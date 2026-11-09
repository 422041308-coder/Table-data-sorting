// S6 前端 · 检索页
// 原生 ES module，无框架、无构建、无 CDN。由 app.js 动态 import，导出 render(host, ctx)。
// ctx = {api, el, toast, state, app}；api 提供 get/post/put/del/download。

function ensureCss() {
  if (document.querySelector('link[data-page-css="common"]')) return;
  const link = document.createElement('link');
  link.rel = 'stylesheet';
  link.href = new URL('../page-common.css', import.meta.url).href;
  link.dataset.pageCss = 'common';
  document.head.appendChild(link);
}

const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
));
const basename = (p) => String(p || '').split(/[\\/]/).pop() || '';
const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);
const fmtBytes = (n) => {
  n = num(n);
  if (!n) return '0 B';
  const u = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(i ? 1 : 0)} ${u[i]}`;
};
const KIND_LABEL = {
  text: '文本', number: '数值', date: '日期', bool: '布尔',
  formula: '公式', blank: '空白', money: '金额', phone: '电话', code: '编码',
};

export async function render(host, ctx) {
  ensureCss();
  const { api, toast, state } = ctx;
  document.querySelectorAll('.pg-modal-mask').forEach((n) => n.remove());

  const st = {
    q: '', mode: 'all', fileId: '', limit: 50, offset: 0,
    total: 0, took: 0, hits: [], files: [], loading: false, error: '', indexed: false,
    edited: new Set(),
  };

  host.innerHTML = `
  <div class="page page-search">
    <section class="panel">
      <div class="page-head">
        <div>
          <h2>全局检索</h2>
          <p class="sub">在目标文件夹的所有工作表中模糊搜索，结果精确到「文件地址 · 工作表 · 第几行 · 第几列 · A1 · 字段名」。</p>
        </div>
        <span class="badge" id="ws-badge">读取工作区…</span>
      </div>
      <div class="toolbar">
        <input id="q" class="pinput grow" type="search" autocomplete="off" spellcheck="false"
               placeholder="输入关键词，按回车搜索；多个关键词用空格分隔，短语用英文双引号">
        <select id="mode" class="pselect" title="匹配模式">
          <option value="all">全部关键词（AND）</option>
          <option value="any">任一关键词（OR）</option>
          <option value="phrase">短语匹配</option>
        </select>
        <select id="file" class="pselect" title="限定文件"><option value="">全部文件</option></select>
        <button id="go" class="btn">搜索</button>
        <button id="replace" class="btn ghost">查找替换</button>
      </div>
      <div class="hint">💡 坐标口径：文件地址 · 工作表 · 第几行 · 第几列（字母+序号）· A1。点击结果行或「编辑」可直接修改并写回原文件（写前自动备份）。</div>
    </section>

    <section class="panel">
      <div class="toolbar between">
        <div id="meta" class="muted">正在读取工作区…</div>
        <div class="toolbar">
          <button id="export" class="btn ghost sm" disabled>导出结果</button>
        </div>
      </div>
      <div id="error" class="error-state" hidden></div>
      <div id="empty" class="empty-state">输入关键词后按回车开始搜索。</div>
      <div id="list" class="tablewrap" hidden>
        <table class="grid grid-search">
          <thead>
            <tr>
              <th style="min-width:240px">文件地址</th>
              <th class="nowrap">工作表</th>
              <th class="num">第几行</th>
              <th class="nowrap">第几列（字母+序号）</th>
              <th class="nowrap">A1</th>
              <th>字段名</th>
              <th style="min-width:220px">命中内容（高亮）</th>
              <th class="nowrap">操作</th>
            </tr>
          </thead>
          <tbody id="rows"></tbody>
        </table>
      </div>
      <div id="pager" class="pager" hidden></div>
    </section>
  </div>`;

  const qEl = host.querySelector('#q');
  const modeEl = host.querySelector('#mode');
  const fileEl = host.querySelector('#file');
  const goEl = host.querySelector('#go');
  const exportEl = host.querySelector('#export');
  const metaEl = host.querySelector('#meta');
  const errEl = host.querySelector('#error');
  const emptyEl = host.querySelector('#empty');
  const listEl = host.querySelector('#list');
  const rowsEl = host.querySelector('#rows');
  const pagerEl = host.querySelector('#pager');
  const wsBadgeEl = host.querySelector('#ws-badge');

  // ------------------------------------------------------------ 工具
  function setError(msg, retry) {
    st.error = msg || '';
    if (!msg) { errEl.hidden = true; errEl.innerHTML = ''; return; }
    errEl.hidden = false;
    errEl.innerHTML = `<b>出错了：</b>${esc(msg)} <button class="linkbtn" id="retry">重试</button>`;
    const btn = errEl.querySelector('#retry');
    if (btn) btn.addEventListener('click', retry || (() => doSearch(st.offset)));
  }

  function rowKey(ref) {
    return [ref.filePath || '', ref.sheet || '', ref.row || 0, ref.col || 0].join('|');
  }

  function setLoading(loading) {
    st.loading = loading;
    goEl.disabled = loading;
    goEl.textContent = loading ? '搜索中…' : '搜索';
    if (loading) {
      setError('');
      emptyEl.hidden = true;
      listEl.hidden = false;
      rowsEl.innerHTML = '<tr><td colspan="8" class="loading">搜索中…</td></tr>';
    }
  }

  function highlightHtml(hit) {
    let text = String(hit.snippet || hit.valueText || '');
    if (!text) return '<span class="muted">（空）</span>';
    let ranges = [];
    const hl = hit.highlight;
    if (Array.isArray(hl)) {
      if (hl.length === 2 && Number.isFinite(Number(hl[0])) && Number.isFinite(Number(hl[1]))) {
        ranges = [[Number(hl[0]), Number(hl[1])]];
      } else if (hl.length && Array.isArray(hl[0])) {
        ranges = hl.filter((x) => Array.isArray(x) && x.length === 2).map((x) => [Number(x[0]), Number(x[1])]);
      }
    }
    const q = st.q.trim();
    if (!ranges.length && q) {
      for (const t of q.split(/\s+/).filter(Boolean).slice(0, 8)) {
        let from = 0;
        while (from < text.length) {
          const at = text.toLowerCase().indexOf(t.toLowerCase(), from);
          if (at < 0) break;
          ranges.push([at, at + t.length]);
          from = at + Math.max(1, t.length);
        }
      }
    }
    ranges = ranges
      .filter(([s, e]) => Number.isFinite(s) && Number.isFinite(e) && e > s && s < text.length)
      .sort((a, b) => a[0] - b[0]);
    let out = '';
    let pos = 0;
    for (const [s, e] of ranges) {
      if (s < pos) continue;
      out += esc(text.slice(pos, s)) + '<mark>' + esc(text.slice(s, e)) + '</mark>';
      pos = e;
    }
    out += esc(text.slice(pos));
    return out;
  }

  function coordText(ref) {
    if (!ref) return '<span class="muted">—</span>';
    return `<span class="coord">第 ${esc(ref.row)} 行 · 第 ${esc(ref.colLetter || '')} 列<span class="muted">（第 ${esc(ref.col)} 列）</span><br>${esc(ref.a1 || '')}</span>`;
  }

  // ------------------------------------------------------------ 工作区 / 文件清单
  async function loadWorkspace() {
    const [ws, files] = await Promise.all([api.get('/workspace'), api.get('/files')]);
    if (!host.isConnected) return;
    const w = ws.data || {};
    const f = files.data || {};
    st.files = Array.isArray(f.files) ? f.files : [];
    state.root = w.root || state.root || '';
    st.indexed = !!w.indexed;
    if (w.indexed) {
      wsBadgeEl.className = 'badge ok';
      wsBadgeEl.textContent = `已索引 ${w.files ?? st.files.length} 文件 / ${w.cells ?? '?'} 单元格${w.indexedAt ? ` · ${w.indexedAt}` : ''}`;
    } else {
      wsBadgeEl.className = 'badge warn';
      wsBadgeEl.textContent = w.root ? '工作区未索引（请到设置页重建索引）' : '未设置工作区';
    }
    wsBadgeEl.title = w.root || '';
    const cur = st.fileId;
    fileEl.innerHTML = '<option value="">全部文件</option>' + st.files.map((x) => (
      `<option value="${esc(x.id)}" title="${esc(x.path)}">${esc(x.name || basename(x.path))}（${esc(x.fmt || '?')}）</option>`
    )).join('');
    if (cur && st.files.some((x) => x.id === cur)) fileEl.value = cur;
    if (!st.files.length) {
      metaEl.textContent = '工作区里还没有可检索的文件，请到「设置」页指定文件夹并重建索引。';
    }
  }

  // ------------------------------------------------------------ 搜索
  async function doSearch(offset = 0) {
    if (st.loading) return;
    st.q = qEl.value;
    st.mode = modeEl.value;
    st.fileId = fileEl.value;
    st.offset = Math.max(0, offset);
    setLoading(true);
    setError('');
    try {
      const res = await api.get('/search', {
        q: st.q, mode: st.mode, file: st.fileId, limit: st.limit, offset: st.offset,
      });
      if (!host.isConnected) return;
      const d = res.data || {};
      st.hits = Array.isArray(d.hits) ? d.hits : [];
      st.total = num(d.total);
      st.took = num(d.took);
      renderResults();
    } catch (e) {
      if (!host.isConnected) return;
      st.hits = [];
      st.total = 0;
      listEl.hidden = true;
      pagerEl.hidden = true;
      exportEl.disabled = true;
      emptyEl.hidden = true;
      metaEl.textContent = '搜索失败';
      setError(e.message || String(e));
    } finally {
      if (host.isConnected) setLoading(false);
    }
  }

  function renderResults() {
    const searched = st.total > 0 || st.hits.length > 0;
    const took = st.took ? `，用时 ${st.took.toFixed(3)} s` : '';
    const scope = st.fileId ? `（限定文件：${basename((st.files.find((x) => x.id === st.fileId) || {}).path || '')}）` : '';
    metaEl.textContent = st.q
      ? `关键词「${st.q}」命中 ${st.total} 条${scope}${took}`
      : `未输入关键词，显示全部命中 ${st.total} 条${scope}${took}`;
    exportEl.disabled = st.total === 0;

    if (!searched) {
      listEl.hidden = true;
      pagerEl.hidden = true;
      emptyEl.hidden = false;
      emptyEl.innerHTML = st.q
        ? `<b>没有找到「${esc(st.q)}」</b>试试更短的关键词、换用同义词，或切换为「任一关键词」。`
        : `<b>还没有结果</b>在工作区里输入关键词，按回车开始检索。`;
      return;
    }

    emptyEl.hidden = true;
    listEl.hidden = false;
    pagerEl.hidden = st.total <= st.limit;

    rowsEl.innerHTML = st.hits.map((h, i) => {
      const ref = h.ref || {};
      const file = ref.filePath || '';
      const edited = st.edited.has(rowKey(ref));
      return `
      <tr data-i="${i}" class="${edited ? 'row-edited' : ''}" title="点击行编辑该单元格">
        <td>
          <div class="filecell" title="${esc(file)}">
            <span class="fname">${esc(basename(file) || '（未知文件）')}</span>
            <span class="pathsmall mono">${esc(file)}</span>
          </div>
          <div class="rowbtns">
            <button class="linkbtn" data-act="folder" data-i="${i}" title="在资源管理器中定位该文件">打开所在文件夹</button>
            <button class="linkbtn" data-act="open" data-i="${i}" title="用 WPS / Excel 打开该文件">用 WPS 打开</button>
            <button class="linkbtn copybtn" data-act="copy" data-i="${i}" title="复制完整路径">复制路径</button>
          </div>
        </td>
        <td>${esc(ref.sheet || '')}</td>
        <td class="num coordcell">第${esc(ref.row ?? '')}行</td>
        <td class="nowrap coordcell">${esc(ref.colLetter || '')}列(${esc(ref.col ?? '')})</td>
        <td class="nowrap coordcell" title="${esc(ref.display || '')}">${esc(ref.a1 || '')}</td>
        <td>${h.fieldName ? esc(h.fieldName) : '<span class="muted">（未识别字段）</span>'}</td>
        <td class="hitv" title="${esc(h.valueText || '')}">${highlightHtml(h)}</td>
        <td class="nowrap"><button class="btn sm ghost" data-act="edit" data-i="${i}">编辑</button></td>
      </tr>`;
    }).join('');

    renderPager();
  }

  function renderPager() {
    const totalPages = Math.max(1, Math.ceil(st.total / st.limit));
    const cur = Math.min(totalPages, Math.floor(st.offset / st.limit) + 1);
    const pages = [];
    if (totalPages <= 7) {
      for (let i = 1; i <= totalPages; i++) pages.push(i);
    } else {
      const set = new Set([1, totalPages, cur - 1, cur, cur + 1].filter((p) => p >= 1 && p <= totalPages));
      const sorted = [...set].sort((a, b) => a - b);
      for (let i = 0; i < sorted.length; i++) {
        if (i && sorted[i] - sorted[i - 1] > 1) pages.push('…');
        pages.push(sorted[i]);
      }
    }
    pagerEl.innerHTML = `
      <span class="pginfo">第 ${cur} / ${totalPages} 页 · 共 ${st.total} 条 · 每页 ${st.limit} 条</span>
      <button class="pgbtn" data-page="${cur - 1}" ${cur <= 1 ? 'disabled' : ''} title="上一页">‹</button>
      ${pages.map((p) => (p === '…'
        ? '<span class="pgbtn gap">…</span>'
        : `<button class="pgbtn ${p === cur ? 'active' : ''}" data-page="${p}">${p}</button>`)).join('')}
      <button class="pgbtn" data-page="${cur + 1}" ${cur >= totalPages ? 'disabled' : ''} title="下一页">›</button>
      <input class="pinput" id="page-jump" type="number" min="1" max="${totalPages}" placeholder="页码" style="width:74px;min-height:28px;padding:3px 8px">
      <button class="pgbtn" id="page-go">跳转</button>`;
  }

  async function copyPath(path) {
    if (!path) { toast('缺少文件路径', 'err'); return; }
    try {
      await navigator.clipboard.writeText(path);
      toast('文件完整路径已复制到剪贴板', 'ok');
    } catch {
      toast('复制失败，请手动选择路径', 'err');
    }
  }

  async function openPath(path, mode) {
    if (!path) { toast('缺少文件路径', 'err'); return; }
    try {
      await api.post('/open', { path, mode });
      toast(mode === 'folder' ? '已在资源管理器中定位文件' : '已调用系统默认程序打开该文件', 'ok');
    } catch (e) {
      let copied = false;
      try { await navigator.clipboard.writeText(path); copied = true; } catch { /* 忽略 */ }
      toast(`打开失败：${e.message || e}${copied ? '；文件路径已复制到剪贴板' : ''}`, 'err');
    }
  }

  // ------------------------------------------------------------ 单元格编辑弹窗
  function createModal(title, wide = false) {
    document.querySelectorAll('.pg-modal-mask').forEach((n) => n.remove());
    const mask = document.createElement('div');
    mask.className = 'pg-modal-mask';
    mask.innerHTML = `
      <div class="pg-modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
        <div class="modal-head"><h3>${esc(title)}</h3><button class="modal-x" data-close title="关闭">×</button></div>
        <div class="modal-body"></div>
        <div class="modal-foot"></div>
      </div>`;
    document.body.appendChild(mask);
    const onKey = (e) => { if (e.key === 'Escape') close(); };
    function close() {
      mask.remove();
      document.removeEventListener('keydown', onKey);
    }
    document.addEventListener('keydown', onKey);
    mask.addEventListener('click', (e) => { if (e.target === mask) close(); });
    mask.querySelector('[data-close]').addEventListener('click', close);
    return {
      mask,
      body: mask.querySelector('.modal-body'),
      foot: mask.querySelector('.modal-foot'),
      close,
    };
  }

  function fieldRow(label, value) {
    return `<div><span>${esc(label)}</span><b>${value === '' || value == null ? '<span class="muted">—</span>' : esc(value)}</b></div>`;
  }

  async function openEdit(i) {
    const hit = st.hits[i];
    if (!hit || !hit.ref) return;
    const ref0 = hit.ref;
    const modal = createModal('单元格详情 / 编辑');
    modal.body.innerHTML = '<div class="loading">正在读取单元格…</div>';
    let d;
    try {
      const res = await api.get('/cell', {
        file: ref0.filePath, sheet: ref0.sheet, row: ref0.row, col: ref0.col,
      });
      d = res.data || {};
    } catch (e) {
      if (!modal.mask.isConnected) return;
      modal.body.innerHTML = `<div class="error-state">读取单元格失败：${esc(e.message || e)}</div>`;
      modal.foot.innerHTML = '<button class="btn ghost" data-close2>关闭</button>';
      modal.foot.querySelector('[data-close2]').addEventListener('click', modal.close);
      return;
    }
    if (!modal.mask.isConnected) return;

    const ref = d.ref || ref0;
    const cur = d.current || {};
    const value = d.value == null ? '' : String(d.value);
    const formula = d.formula == null ? '' : String(d.formula);
    const editable = d.editable !== false;
    modal.body.innerHTML = `
      <div class="coord-line mono">${esc(ref.display || `${ref.a1 || ''} · 第${ref.row}行 · 第${ref.col}列`)}</div>
      <div class="kvline"><span class="k">文件</span><span class="mono wrapdot">${esc(ref.filePath || '')}</span>
        <button class="linkbtn" id="m-folder">打开所在文件夹</button>
        <button class="linkbtn" id="m-open">用 WPS 打开</button>
      </div>
      <div class="kv-grid">
        ${fieldRow('工作表', ref.sheet)}
        ${fieldRow('第几行', ref.row)}
        ${fieldRow('第几列', ref.colLetter ? `${ref.colLetter}（第 ${ref.col} 列）` : ref.col)}
        ${fieldRow('A1 引用', ref.a1)}
        ${fieldRow('字段名', d.fieldName)}
        ${fieldRow('单元格类型', KIND_LABEL[d.kind] || d.kind)}
        ${fieldRow('公式', formula)}
        ${fieldRow('是否小计行', d.isSubtotal ? '是（小计/占位行）' : '否')}
        ${fieldRow('合并归属', d.mergeOwner ? (typeof d.mergeOwner === 'string' ? d.mergeOwner : JSON.stringify(d.mergeOwner)) : '')}
      </div>
      ${d.isSubtotal ? '<div class="warn-state">该行被识别为小计或分类占位行，修改前请确认不会破坏模板公式。</div>' : ''}
      ${editable ? '' : '<div class="warn-state">该文件格式不支持原地写回（如 .xls），请另存为 .xlsx 后再编辑。</div>'}
      <div class="formline"><label>修改类型</label>
        <select id="m-kind" class="pselect">
          <option value="value">修改显示值（value）</option>
          <option value="formula" ${formula ? '' : 'disabled'}>修改公式（formula）${formula ? '' : '（无公式）'}</option>
        </select>
      </div>
      <div class="formline"><label>当前值</label><div class="readonly mono">${esc(value) || '<span class="muted">（空）</span>'}</div></div>
      <div class="formline"><label>新值</label><textarea id="m-value" class="ptextarea mono" rows="3" placeholder="输入新的单元格内容"></textarea></div>
      <div class="hint">保存前会自动备份原文件；若文件已被 WPS/Excel 等外部程序修改，保存会被拒绝并提示冲突。修改会记入审计日志。</div>
      <div id="m-err" class="error-state" hidden></div>`;
    const kindEl = modal.body.querySelector('#m-kind');
    const valEl = modal.body.querySelector('#m-value');
    const errBox = modal.body.querySelector('#m-err');
    const btnFolder = modal.body.querySelector('#m-folder');
    const btnOpen = modal.body.querySelector('#m-open');
    if (btnFolder) btnFolder.addEventListener('click', () => openPath(ref.filePath, 'folder'));
    if (btnOpen) btnOpen.addEventListener('click', () => openPath(ref.filePath, 'file'));

    let filledFor = '';
    function fillForKind() {
      const k = kindEl.value;
      if (filledFor === k) return;
      filledFor = k;
      valEl.value = k === 'formula' ? formula : value;
    }
    fillForKind();
    kindEl.addEventListener('change', fillForKind);
    valEl.focus();

    function showErr(msg) {
      errBox.hidden = !msg;
      errBox.innerHTML = msg ? `<b>保存失败：</b>${esc(msg)}` : '';
    }

    modal.foot.innerHTML = `
      <button class="btn ghost" data-cancel>取消</button>
      <button class="btn" id="m-save" ${editable ? '' : 'disabled'}>保存并写回</button>`;
    modal.foot.querySelector('[data-cancel]').addEventListener('click', modal.close);
    const saveBtn = modal.foot.querySelector('#m-save');
    saveBtn.addEventListener('click', async () => {
      const kind = kindEl.value;
      const newValue = valEl.value;
      if (kind === 'value' && newValue === '') {
        if (!confirm('新值为空，确定要清空该单元格吗？')) return;
      }
      if (kind === 'formula' && !newValue.trim().startsWith('=')) {
        showErr('公式必须以 = 开头，例如 =SUM(A1:A10)。');
        return;
      }
      saveBtn.disabled = true;
      const oldText = saveBtn.textContent;
      saveBtn.textContent = '保存中…';
      showErr('');
      try {
        const res = await api.put('/cell', {
          ref: { filePath: ref.filePath, sheet: ref.sheet, row: ref.row, col: ref.col },
          newValue,
          kind,
          expectMtime: num(cur.mtime),
          expectSize: num(cur.size),
        });
        const out = res.data || {};
        if (out.conflict) throw Object.assign(new Error('文件已被外部修改，请刷新后重试'), { code: 'CONFLICT' });
        toast(`已保存 ${ref.a1 || ''}，原文件已自动备份`, 'ok');
        st.edited.add(rowKey(ref0));
        modal.close();
        await doSearch(st.offset);
      } catch (e) {
        saveBtn.disabled = false;
        saveBtn.textContent = oldText;
        const msg = e.code === 'CONFLICT'
          ? '文件已被 WPS/Excel 等外部程序修改，本次保存被拒绝。请关闭该文件后刷新结果再试。'
          : (e.message || String(e));
        showErr(msg);
        toast(msg, 'err');
      }
    });
  }

  // ------------------------------------------------------------ 查找替换弹窗
  function openReplace() {
    const modal = createModal('查找替换（预演后批量写回）', true);
    const scope = st.fileId
      ? `当前限定文件：${basename((st.files.find((x) => x.id === st.fileId) || {}).path || st.fileId)}`
      : '范围：工作区全部文件';
    modal.body.innerHTML = `
      <div class="inline-form">
        <input id="r-find" class="pinput grow" placeholder="查找内容（必填）" autocomplete="off">
        <input id="r-replace" class="pinput grow" placeholder="替换为（可为空 = 删除）" autocomplete="off">
        <label class="checkline"><input id="r-case" type="checkbox">区分大小写</label>
        <button class="btn" id="r-preview">预演</button>
      </div>
      <div class="hint">${esc(scope)}。先预演生成「旧值 → 新值」清单，逐条勾选确认后才会写回；写回前同样自动备份。</div>
      <div id="r-err" class="error-state" hidden></div>
      <div id="r-result"></div>`;
    const findEl = modal.body.querySelector('#r-find');
    const replaceEl = modal.body.querySelector('#r-replace');
    const caseEl = modal.body.querySelector('#r-case');
    const errEl = modal.body.querySelector('#r-err');
    const resultEl = modal.body.querySelector('#r-result');
    const previewBtn = modal.body.querySelector('#r-preview');
    modal.foot.innerHTML = '<button class="btn ghost" data-close3>关闭</button>';
    modal.foot.querySelector('[data-close3]').addEventListener('click', modal.close);
    findEl.focus();

    let matches = [];

    function showErr(msg) {
      errEl.hidden = !msg;
      errEl.innerHTML = msg ? `<b>操作失败：</b>${esc(msg)}` : '';
    }

    async function doPreview() {
      const find = findEl.value;
      if (!find) { showErr('请先填写查找内容。'); return; }
      previewBtn.disabled = true;
      previewBtn.textContent = '预演中…';
      showErr('');
      resultEl.innerHTML = '<div class="loading">正在生成替换预演…</div>';
      try {
        // body 形状按 CONTRACT 5.1 冻结：{find, replace, filePath|null, fileId|null, matchCase, limit}
        const scoped = st.files.find((f) => String(f.id) === String(st.fileId)) || null;
        const res = await api.post('/replace/preview', {
          find,
          replace: replaceEl.value,
          filePath: scoped ? scoped.path : null,
          fileId: st.fileId || null,
          matchCase: !!caseEl.checked,
          limit: 500,
        });
        const d = res.data || {};
        matches = Array.isArray(d.matches) ? d.matches
          : Array.isArray(d.items) ? d.items
            : Array.isArray(d) ? d : [];
        renderPreview(d);
      } catch (e) {
        resultEl.innerHTML = '';
        showErr(e.message || String(e));
      } finally {
        previewBtn.disabled = false;
        previewBtn.textContent = '预演';
      }
    }

    function renderPreview(d) {
      if (!matches.length) {
        resultEl.innerHTML = '<div class="empty-state">没有找到可替换的单元格。</div>';
        return;
      }
      const count = num(d.count) || matches.length;
      resultEl.innerHTML = `
        <div class="toolbar between" style="margin-top:12px">
          <div class="muted">共 ${count} 处可替换${d.truncated ? '（已截断，仅显示前 500 条）' : ''}</div>
          <div class="toolbar">
            <label class="checkline"><input type="checkbox" id="r-all" checked>全选</label>
            <button class="btn sm" id="r-apply">替换选中项</button>
          </div>
        </div>
        <div class="tablewrap" style="max-height:42vh;margin-top:8px">
          <table class="grid">
            <thead><tr><th style="width:38px">选</th><th>坐标（文件 · 工作表 · 行 · 列 · A1）</th><th>旧值</th><th>新值</th></tr></thead>
            <tbody>${matches.map((m, i) => {
              const ref = m.ref || {};
              return `<tr>
                <td><input type="checkbox" class="r-pick" data-i="${i}" checked></td>
                <td><div class="mini">${esc(basename(ref.filePath || ''))}</div>
                    <div class="mono wrapdot">${esc(ref.display || '')}</div></td>
                <td class="hitv">${esc(m.oldValue ?? '')}</td>
                <td class="hitv">${esc(m.newValue ?? '')}</td>
              </tr>`;
            }).join('')}</tbody>
          </table>
        </div>`;
      const allEl = resultEl.querySelector('#r-all');
      const applyBtn = resultEl.querySelector('#r-apply');
      const picks = () => Array.from(resultEl.querySelectorAll('.r-pick'));
      allEl.addEventListener('change', () => picks().forEach((x) => { x.checked = allEl.checked; }));
      picks().forEach((x) => x.addEventListener('change', () => {
        allEl.checked = picks().every((y) => y.checked);
      }));
      applyBtn.addEventListener('click', async () => {
        const selected = picks().filter((x) => x.checked).map((x) => matches[num(x.dataset.i)]).filter(Boolean);
        if (!selected.length) { showErr('请至少勾选一条要替换的记录。'); return; }
        if (!confirm(`确认替换 ${selected.length} 处？写回前会自动备份原文件。`)) return;
        applyBtn.disabled = true;
        applyBtn.textContent = '写回中…';
        showErr('');
        try {
          // 按 CONTRACT 5.1：apply 只提交用户确认过的条目，逐条带 newValue
          const res = await api.post('/replace/apply', {
            items: selected.map((m) => ({ ref: m.ref, newValue: m.newValue })),
          });
          const out = res.data || {};
          toast(`查找替换完成：写回 ${out.written ?? selected.length} 处`, 'ok');
          modal.close();
          await doSearch(st.offset);
        } catch (e) {
          applyBtn.disabled = false;
          applyBtn.textContent = '替换选中项';
          showErr(e.message || String(e));
        }
      });
    }

    previewBtn.addEventListener('click', doPreview);
    findEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') doPreview(); });
  }

  // ------------------------------------------------------------ 事件绑定
  goEl.addEventListener('click', () => doSearch(0));
  qEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') doSearch(0); });
  modeEl.addEventListener('change', () => doSearch(0));
  fileEl.addEventListener('change', () => doSearch(0));

  rowsEl.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-act]');
    if (btn) {
      const i = num(btn.dataset.i);
      const hit = st.hits[i];
      if (!hit || !hit.ref) return;
      if (btn.dataset.act === 'edit') openEdit(i);
      else if (btn.dataset.act === 'folder') openPath(hit.ref.filePath, 'folder');
      else if (btn.dataset.act === 'open') openPath(hit.ref.filePath, 'file');
      else if (btn.dataset.act === 'copy') copyPath(hit.ref.filePath);
      return;
    }
    const tr = e.target.closest('tr[data-i]');
    if (tr) openEdit(num(tr.dataset.i));
  });

  pagerEl.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-page]');
    if (btn && !btn.disabled) {
      doSearch(num(btn.dataset.page) * st.limit - st.limit);
      return;
    }
    if (e.target.id === 'page-go') {
      const inp = pagerEl.querySelector('#page-jump');
      const p = Math.max(1, num(inp && inp.value) || 1);
      doSearch((p - 1) * st.limit);
    }
  });
  pagerEl.addEventListener('keydown', (e) => {
    if (e.target.id === 'page-jump' && e.key === 'Enter') {
      const p = Math.max(1, num(e.target.value) || 1);
      doSearch((p - 1) * st.limit);
    }
  });

  exportEl.addEventListener('click', async () => {
    exportEl.disabled = true;
    exportEl.textContent = '导出中…';
    try {
      const res = await api.post('/search/export', {
        q: st.q, mode: st.mode, fileId: st.fileId || null, limit: 5000,
      });
      const d = res.data || {};
      if (d.token) {
        api.download(d.token);
        toast(`已生成导出文件 ${d.name || ''}（${d.count ?? st.total} 条）`, 'ok');
      } else {
        toast(`导出已生成：${d.file || d.path || '（未返回路径）'}`, 'ok');
      }
    } catch (e) {
      toast(`导出失败：${e.message || e}`, 'err');
    } finally {
      if (host.isConnected) {
        exportEl.disabled = st.total === 0;
        exportEl.textContent = '导出结果';
      }
    }
  });

  host.querySelector('#replace').addEventListener('click', openReplace);

  // 支持深链：#search?q=合同&mode=all（便于分享/刷新后保留查询，也用于截图）
  const initial = (ctx.state && ctx.state.params) || {};
  if (initial.q) {
    qEl.value = initial.q;
    st.q = initial.q;
  }
  if (initial.mode && ['all', 'any', 'phrase'].includes(initial.mode)) {
    modeEl.value = initial.mode;
    st.mode = initial.mode;
  }

  // ------------------------------------------------------------ 初始化
  try {
    await loadWorkspace();
  } catch (e) {
    if (!host.isConnected) return;
    wsBadgeEl.className = 'badge err';
    wsBadgeEl.textContent = '服务未连接';
    metaEl.textContent = '无法读取工作区状态，请确认服务已启动。';
    setError(e.message || String(e), () => loadWorkspace());
    return;
  }
  if (st.indexed) {
    // 已建索引：直接把全部命中列出来，减少一次额外点击
    await doSearch(0);
  } else {
    // 未建索引时不自动搜索，避免打开页面就触发几十秒的首次建索引
    metaEl.textContent = st.files.length
      ? '工作区尚未索引：输入关键词后按回车搜索（首次搜索会建立索引，可能需要几十秒）。'
      : '工作区里还没有可检索的文件，请到「设置」页指定文件夹并重建索引。';
    emptyEl.innerHTML = '<b>输入关键词开始搜索</b>支持多关键词（空格分隔）与短语（英文双引号）；结果会给出文件、工作表、行、列与 A1 坐标。';
  }
}
