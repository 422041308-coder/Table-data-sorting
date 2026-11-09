// S6 前端 · 数据对比页
// 原生 ES module，无框架、无构建、无 CDN。render(host, ctx) 由 app.js 动态调用。
// 流程：左右文件/工作表 + 主键列（1 基列号，可多选）+ 选项 → POST /api/compare
//       → 汇总卡片 + 差异表（左右两侧文件/行/列/A1）→ 导出下载。

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
const num = (v) => (Number.isFinite(Number(v)) ? Number(v) : 0);
const basename = (p) => String(p || '').split(/[\\/]/).pop() || '';
const clone = (v) => JSON.parse(JSON.stringify(v));

const KIND_META = {
  added: { label: '新增（仅右）', cls: 'added' },
  removed: { label: '删除（仅左）', cls: 'removed' },
  changed: { label: '修改', cls: 'changed' },
};

function letterIndex(letter) {
  let n = 0;
  for (const ch of String(letter || '').trim().toUpperCase()) {
    const c = ch.charCodeAt(0);
    if (c < 65 || c > 90) continue;
    n = n * 26 + (c - 64);
  }
  return n;
}
function colLetter(col) {
  let n = Number(col);
  if (!Number.isFinite(n) || n < 1) return '';
  let out = '';
  while (n) { n--; out = String.fromCharCode(65 + (n % 26)) + out; n = Math.floor(n / 26); }
  return out;
}

export async function render(host, ctx) {
  ensureCss();
  const { api, toast } = ctx;
  document.querySelectorAll('.pg-modal-mask').forEach((n) => n.remove());

  const st = {
    files: [],
    left: { fileId: '', sheetIndex: 0, sheets: [] },
    right: { fileId: '', sheetIndex: 0, sheets: [] },
    keys: [],
    candidates: [],
    filter: 'all',
    offset: 0,
    result: null,
    loading: false,
  };

  host.innerHTML = `
  <div class="page page-compare">
    <section class="panel">
      <div class="page-head">
        <div>
          <h2>数据对比</h2>
          <p class="sub">选两张表（可不同文件/工作表），按主键匹配条目，逐字段输出「新增 / 删除 / 修改」，每条差异都带左右两侧的文件、工作表、行、列与 A1。</p>
        </div>
        <button class="btn ghost sm" id="refresh">刷新文件清单</button>
      </div>
      <div class="grid2">
        <div class="fields-box" style="max-height:none;background:#fff;padding:10px">
          <div class="subhead" style="margin-top:0"><h3>左侧表（基准）</h3><span class="mini" id="left-state"></span></div>
          <div class="toolbar">
            <select id="left-file" class="pselect grow"></select>
            <select id="left-sheet" class="pselect" style="min-width:170px"></select>
          </div>
          <div class="mini mono wrapdot" id="left-path" style="margin-top:6px"></div>
        </div>
        <div class="fields-box" style="max-height:none;background:#fff;padding:10px">
          <div class="subhead" style="margin-top:0"><h3>右侧表（对照）</h3><span class="mini" id="right-state"></span></div>
          <div class="toolbar">
            <select id="right-file" class="pselect grow"></select>
            <select id="right-sheet" class="pselect" style="min-width:170px"></select>
          </div>
          <div class="mini mono wrapdot" id="right-path" style="margin-top:6px"></div>
        </div>
      </div>

      <div class="subhead"><h3>主键列（1 基列号，可多选）</h3>
        <span class="mini">主键用于匹配两张表的同一行；组合键可多选</span>
      </div>
      <div class="toolbar">
        <input id="keys-input" class="pinput grow" placeholder="输入列字母或列号，逗号/空格分隔，如：AF 或 32,33">
        <button id="keys-parse" class="btn ghost">解析并应用</button>
        <button id="keys-clear" class="btn ghost">清空</button>
      </div>
      <div class="chips" id="keys-chips"><span class="mini">尚未选择主键</span></div>
      <div id="key-candidates-wrap" hidden>
        <div class="mini" style="margin:4px 0">从表头识别的主键候选（勾选即加入；左表列名 / 右表列名）：</div>
        <div class="keygrid" id="key-candidates"></div>
      </div>

      <div class="subhead"><h3>对比选项</h3></div>
      <div class="toolbar">
        <label class="checkline"><input type="checkbox" id="opt-trim" checked>忽略文本首尾空白（trim）</label>
        <label class="checkline"><input type="checkbox" id="opt-filldown" checked>合并单元格/空值向下填充（filldown）</label>
        <label class="checkline">数值容差<input id="opt-tol" class="pinput" type="number" step="0.001" min="0" value="0.001" style="width:100px;min-height:28px;padding:3px 8px"></label>
        <span class="spring"></span>
        <button id="compare" class="btn">开始对比</button>
        <button id="export" class="btn ghost" disabled>导出对比报告</button>
      </div>
      <div class="hint">差异口径：<b>added</b>=仅右表有，<b>removed</b>=仅左表有，<b>changed</b>=两边都有但字段值不同。主键重复/为空会写入告警，不会静默错配。</div>
      <div id="error" class="error-state" hidden></div>
    </section>

    <section class="panel">
      <div class="subhead"><h3>对比结果</h3><span class="mini" id="result-state"></span></div>
      <div id="warnings"></div>
      <div id="summary"></div>
      <div id="empty" class="empty-state">选择左右两张表并指定主键列后，点「开始对比」。</div>
      <div id="list" class="tablewrap" hidden>
        <table class="grid diff-table">
          <thead><tr>
            <th class="nowrap">差异类型</th>
            <th class="nowrap">主键值</th>
            <th>字段名</th>
            <th>左值</th>
            <th style="min-width:160px">左文件</th>
            <th class="nowrap coordcell">左 行 / 列 / A1</th>
            <th>右值</th>
            <th style="min-width:160px">右文件</th>
            <th class="nowrap coordcell">右 行 / 列 / A1</th>
          </tr></thead>
          <tbody id="rows"></tbody>
        </table>
      </div>
      <div id="pager" class="pager" hidden></div>
      <div class="toolbar between" id="filter-bar" hidden style="margin-top:10px">
        <div class="toolbar">
          <span class="mini">筛选差异类型</span>
          <select id="filter" class="pselect">
            <option value="all">全部</option>
            <option value="added">仅新增</option>
            <option value="removed">仅删除</option>
            <option value="changed">仅修改</option>
          </select>
        </div>
        <span class="mini" id="filter-count"></span>
      </div>
    </section>
  </div>`;

  const leftFileEl = host.querySelector('#left-file');
  const rightFileEl = host.querySelector('#right-file');
  const leftSheetEl = host.querySelector('#left-sheet');
  const rightSheetEl = host.querySelector('#right-sheet');
  const leftStateEl = host.querySelector('#left-state');
  const rightStateEl = host.querySelector('#right-state');
  const leftPathEl = host.querySelector('#left-path');
  const rightPathEl = host.querySelector('#right-path');
  const keysInputEl = host.querySelector('#keys-input');
  const keysChipsEl = host.querySelector('#keys-chips');
  const keyWrapEl = host.querySelector('#key-candidates-wrap');
  const keyCandidatesEl = host.querySelector('#key-candidates');
  const errEl = host.querySelector('#error');
  const summaryEl = host.querySelector('#summary');
  const warningsEl = host.querySelector('#warnings');
  const emptyEl = host.querySelector('#empty');
  const listEl = host.querySelector('#list');
  const rowsEl = host.querySelector('#rows');
  const pagerEl = host.querySelector('#pager');
  const resultStateEl = host.querySelector('#result-state');
  const filterEl = host.querySelector('#filter');
  const filterBarEl = host.querySelector('#filter-bar');
  const filterCountEl = host.querySelector('#filter-count');
  const exportEl = host.querySelector('#export');
  const compareBtn = host.querySelector('#compare');

  function setError(msg) {
    errEl.hidden = !msg;
    errEl.innerHTML = msg ? `<b>出错了：</b>${esc(msg)}` : '';
  }

  const sideFile = (side) => st.files.find((f) => String(f.id) === String(side.fileId));
  const sideSheet = (side) => side.sheets[side.sheetIndex];
  const sidePath = (side) => (sideFile(side) || {}).path || '';
  const sideSheetName = (side) => (sideSheet(side) || {}).name || '';

  function options() {
    return {
      trim: host.querySelector('#opt-trim').checked,
      filldown: host.querySelector('#opt-filldown').checked,
      tolerance: Math.max(0, num(host.querySelector('#opt-tol').value)),
    };
  }

  // ------------------------------------------------------------ 文件/工作表
  async function loadFiles() {
    const res = await api.get('/files');
    if (!host.isConnected) return;
    const f = res.data || {};
    st.files = Array.isArray(f.files) ? f.files : [];
    const optsHtml = st.files.length
      ? st.files.map((x) => `<option value="${esc(x.id)}" title="${esc(x.path)}">${esc(x.name || basename(x.path))}（${esc(x.fmt || '?')}）</option>`).join('')
      : '<option value="">（工作区没有文件）</option>';
    leftFileEl.innerHTML = optsHtml;
    rightFileEl.innerHTML = optsHtml;
    if (st.files.length) {
      if (!st.left.fileId) st.left.fileId = st.files[0].id;
      if (!st.right.fileId) st.right.fileId = st.files[Math.min(1, st.files.length - 1)].id;
      leftFileEl.value = st.left.fileId;
      rightFileEl.value = st.right.fileId;
    }
    await Promise.all([loadSheets('left'), loadSheets('right')]);
    refreshSideText();
  }

  async function loadSheets(sideKey) {
    const side = st[sideKey];
    side.sheets = [];
    side.sheetIndex = 0;
    const sheetEl = sideKey === 'left' ? leftSheetEl : rightSheetEl;
    if (!side.fileId) {
      sheetEl.innerHTML = '<option value="">（未选择文件）</option>';
      return;
    }
    sheetEl.innerHTML = '<option value="">读取中…</option>';
    try {
      const res = await api.get(`/files/${encodeURIComponent(side.fileId)}/sheets`);
      if (!host.isConnected) return;
      const d = res.data || {};
      side.sheets = Array.isArray(d.sheets) ? d.sheets : [];
      sheetEl.innerHTML = side.sheets.length
        ? side.sheets.map((s, i) => (
          `<option value="${i}" ${i === 0 ? 'selected' : ''}>${esc(s.name || ('Sheet' + (i + 1)))}（${num(s.nRows)}×${num(s.nCols)}）</option>`
        )).join('')
        : '<option value="">（该文件没有工作表）</option>';
      side.sheetIndex = 0;
    } catch (e) {
      sheetEl.innerHTML = '<option value="">读取失败</option>';
      toast(`工作表读取失败：${e.message || e}`, 'err');
    }
  }

  function refreshSideText() {
    leftStateEl.textContent = sideSheet(st.left) ? `${num(sideSheet(st.left).nRows)} 行` : '';
    rightStateEl.textContent = sideSheet(st.right) ? `${num(sideSheet(st.right).nRows)} 行` : '';
    leftPathEl.textContent = sidePath(st.left);
    rightPathEl.textContent = sidePath(st.right);
  }

  // ------------------------------------------------------------ 主键列
  function normalizeKeys(list) {
    const out = [];
    for (const raw of list) {
      const token = String(raw || '').trim();
      if (!token) continue;
      let n = 0;
      if (/^\d+$/.test(token)) n = Number(token);
      else if (/^[A-Za-z]{1,3}$/.test(token)) n = letterIndex(token);
      if (n >= 1 && n <= 16384 && !out.includes(n)) out.push(n);
    }
    return out.sort((a, b) => a - b);
  }

  function syncKeyUI() {
    keysInputEl.value = st.keys.join(', ');
    if (!st.keys.length) {
      keysChipsEl.innerHTML = '<span class="mini">尚未选择主键</span>';
    } else {
      keysChipsEl.innerHTML = st.keys.map((n) => (
        `<span class="chip">${colLetter(n)}（第 ${n} 列）<button data-key="${n}" title="移除">×</button></span>`
      )).join('');
    }
    keyCandidatesEl.querySelectorAll('input[type=checkbox]').forEach((cb) => {
      cb.checked = st.keys.includes(num(cb.value));
    });
  }

  async function loadCandidates() {
    st.candidates = [];
    if (!sideFile(st.left) || !sideFile(st.right)) {
      keyWrapEl.hidden = true;
      return;
    }
    keyWrapEl.hidden = true;
    try {
      const [lp, rp] = await Promise.all([
        api.get(`/files/${encodeURIComponent(st.left.fileId)}/sheet/${st.left.sheetIndex}/preview`, { rows: 6 }),
        api.get(`/files/${encodeURIComponent(st.right.fileId)}/sheet/${st.right.sheetIndex}/preview`, { rows: 6 }),
      ]);
      if (!host.isConnected) return;
      const lcols = deriveHeaderCols(lp.data || {}, sideSheet(st.left) || {});
      const rcols = deriveHeaderCols(rp.data || {}, sideSheet(st.right) || {});
      const map = new Map();
      for (const c of lcols) map.set(c.col, { col: c.col, left: c.name, right: '' });
      for (const c of rcols) {
        if (map.has(c.col)) map.get(c.col).right = c.name;
        else map.set(c.col, { col: c.col, left: '', right: c.name });
      }
      st.candidates = [...map.values()].sort((a, b) => a.col - b.col).slice(0, 60);
    } catch (e) {
      st.candidates = [];
    }
    if (!st.candidates.length) {
      keyWrapEl.hidden = true;
      return;
    }
    keyWrapEl.hidden = false;
    keyCandidatesEl.innerHTML = st.candidates.map((c) => `
      <label class="checkline" title="第 ${c.col} 列 · ${esc(c.left || '')}${c.right && c.right !== c.left ? ` / ${esc(c.right)}` : ''}">
        <input type="checkbox" value="${c.col}">
        <span>${esc(c.left || c.right || `第 ${c.col} 列`)} <span class="mini">（第 ${c.col} 列）</span></span>
      </label>`).join('');
    syncKeyUI();
  }

  function deriveHeaderCols(preview, sheet) {
    const rows = Array.isArray(preview.rows) ? preview.rows : [];
    const dataStart = num(preview.dataStart) || num(sheet.dataStart) || 1;
    let candidates = rows.filter((r) => num(r.r) < dataStart);
    if (!candidates.length) candidates = rows.slice(0, 2);
    candidates = candidates.filter((r) => r.cells && Object.keys(r.cells).length);
    if (!candidates.length) return [];
    const counts = candidates.map((r) => Object.keys(r.cells).length);
    const maxCount = Math.max(...counts);
    const chosen = [];
    for (const row of candidates) {
      const vals = Object.entries(row.cells || {}).filter(([, v]) => v != null && String(v).trim() !== '');
      if (!vals.length) continue;
      if (vals.length < Math.max(2, Math.floor(maxCount * 0.3))) continue;
      if (vals.some(([, v]) => typeof v === 'number' || /^-?\d+([.,]\d+)?$/.test(String(v).trim()))) break;
      chosen.push(row);
    }
    const keys = [...new Set(chosen.flatMap((r) => Object.keys(r.cells || {})))]
      .filter((k) => /^[A-Za-z]+$/.test(k))
      .sort((a, b) => letterIndex(a) - letterIndex(b));
    return keys.map((letter) => {
      const name = chosen.map((r) => r.cells[letter]).filter((v) => v != null && String(v).trim() !== '').map((v) => String(v).trim()).join(' / ');
      return { col: letterIndex(letter), name: name || `第 ${letterIndex(letter)} 列` };
    });
  }

  // ------------------------------------------------------------ 对比
  function refCoord(ref) {
    if (!ref) return '<span class="muted">—</span>';
    const text = `第${ref.row}行 / ${ref.colLetter || ''}列(${ref.col}) / ${ref.a1 || ''}`;
    return `<span class="coord" title="${esc(ref.display || text)}">${esc(text)}</span>`;
  }

  function refFile(ref) {
    if (!ref) return '<span class="muted">—</span>';
    return `<div class="filecell" title="${esc(ref.filePath || '')}">
      <span class="fname">${esc(basename(ref.filePath || ''))}</span>
      <span class="pathsmall mono">${esc(ref.filePath || '')}</span>
      <span class="rowbtns"><button class="linkbtn copybtn" data-copy="${esc(ref.filePath || '')}" title="复制完整路径">复制路径</button></span>
    </div>`;
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

  function filteredItems() {
    const items = Array.isArray((st.result || {}).items) ? st.result.items : [];
    return st.filter === 'all' ? items : items.filter((x) => x.kind === st.filter);
  }

  function renderResult() {
    const d = st.result || {};
    const summary = d.summary || {};
    const warnings = Array.isArray(d.warnings) ? d.warnings : [];
    const items = Array.isArray(d.items) ? d.items : [];
    resultStateEl.textContent = items.length ? `共 ${items.length} 条差异` : '';
    exportEl.disabled = !items.length;

    warningsEl.innerHTML = warnings.length
      ? `<div class="warn-state"><b>告警：</b>${warnings.map((w) => esc(typeof w === 'string' ? w : JSON.stringify(w))).join('<br>')}</div>`
      : '';
    summaryEl.innerHTML = `
      <div class="stat-grid">
        <div class="stat"><b>${num(summary.matched)}</b><span>匹配成功</span></div>
        <div class="stat ok"><b>${num(summary.added)}</b><span>新增（仅右）</span></div>
        <div class="stat err"><b>${num(summary.removed)}</b><span>删除（仅左）</span></div>
        <div class="stat warn"><b>${num(summary.changed)}</b><span>修改</span></div>
        <div class="stat accent"><b>${num(summary.leftRows)}</b><span>左侧数据行</span></div>
        <div class="stat accent"><b>${num(summary.rightRows)}</b><span>右侧数据行</span></div>
      </div>`;

    if (!items.length) {
      listEl.hidden = true;
      pagerEl.hidden = true;
      filterBarEl.hidden = true;
      emptyEl.hidden = false;
      emptyEl.innerHTML = '<b>两表没有差异</b>在选定的主键与选项下，左右两侧数据一致（或都没有差异字段）。';
      return;
    }
    emptyEl.hidden = true;
    listEl.hidden = false;
    filterBarEl.hidden = false;
    filterBarEl.style.display = 'flex';
    listEl.scrollTop = 0;
    renderRows();
  }

  function renderRows() {
    const items = filteredItems();
    const pageSize = 50;
    const pages = Math.max(1, Math.ceil(items.length / pageSize));
    const page = Math.min(pages, Math.floor(st.offset / pageSize) + 1);
    const slice = items.slice((page - 1) * pageSize, page * pageSize);
    filterCountEl.textContent = `当前筛选 ${items.length} 条 / 共 ${(st.result.items || []).length} 条`;
    rowsEl.innerHTML = slice.map((it) => {
      const meta = KIND_META[it.kind] || { label: it.kind || '?', cls: '' };
      return `<tr class="diff-${esc(it.kind || '')}">
        <td class="nowrap"><span class="badge ${esc(meta.cls)}">${esc(meta.label)}</span></td>
        <td class="nowrap">${esc(it.key ?? '')}</td>
        <td>${it.fieldName ? esc(it.fieldName) : '<span class="muted">（整行）</span>'}</td>
        <td class="valcell">${it.leftValue === '' || it.leftValue == null ? '<span class="muted">（空）</span>' : esc(it.leftValue)}</td>
        <td>${refFile(it.leftRef)}</td>
        <td class="coordcell">${refCoord(it.leftRef)}</td>
        <td class="valcell">${it.rightValue === '' || it.rightValue == null ? '<span class="muted">（空）</span>' : esc(it.rightValue)}</td>
        <td>${refFile(it.rightRef)}</td>
        <td class="coordcell">${refCoord(it.rightRef)}</td>
      </tr>`;
    }).join('');
    pagerEl.hidden = pages <= 1;
    if (pages > 1) {
      pagerEl.innerHTML = `
        <span class="pginfo">第 ${page} / ${pages} 页 · 每页 ${pageSize} 条</span>
        <button class="pgbtn" data-page="${page - 1}" ${page <= 1 ? 'disabled' : ''}>‹</button>
        <button class="pgbtn" data-page="${page + 1}" ${page >= pages ? 'disabled' : ''}>›</button>`;
    }
  }

  async function doCompare() {
    if (st.loading) return;
    setError('');
    if (!sideFile(st.left) || !sideSheet(st.left)) { setError('请先选择左侧表的文件与工作表。'); return; }
    if (!sideFile(st.right) || !sideSheet(st.right)) { setError('请先选择右侧表的文件与工作表。'); return; }
    if (!st.keys.length) { setError('请至少指定一个主键列（1 基列号，可用列字母）。'); return; }
    st.loading = true;
    compareBtn.disabled = true;
    compareBtn.textContent = '对比中…';
    summaryEl.innerHTML = '<div class="loading">正在按主键匹配并逐字段比较…</div>';
    emptyEl.hidden = true;
    listEl.hidden = true;
    pagerEl.hidden = true;
    filterBarEl.hidden = true;
    warningsEl.innerHTML = '';
    exportEl.disabled = true;
    try {
      const body = {
        left: { path: sidePath(st.left), sheet: sideSheetName(st.left) },
        right: { path: sidePath(st.right), sheet: sideSheetName(st.right) },
        keys: clone(st.keys),
        options: options(),
      };
      const res = await api.post('/compare', body);
      if (!host.isConnected) return;
      st.result = res.data || {};
      st.offset = 0;
      st.filter = 'all';
      filterEl.value = 'all';
      renderResult();
      const sum = (st.result.summary || {});
      toast(`对比完成：新增 ${num(sum.added)} / 删除 ${num(sum.removed)} / 修改 ${num(sum.changed)}`, 'ok');
    } catch (e) {
      if (!host.isConnected) return;
      summaryEl.innerHTML = '';
      resultStateEl.textContent = '';
      setError(e.message || String(e));
      toast(`对比失败：${e.message || e}`, 'err');
    } finally {
      if (host.isConnected) {
        st.loading = false;
        compareBtn.disabled = false;
        compareBtn.textContent = '开始对比';
      }
    }
  }

  async function doExport() {
    if (!st.result || !(st.result.items || []).length) { toast('请先完成一次对比', 'err'); return; }
    exportEl.disabled = true;
    exportEl.textContent = '导出中…';
    try {
      const body = {
        left: { path: sidePath(st.left), sheet: sideSheetName(st.left) },
        right: { path: sidePath(st.right), sheet: sideSheetName(st.right) },
        keys: clone(st.keys),
        options: options(),
      };
      const res = await api.post('/compare/export', body);
      const d = res.data || {};
      if (d.token) {
        api.download(d.token);
        toast(`已生成对比报告 ${d.name || ''}（${d.count ?? (st.result.items || []).length} 条差异）`, 'ok');
      } else {
        toast(`报告已生成：${d.file || d.path || '（未返回路径）'}`, 'ok');
      }
    } catch (e) {
      toast(`导出失败：${e.message || e}`, 'err');
    } finally {
      if (host.isConnected) {
        exportEl.disabled = false;
        exportEl.textContent = '导出对比报告';
      }
    }
  }

  // ------------------------------------------------------------ 事件
  leftFileEl.addEventListener('change', async () => {
    st.left.fileId = leftFileEl.value;
    await loadSheets('left');
    refreshSideText();
    await loadCandidates();
  });
  rightFileEl.addEventListener('change', async () => {
    st.right.fileId = rightFileEl.value;
    await loadSheets('right');
    refreshSideText();
    await loadCandidates();
  });
  leftSheetEl.addEventListener('change', async () => {
    st.left.sheetIndex = num(leftSheetEl.value);
    refreshSideText();
    await loadCandidates();
  });
  rightSheetEl.addEventListener('change', async () => {
    st.right.sheetIndex = num(rightSheetEl.value);
    refreshSideText();
    await loadCandidates();
  });
  host.querySelector('#keys-parse').addEventListener('click', () => {
    const parsed = normalizeKeys(keysInputEl.value.split(/[,，;；\s]+/));
    if (!parsed.length) { toast('没有解析出有效列号：请输入如 AF 或 32,33', 'err'); return; }
    st.keys = parsed;
    syncKeyUI();
    toast(`主键已设为：${st.keys.map((n) => `${colLetter(n)}（第 ${n} 列）`).join('、')}`, 'ok');
  });
  keysInputEl.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') host.querySelector('#keys-parse').click();
  });
  keysChipsEl.addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-key]');
    if (!btn) return;
    st.keys = st.keys.filter((n) => n !== num(btn.dataset.key));
    syncKeyUI();
  });
  keyCandidatesEl.addEventListener('change', (e) => {
    const cb = e.target.closest('input[type=checkbox]');
    if (!cb) return;
    const n = num(cb.value);
    if (cb.checked) {
      if (!st.keys.includes(n)) st.keys.push(n);
      st.keys.sort((a, b) => a - b);
    } else {
      st.keys = st.keys.filter((x) => x !== n);
    }
    syncKeyUI();
  });
  host.querySelector('#keys-clear').addEventListener('click', () => {
    st.keys = [];
    syncKeyUI();
  });
  rowsEl.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-copy]');
    if (btn) copyPath(btn.dataset.copy);
  });
  filterEl.addEventListener('change', () => {
    st.filter = filterEl.value;
    st.offset = 0;
    renderRows();
  });
  pagerEl.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-page]');
    if (btn && !btn.disabled) {
      st.offset = (num(btn.dataset.page) - 1) * 50;
      renderRows();
    }
  });
  compareBtn.addEventListener('click', doCompare);
  exportEl.addEventListener('click', doExport);
  host.querySelector('#refresh').addEventListener('click', async () => {
    try { await loadFiles(); toast('文件清单已刷新', 'ok'); } catch (e) { toast(`刷新失败：${e.message || e}`, 'err'); }
  });

  // ------------------------------------------------------------ 初始化
  try {
    await loadFiles();
    await loadCandidates();
  } catch (e) {
    if (!host.isConnected) return;
    setError(`初始化失败：${e.message || e}`);
    toast(`对比页加载失败：${e.message || e}`, 'err');
  }
}
