// S6 前端 · 设置页
// 原生 ES module，无框架、无构建、无 CDN。render(host, ctx) 由 app.js 动态调用。
// 覆盖：工作区目录/索引状态、重建索引、文件清单（格式/大小/状态/错误）、审计记录。

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
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const fmtBytes = (n) => {
  n = num(n);
  if (!n) return '0 B';
  const u = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(i ? 1 : 0)} ${u[i]}`;
};
const fmtTime = (ts) => {
  const n = num(ts);
  if (!n) return '';
  const d = new Date(n * 1000);
  const p = (x) => String(x).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
};
const PAGE_SIZE = 50;

const STATUS_BADGE = {
  ok: '<span class="badge ok">正常</span>',
  error: '<span class="badge err">错误</span>',
  skipped: '<span class="badge warn">跳过</span>',
};

export async function render(host, ctx) {
  ensureCss();
  const { api, toast } = ctx;
  document.querySelectorAll('.pg-modal-mask').forEach((n) => n.remove());

  const st = {
    ws: null, files: [], audit: [],
    fileOffset: 0, auditLimit: 200,
    pendingIndex: false,
  };

  host.innerHTML = `
  <div class="page page-settings">
    <section class="panel">
      <div class="page-head">
        <div>
          <h2>设置</h2>
          <p class="sub">工作区目录、索引状态、文件清单与修改审计。扫描目录为递归只读，写回前会自动备份到 data/backup。</p>
        </div>
        <button class="btn ghost sm" id="refresh-all">刷新</button>
      </div>
      <div class="toolbar">
        <input id="set-root" class="pinput grow mono" placeholder="目标文件夹绝对路径（递归扫描）">
        <button id="set-browse" class="btn ghost">浏览…</button>
        <button id="ws-apply" class="btn">切换工作区</button>
        <button id="ws-index" class="btn ghost">重建索引</button>
      </div>
      <div class="progress" id="idx-progress" hidden><i id="idx-bar"></i></div>
      <div class="stat-grid" id="ws-stats">
        <div class="stat"><b>—</b><span>文件数</span></div>
        <div class="stat"><b>—</b><span>工作表数</span></div>
        <div class="stat"><b>—</b><span>单元格数</span></div>
        <div class="stat"><b>—</b><span>索引状态</span></div>
      </div>
      <div class="kvline"><span class="k">索引时间</span><span id="ws-indexed-at" class="mini"></span></div>
      <div id="ws-note" class="warn-state" hidden></div>
      <div id="ws-error" class="error-state" hidden></div>
    </section>

    <section class="panel">
      <div class="subhead"><h3>文件清单</h3><span class="mini" id="files-count"></span></div>
      <div class="toolbar">
        <input id="f-q" class="pinput grow" placeholder="按文件名或路径筛选…">
        <select id="f-fmt" class="pselect"><option value="">全部格式</option></select>
        <button id="f-refresh" class="btn ghost sm">刷新</button>
      </div>
      <div id="f-error" class="error-state" hidden></div>
      <div id="f-empty" class="empty-state"><b>还没有文件</b>请在上方设置工作区目录并重建索引。</div>
      <div id="f-list" class="tablewrap" hidden>
        <table class="grid file-table">
          <thead><tr>
            <th style="min-width:230px">文件</th>
            <th class="nowrap">格式</th>
            <th class="num">大小</th>
            <th class="num">工作表数</th>
            <th class="nowrap">修改时间</th>
            <th class="nowrap">解析状态</th>
            <th>错误信息</th>
            <th class="nowrap">索引时间</th>
          </tr></thead>
          <tbody id="f-rows"></tbody>
        </table>
      </div>
      <div id="f-pager" class="pager" hidden></div>
    </section>

    <section class="panel">
      <div class="subhead">
        <h3>审计记录（修改历史）</h3>
        <div class="toolbar">
          <input id="a-file" class="pinput" style="min-width:260px" placeholder="按文件路径筛选（可选）">
          <button id="a-refresh" class="btn ghost sm">刷新</button>
        </div>
      </div>
      <div id="a-error" class="error-state" hidden></div>
      <div id="a-empty" class="empty-state"><b>暂无修改记录</b>通过检索页编辑单元格后，这里会显示时间、坐标、旧值、新值与备份路径。</div>
      <div id="a-list" class="tablewrap" hidden>
        <table class="grid">
          <thead><tr>
            <th class="nowrap">时间</th>
            <th style="min-width:200px">文件</th>
            <th>坐标（工作表 · 行 · 列 · A1）</th>
            <th>旧值</th>
            <th>新值</th>
            <th class="nowrap">结果</th>
            <th style="min-width:180px">备份</th>
          </tr></thead>
          <tbody id="a-rows"></tbody>
        </table>
      </div>
      <div class="hint">审计日志追加写入 <code>data/audit/*.jsonl</code>；<code>result</code> 为 ok / conflict / rolled_back。</div>
    </section>
  </div>`;

  const wsRootEl = host.querySelector('#set-root');
  const wsStatsEl = host.querySelector('#ws-stats');
  const wsIndexedAtEl = host.querySelector('#ws-indexed-at');
  const wsNoteEl = host.querySelector('#ws-note');
  const wsErrEl = host.querySelector('#ws-error');
  const idxProgressEl = host.querySelector('#idx-progress');
  const idxBarEl = host.querySelector('#idx-bar');
  const fQEl = host.querySelector('#f-q');
  const fFmtEl = host.querySelector('#f-fmt');
  const fCountEl = host.querySelector('#files-count');
  const fErrEl = host.querySelector('#f-error');
  const fEmptyEl = host.querySelector('#f-empty');
  const fListEl = host.querySelector('#f-list');
  const fRowsEl = host.querySelector('#f-rows');
  const fPagerEl = host.querySelector('#f-pager');
  const aFileEl = host.querySelector('#a-file');
  const aErrEl = host.querySelector('#a-error');
  const aEmptyEl = host.querySelector('#a-empty');
  const aListEl = host.querySelector('#a-list');
  const aRowsEl = host.querySelector('#a-rows');

  function setErr(target, msg) {
    target.hidden = !msg;
    target.innerHTML = msg ? `<b>出错了：</b>${esc(msg)}` : '';
  }

  // ------------------------------------------------------------ 工作区
  async function loadStatus() {
    try {
      const res = await api.get('/workspace');
      if (!host.isConnected) return;
      const d = res.data || {};
      st.ws = d;
      const root = d.root || '';
      wsRootEl.value = root;
      const topInput = document.getElementById('ws-root');
      if (topInput && root) topInput.value = root;
      wsStatsEl.innerHTML = `
        <div class="stat accent"><b>${num(d.files)}</b><span>文件数</span></div>
        <div class="stat accent"><b>${esc(d.sheets ?? '—')}</b><span>工作表数</span></div>
        <div class="stat accent"><b>${num(d.cells).toLocaleString('zh-CN')}</b><span>单元格数</span></div>
        <div class="stat ${d.indexed ? 'ok' : 'warn'}"><b>${d.indexed ? '已索引' : '未索引'}</b><span>索引状态</span></div>`;
      wsIndexedAtEl.textContent = [d.indexedAt && `索引：${d.indexedAt}`, d.scannedAt && `扫描：${d.scannedAt}`]
        .filter(Boolean).join('　') || '—';
      const notes = [
        ...(Array.isArray(d.notes) ? d.notes : []),
        ...(Array.isArray(d.errors) ? d.errors : []),
      ];
      wsNoteEl.hidden = !notes.length;
      wsNoteEl.innerHTML = notes.map((x) => esc(typeof x === 'string' ? x : JSON.stringify(x))).join('<br>');
      setErr(wsErrEl, '');
      if (!root) {
        wsNoteEl.hidden = false;
        wsNoteEl.textContent = '还没有指定工作区目录：输入目标文件夹路径后点「切换工作区」。';
      }
    } catch (e) {
      if (!host.isConnected) return;
      setErr(wsErrEl, e.message || String(e));
    }
  }

  async function setWorkspace() {
    const dir = wsRootEl.value.trim();
    if (!dir) { toast('请填写工作区目录路径', 'err'); wsRootEl.focus(); return; }
    const btn = host.querySelector('#ws-apply');
    btn.disabled = true;
    btn.textContent = '切换中…';
    setErr(wsErrEl, '');
    try {
      const res = await api.post('/workspace', { dir });
      const root = (res.data || {}).root || dir;
      toast(`已切换工作区：${root}`, 'ok');
      await Promise.all([loadStatus(), loadFiles()]);
    } catch (e) {
      setErr(wsErrEl, e.message || String(e));
      toast(`切换失败：${e.message || e}`, 'err');
    } finally {
      if (host.isConnected) {
        btn.disabled = false;
        btn.textContent = '切换工作区';
      }
    }
  }

  async function pollIndex(taskId) {
    for (let i = 0; i < 1800; i++) {
      const res = await api.get('/tasks/' + encodeURIComponent(taskId));
      if (!host.isConnected) throw new Error('页面已切换');
      const t = res.data || {};
      const pct = num(t.progress);
      idxProgressEl.hidden = false;
      idxBarEl.style.width = `${Math.max(0, Math.min(100, pct))}%`;
      idxProgressEl.title = t.message || `${pct}%`;
      if (t.status === 'done') return t.result;
      if (t.status === 'error') throw new Error(t.error || '索引失败');
      await sleep(500);
    }
    throw new Error('索引任务超时');
  }

  async function rebuildIndex() {
    if (st.pendingIndex) return;
    const btn = host.querySelector('#ws-index');
    st.pendingIndex = true;
    btn.disabled = true;
    btn.textContent = '重建中…';
    setErr(wsErrEl, '');
    idxProgressEl.hidden = false;
    idxBarEl.style.width = '0%';
    try {
      const res = await api.post('/index/refresh', {});
      const d = res.data || {};
      if (d.taskId) {
        await pollIndex(d.taskId);
        toast('索引已重建完成', 'ok');
      } else {
        toast(d.mock ? '示例数据模式：索引刷新接口未返回任务，已重新读取状态' : '索引刷新请求已提交', 'info');
      }
      await Promise.all([loadStatus(), loadFiles()]);
    } catch (e) {
      setErr(wsErrEl, e.message || String(e));
      toast(`重建索引失败：${e.message || e}`, 'err');
    } finally {
      st.pendingIndex = false;
      if (host.isConnected) {
        btn.disabled = false;
        btn.textContent = '重建索引';
        idxProgressEl.hidden = true;
        idxBarEl.style.width = '0%';
      }
    }
  }

  // ------------------------------------------------------------ 文件清单
  async function loadFiles() {
    setErr(fErrEl, '');
    try {
      const res = await api.get('/files');
      if (!host.isConnected) return;
      const d = res.data || {};
      st.files = Array.isArray(d.files) ? d.files : [];
      const fmts = [...new Set(st.files.map((f) => f.fmt).filter(Boolean))].sort();
      const cur = fFmtEl.value;
      fFmtEl.innerHTML = '<option value="">全部格式</option>' + fmts.map((x) => (
        `<option value="${esc(x)}" ${x === cur ? 'selected' : ''}>${esc(x)}</option>`
      )).join('');
      st.fileOffset = 0;
      renderFiles();
    } catch (e) {
      if (!host.isConnected) return;
      setErr(fErrEl, e.message || String(e));
    }
  }

  function filteredFiles() {
    const q = fQEl.value.trim().toLowerCase();
    const fmt = fFmtEl.value;
    return st.files.filter((f) => {
      if (fmt && f.fmt !== fmt) return false;
      if (q && !(`${f.name || ''}\n${f.path || ''}`.toLowerCase().includes(q))) return false;
      return true;
    });
  }

  function renderFiles() {
    const list = filteredFiles();
    const total = list.length;
    const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    const page = Math.min(pages, Math.floor(st.fileOffset / PAGE_SIZE) + 1);
    const slice = list.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
    fCountEl.textContent = total
      ? `共 ${total} 个文件${st.files.length !== total ? `（已从 ${st.files.length} 个中筛选）` : ''}`
      : '';
    if (!st.files.length) {
      fListEl.hidden = true;
      fPagerEl.hidden = true;
      fEmptyEl.hidden = false;
      fEmptyEl.innerHTML = '<b>还没有文件</b>请在上方设置工作区目录并重建索引。';
      return;
    }
    if (!total) {
      fListEl.hidden = true;
      fPagerEl.hidden = true;
      fEmptyEl.hidden = false;
      fEmptyEl.innerHTML = '<b>没有匹配的文件</b>换个筛选条件试试。';
      return;
    }
    fEmptyEl.hidden = true;
    fListEl.hidden = false;
    fRowsEl.innerHTML = slice.map((f) => {
      const status = STATUS_BADGE[f.status] || `<span class="badge">${esc(f.status || '未知')}</span>`;
      return `<tr>
        <td>
          <div class="filecell" title="${esc(f.path || '')}">
            <span class="fname">${esc(f.name || basename(f.path))}</span>
            <span class="pathsmall mono">${esc(f.path || '')}</span>
          </div>
        </td>
        <td class="nowrap"><span class="badge">${esc(f.fmt || '?')}</span></td>
        <td class="num">${fmtBytes(f.size)}</td>
        <td class="num">${esc(f.sheetsCount ?? '')}</td>
        <td class="nowrap mini">${fmtTime(f.mtime)}</td>
        <td class="nowrap">${status}</td>
        <td class="mini">${f.error ? esc(f.error) : '<span class="muted">—</span>'}</td>
        <td class="nowrap mini">${esc(f.indexedAt || '—')}</td>
      </tr>`;
    }).join('');
    fPagerEl.hidden = pages <= 1;
    if (pages > 1) {
      const win = [];
      const from = Math.max(1, page - 3);
      const to = Math.min(pages, from + 6);
      for (let i = from; i <= to; i++) win.push(i);
      fPagerEl.innerHTML = `
        <span class="pginfo">第 ${page} / ${pages} 页 · 每页 ${PAGE_SIZE} 个</span>
        <button class="pgbtn" data-page="${page - 1}" ${page <= 1 ? 'disabled' : ''}>‹</button>
        ${win.map((p) => `<button class="pgbtn ${p === page ? 'active' : ''}" data-page="${p}">${p}</button>`).join('')}
        <button class="pgbtn" data-page="${page + 1}" ${page >= pages ? 'disabled' : ''}>›</button>`;
    }
  }

  // ------------------------------------------------------------ 审计
  async function loadAudit() {
    setErr(aErrEl, '');
    const file = aFileEl.value.trim();
    try {
      const res = await api.get('/audit', { limit: st.auditLimit, file: file || undefined });
      if (!host.isConnected) return;
      const d = res.data || {};
      st.audit = Array.isArray(d.entries) ? d.entries
        : Array.isArray(d) ? d
          : Array.isArray(d.audit) ? d.audit : [];
      renderAudit();
    } catch (e) {
      if (!host.isConnected) return;
      setErr(aErrEl, e.message || String(e));
      aListEl.hidden = true;
      aEmptyEl.hidden = true;
    }
  }

  function renderAudit() {
    if (!st.audit.length) {
      aListEl.hidden = true;
      aEmptyEl.hidden = false;
      aEmptyEl.innerHTML = '<b>暂无修改记录</b>通过检索页编辑单元格后，这里会显示时间、坐标、旧值、新值与备份路径。';
      return;
    }
    aEmptyEl.hidden = true;
    aListEl.hidden = false;
    aRowsEl.innerHTML = st.audit.slice(0, st.auditLimit).map((entry) => {
      const ref = entry.ref || {};
      const result = entry.result || 'ok';
      const badge = result === 'ok'
        ? '<span class="badge ok">成功</span>'
        : result === 'conflict'
          ? '<span class="badge err">冲突拒绝</span>'
          : result === 'rolled_back'
            ? '<span class="badge warn">已回滚</span>'
            : `<span class="badge">${esc(result)}</span>`;
      return `<tr>
        <td class="nowrap mini">${esc(entry.ts || '')}</td>
        <td><div class="filecell" title="${esc(ref.filePath || '')}">
          <span class="fname">${esc(basename(ref.filePath || ''))}</span>
          <span class="pathsmall mono">${esc(ref.filePath || '')}</span>
        </div></td>
        <td class="mono wrapdot">${esc(ref.display || `${ref.sheet || ''} · 第${ref.row ?? '?'}行 · ${ref.a1 || ''}`)}</td>
        <td class="hitv">${esc(entry.oldValue ?? '')}</td>
        <td class="hitv">${esc(entry.newValue ?? '')}</td>
        <td class="nowrap">${badge}</td>
        <td class="mini wrapdot">${esc(entry.backupPath || '—')}</td>
      </tr>`;
    }).join('');
  }

  // ------------------------------------------------------------ 事件
  host.querySelector('#ws-apply').addEventListener('click', setWorkspace);
  wsRootEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') setWorkspace(); });
  host.querySelector('#set-browse').addEventListener('click', async () => {
    const btn = host.querySelector('#set-browse');
    btn.disabled = true;
    btn.textContent = '选择中…';
    try {
      const res = await api.post('/browse', { title: '选择目标文件夹', initial: wsRootEl.value.trim() });
      const p = res.data && res.data.path;
      if (p) {
        wsRootEl.value = p;
        await setWorkspace();
      }
    } catch (e) {
      toast('打开文件夹选择框失败：' + (e.message || e), 'err');
    } finally {
      if (host.isConnected) { btn.disabled = false; btn.textContent = '浏览…'; }
    }
  });
  host.querySelector('#ws-index').addEventListener('click', rebuildIndex);
  host.querySelector('#refresh-all').addEventListener('click', async () => {
    await Promise.all([loadStatus(), loadFiles(), loadAudit()]);
    toast('已刷新', 'ok');
  });
  fQEl.addEventListener('input', () => { st.fileOffset = 0; renderFiles(); });
  fFmtEl.addEventListener('change', () => { st.fileOffset = 0; renderFiles(); });
  host.querySelector('#f-refresh').addEventListener('click', loadFiles);
  fPagerEl.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-page]');
    if (btn && !btn.disabled) {
      st.fileOffset = (num(btn.dataset.page) - 1) * PAGE_SIZE;
      renderFiles();
    }
  });
  host.querySelector('#a-refresh').addEventListener('click', loadAudit);
  aFileEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') loadAudit(); });

  // ------------------------------------------------------------ 初始化
  try {
    await Promise.all([loadStatus(), loadFiles(), loadAudit()]);
  } catch (e) {
    if (!host.isConnected) return;
    setErr(wsErrEl, e.message || String(e));
  }
}
