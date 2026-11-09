// 前端骨架（Lead 维护）：路由 + API 封装。页面实现放在 pages/*.js。
// 不依赖任何第三方库，也不做构建。

const API_BASE = '/api';

async function request(path, { method = 'GET', body, params } = {}) {
  let url = API_BASE + path;
  if (params) {
    const q = new URLSearchParams();
    for (const [k, v] of Object.entries(params)) {
      if (v !== undefined && v !== null && v !== '') q.set(k, v);
    }
    const s = q.toString();
    if (s) url += (url.includes('?') ? '&' : '?') + s;
  }
  const init = { method, headers: {} };
  if (body !== undefined) {
    init.headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(body);
  }
  const res = await fetch(url, init);
  let data;
  try { data = await res.json(); } catch { data = { ok: false, message: await res.text() }; }
  if (!res.ok || data.ok === false) {
    const err = new Error(data.message || `HTTP ${res.status}`);
    err.code = data.code;
    err.detail = data.detail;
    throw err;
  }
  return data;
}

export const api = {
  get: (path, params) => request(path, { params }),
  post: (path, body) => request(path, { method: 'POST', body: body ?? {} }),
  put: (path, body) => request(path, { method: 'PUT', body: body ?? {} }),
  del: (path) => request(path, { method: 'DELETE' }),
  download: (token) => window.open(`${API_BASE}/download/${encodeURIComponent(token)}`, '_blank'),
};

// ---------------------------------------------------------------- 小工具
export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

export function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
    else if (v !== undefined && v !== null) node.setAttribute(k, v);
  }
  for (const c of [].concat(children)) {
    if (c == null) continue;
    node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  }
  return node;
}

let toastTimer = null;
export function toast(msg, kind = 'info') {
  const box = document.getElementById('toast');
  if (!box) return;
  box.textContent = msg;
  box.className = 'toast show ' + kind;
  box.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { box.className = 'toast'; box.hidden = true; }, 3200);
}

export function fmtBytes(n) {
  if (!n) return '0 B';
  const u = ['B', 'KB', 'MB', 'GB'];
  let i = 0;
  while (n >= 1024 && i < u.length - 1) { n /= 1024; i++; }
  return `${n.toFixed(i ? 1 : 0)} ${u[i]}`;
}

export function fmtTime(ts) {
  if (!ts) return '';
  const d = new Date(ts * 1000);
  const p = (x) => String(x).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

// ---------------------------------------------------------------- 路由
const PAGES = {
  search: { title: '检索', module: './pages/search.js' },
  mapping: { title: '字段映射', module: './pages/mapping.js' },
  convert: { title: '模板转换', module: './pages/convert.js' },
  compare: { title: '数据对比', module: './pages/compare.js' },
  settings: { title: '设置', module: './pages/settings.js' },
};

const state = {
  page: 'search',
  root: '',
  params: {},
};

// 支持 #search?q=xx&mode=all 形式（深链/分享/刷新保留查询）
function parseHash() {
  const raw = (location.hash || '').replace(/^#/, '');
  const qi = raw.indexOf('?');
  const name = qi >= 0 ? raw.slice(0, qi) : raw;
  const params = {};
  if (qi >= 0) {
    try {
      new URLSearchParams(raw.slice(qi + 1)).forEach((v, k) => { params[k] = v; });
    } catch { /* 忽略非法参数 */ }
  }
  return { page: PAGES[name] ? name : 'search', params };
}

function hashFor(page, params) {
  const qs = new URLSearchParams(params || {}).toString();
  return '#' + page + (qs ? '?' + qs : '');
}

async function render(page, params) {
  if (!PAGES[page]) page = 'search';
  state.page = page;
  state.params = params || {};
  // 地址栏同步（含查询参数）
  const want = hashFor(page, state.params);
  if (location.hash !== want) {
    try { history.replaceState(null, '', want); } catch { /* file:// 下忽略 */ }
  }
  for (const btn of document.querySelectorAll('#tabs .tab')) {
    btn.classList.toggle('active', btn.dataset.page === page);
  }
  const host = document.getElementById('app');
  host.innerHTML = '';
  const def = PAGES[page];
  try {
    const mod = await import(def.module);
    host.classList.remove('empty');
    await mod.render(host, { api, el, toast, state, app });
  } catch (e) {
    if (e && /Failed to fetch dynamically imported module|Cannot find module|404/.test(String(e.message || e))) {
      host.innerHTML = `<div class="placeholder"><h3>${def.title}（页面开发中）</h3>
        <p>该页面由前端切片实现：<code>app/web/pages/${page}.js</code></p></div>`;
      host.classList.add('empty');
    } else {
      console.error(e);
      host.innerHTML = `<div class="placeholder"><h3>页面加载失败</h3><pre>${String(e && e.stack || e)}</pre></div>`;
    }
  }
}

// ---------------------------------------------------------------- 全局应用
export const app = {
  state,
  async init() {
    document.getElementById('tabs').addEventListener('click', (e) => {
      const btn = e.target.closest('.tab');
      if (btn) render(btn.dataset.page);
    });

    const rootInput = document.getElementById('ws-root');
    const stateEl = document.getElementById('index-state');

    document.getElementById('ws-apply').addEventListener('click', async () => {
      try {
        const res = await api.post('/workspace', { dir: rootInput.value.trim() });
        state.root = (res.data && res.data.root) || rootInput.value.trim();
        toast('已切换工作区：' + state.root, 'ok');
        await refreshStatus();
      } catch (e) { toast(e.message, 'err'); }
    });

    // 浏览…：调本地服务弹出系统"选择文件夹"对话框（浏览器本身拿不到绝对路径）
    document.getElementById('ws-browse').addEventListener('click', async () => {
      const btn = document.getElementById('ws-browse');
      btn.disabled = true;
      btn.textContent = '选择中…';
      try {
        const res = await api.post('/browse', {
          title: '选择表格所在文件夹',
          initial: (rootInput.value || '').trim(),
        });
        const p = res.data && res.data.path;
        if (p) {
          rootInput.value = p;
          document.getElementById('ws-apply').click();
        }
      } catch (e) {
        toast('打开文件夹选择框失败：' + (e.message || e), 'err');
      } finally {
        btn.disabled = false;
        btn.textContent = '浏览…';
      }
    });

    document.getElementById('ws-refresh').addEventListener('click', async () => {
      try {
        const res = await api.post('/index/refresh', {});
        if (res.data && res.data.taskId) {
          stateEl.textContent = '索引中…';
          await pollTask(res.data.taskId, stateEl);
        }
        await refreshStatus();
        toast('索引已刷新', 'ok');
      } catch (e) { toast(e.message, 'err'); }
    });

    async function refreshStatus() {
      try {
        const res = await api.get('/workspace');
        const d = res.data || {};
        state.root = d.root || state.root;
        if (!rootInput.value && state.root) rootInput.value = state.root;
        stateEl.textContent = d.indexed
          ? `已索引 ${d.files ?? 0} 个文件${d.cells ? ` / ${d.cells} 个单元格` : ''}`
          : '未索引';
        stateEl.classList.toggle('ok', !!d.indexed);
      } catch (e) { stateEl.textContent = '服务未连接'; }
    }

    async function pollTask(taskId, label) {
      for (let i = 0; i < 600; i++) {
        const res = await api.get('/tasks/' + taskId);
        const t = res.data || {};
        if (label) label.textContent = `${t.kind || '任务'}：${t.progress || 0}%`;
        if (t.status === 'done') return t.result;
        if (t.status === 'error') throw new Error(t.error || '任务失败');
        await new Promise((r) => setTimeout(r, 500));
      }
      throw new Error('任务超时');
    }
    app.pollTask = pollTask;

    try {
      const h = await api.get('/health');
      state.root = (h.data && h.data.root) || '';
      if (state.root) rootInput.value = state.root;
    } catch { /* 忽略 */ }
    await refreshStatus();
    const start = parseHash();
    render(start.page, start.params);
    window.addEventListener('hashchange', () => {
      const h = parseHash();
      const sameParams = JSON.stringify(h.params) === JSON.stringify(state.params);
      if (PAGES[h.page] && (h.page !== state.page || !sameParams)) render(h.page, h.params);
    });
  },
};

app.init();
