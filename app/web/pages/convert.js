// S6 前端 · 模板转换页
// 原生 ES module，无框架、无构建、无 CDN。render(host, ctx) 由 app.js 动态调用。
// 流程：选来源文件 + 映射方案 + 目标模板 → POST /api/convert → 轮询任务 → 展示结果并下载。

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

export async function render(host, ctx) {
  ensureCss();
  const { api, toast } = ctx;
  document.querySelectorAll('.pg-modal-mask').forEach((n) => n.remove());

  const st = {
    files: [], profiles: [], templates: [], sheets: [],
    fileId: '', sheetIndex: 0, profileId: '', templateId: '',
    mode: 'append', busy: false, result: null,
  };

  host.innerHTML = `
  <div class="page page-convert">
    <section class="panel">
      <div class="page-head">
        <div>
          <h2>模板转换</h2>
          <p class="sub">把来源表按映射方案写入目标模板：保留模板原有样式、合并区、列宽与公式，输出新的 xlsx。</p>
        </div>
        <button class="btn ghost sm" id="refresh">刷新清单</button>
      </div>
      <div class="toolbar">
        <span class="mini">来源文件</span>
        <select id="src-file" class="pselect" style="min-width:280px"></select>
        <span class="mini">来源工作表</span>
        <select id="src-sheet" class="pselect" style="min-width:200px"></select>
        <span class="mini">映射方案</span>
        <select id="profile" class="pselect" style="min-width:220px"></select>
        <span class="mini">目标模板</span>
        <select id="tpl" class="pselect" style="min-width:220px"></select>
      </div>
      <div class="toolbar" style="margin-top:10px">
        <span class="mini">写入模式</span>
        <select id="mode" class="pselect">
          <option value="append">追加到数据区（append）</option>
          <option value="replace">替换数据区（replace）</option>
        </select>
        <input id="out-name" class="pinput grow" placeholder="输出文件名（可选，留空自动按时间命名，如：台账_20241109_1540.xlsx）">
        <button id="convert" class="btn">开始转换</button>
      </div>
      <div class="hint">转换是长任务：提交后轮询 <code>/api/tasks/&lt;id&gt;</code> 显示进度，完成后生成文件到 <code>data/output/</code>。来源表的小计/占位行不会当作数据写入。</div>
    </section>

    <div class="grid2">
      <section class="panel">
        <div class="subhead"><h3>方案摘要</h3><span class="mini" id="profile-state"></span></div>
        <div id="summary"><div class="empty-state">选择来源文件与映射方案后显示摘要。</div></div>
      </section>
      <section class="panel">
        <div class="subhead"><h3>任务进度</h3><span class="mini" id="task-state">未开始</span></div>
        <div class="progress"><i id="bar"></i></div>
        <div id="task-detail" class="muted">提交转换后可在这里看到进度与错误信息。</div>
        <div id="task-error" class="error-state" hidden></div>
      </section>
    </div>

    <section class="panel">
      <div class="subhead"><h3>转换结果</h3><span class="mini" id="result-state"></span></div>
      <div id="result"><div class="empty-state">还没有生成文件。选好来源与方案后点「开始转换」。</div></div>
    </section>
  </div>`;

  const srcFileEl = host.querySelector('#src-file');
  const srcSheetEl = host.querySelector('#src-sheet');
  const profileEl = host.querySelector('#profile');
  const tplEl = host.querySelector('#tpl');
  const modeEl = host.querySelector('#mode');
  const outNameEl = host.querySelector('#out-name');
  const convertBtn = host.querySelector('#convert');
  const summaryEl = host.querySelector('#summary');
  const profileStateEl = host.querySelector('#profile-state');
  const taskStateEl = host.querySelector('#task-state');
  const barEl = host.querySelector('#bar');
  const taskDetailEl = host.querySelector('#task-detail');
  const taskErrEl = host.querySelector('#task-error');
  const resultEl = host.querySelector('#result');
  const resultStateEl = host.querySelector('#result-state');

  function setBusy(busy) {
    st.busy = busy;
    convertBtn.disabled = busy;
    convertBtn.textContent = busy ? '转换中…' : '开始转换';
  }

  function setBar(pct, text) {
    const p = Math.max(0, Math.min(100, num(pct)));
    barEl.style.width = `${p}%`;
    taskStateEl.textContent = text || (p ? `进行中 ${p}%` : '未开始');
  }

  function setTaskError(msg) {
    taskErrEl.hidden = !msg;
    taskErrEl.innerHTML = msg ? `<b>转换失败：</b>${esc(msg)}` : '';
  }

  function renderSummary() {
    const profile = st.profiles.find((p) => String(p.id) === String(st.profileId));
    const tpl = st.templates.find((t) => String(t.id) === String(st.templateId));
    const file = st.files.find((f) => String(f.id) === String(st.fileId));
    profileStateEl.textContent = profile ? `${(profile.rules || []).length} 条规则` : '';
    if (!profile && !tpl) {
      summaryEl.innerHTML = '<div class="empty-state">选择来源文件与映射方案后显示摘要。</div>';
      return;
    }
    summaryEl.innerHTML = `
      <div class="kv-grid" style="margin-top:0">
        <div><span>来源文件</span><b>${esc(file ? (file.name || basename(file.path)) : '（未选择）')}</b></div>
        <div><span>来源格式 / 工作表</span><b>${esc(file ? file.fmt : '')} / ${esc((st.sheets[st.sheetIndex] || {}).name || '（未选择）')}</b></div>
        <div><span>映射方案</span><b>${esc(profile ? (profile.name || profile.id) : '（未选择）')}</b></div>
        <div><span>模板</span><b>${esc(tpl ? (tpl.name || tpl.id) : '（未选择）')}</b></div>
        <div><span>模板字段数</span><b>${esc(tpl ? (tpl.fields || []).length : '')}</b></div>
        <div><span>模板数据起始</span><b>${esc(tpl && tpl.dataStart ? `第 ${tpl.dataStart[0]} 行第 ${tpl.dataStart[1]} 列` : '')}</b></div>
      </div>
      ${profile && !(tpl && (tpl.fields || []).length) ? '<div class="warn-state">该模板没有解析出字段列表，转换前建议先到映射页确认字段与数据起始单元格。</div>' : ''}`;
  }

  async function pollTask(taskId) {
    for (let i = 0; i < 1800; i++) {
      const res = await api.get('/tasks/' + encodeURIComponent(taskId));
      if (!host.isConnected) throw new Error('页面已切换');
      const t = res.data || {};
      const pct = num(t.progress);
      setBar(pct, t.message || `进行中 ${pct}%`);
      taskDetailEl.textContent = t.message || (t.status === 'running' ? '正在生成…' : '');
      if (t.status === 'done') return t.result;
      if (t.status === 'error') throw new Error(t.error || '任务失败');
      // 进度不前进时也给出提示，避免用户以为卡死
      if (!t.progress && i > 0 && i % 20 === 0) taskDetailEl.textContent = '任务仍在运行，请稍候…';
      await sleep(400);
    }
    throw new Error('任务超时（超过 12 分钟）');
  }

  function showResult(data) {
    st.result = data || {};
    const path = st.result.path || st.result.file || '';
    const token = st.result.token || st.result.downloadToken || '';
    const count = st.result.count ?? st.result.rows ?? '—';
    resultStateEl.textContent = path ? '生成成功' : '';
    resultEl.innerHTML = `
      <div class="result-file">
        <div><b>生成文件：</b><span class="mono wrapdot">${esc(path || '（后端未返回文件路径）')}</span></div>
        <div class="mini" style="margin-top:4px">数据行数：${esc(count)}${st.result.mock ? ' · 当前为 mock 示例数据' : ''}</div>
      </div>
      <div class="toolbar">
        ${token ? '<button class="btn" id="download">下载生成文件</button>' : ''}
        <button class="btn ghost" id="open-folder" ${path ? '' : 'disabled'}>打开所在文件夹</button>
        <button class="btn ghost" id="copy-path" ${path ? '' : 'disabled'}>复制路径</button>
        ${path ? '' : '<span class="mini">后端没返回下载令牌（token），请让管理员在 /api/convert 完成时注册下载令牌。</span>'}
      </div>`;
    const dlBtn = resultEl.querySelector('#download');
    if (dlBtn) {
      dlBtn.addEventListener('click', () => {
        api.download(token);
        toast(`开始下载：${basename(path) || '生成文件'}`, 'ok');
      });
    }
    const openBtn = resultEl.querySelector('#open-folder');
    if (openBtn) {
      openBtn.addEventListener('click', async () => {
        try {
          await api.post('/open', { path, mode: 'folder' });
          toast('已在资源管理器中定位生成文件', 'ok');
        } catch (e) {
          toast(`打开失败：${e.message || e}`, 'err');
        }
      });
    }
    const copyBtn = resultEl.querySelector('#copy-path');
    if (copyBtn) {
      copyBtn.addEventListener('click', async () => {
        try {
          await navigator.clipboard.writeText(path);
          toast('路径已复制到剪贴板', 'ok');
        } catch {
          toast('复制失败，请手动选择路径', 'err');
        }
      });
    }
  }

  async function loadSheets() {
    st.sheets = [];
    st.sheetIndex = 0;
    srcSheetEl.innerHTML = '<option value="">读取中…</option>';
    if (!st.fileId) {
      srcSheetEl.innerHTML = '<option value="">（无文件）</option>';
      return;
    }
    try {
      const res = await api.get(`/files/${encodeURIComponent(st.fileId)}/sheets`);
      if (!host.isConnected) return;
      const d = res.data || {};
      st.sheets = Array.isArray(d.sheets) ? d.sheets : [];
      srcSheetEl.innerHTML = st.sheets.length
        ? st.sheets.map((s, i) => (
          `<option value="${i}" ${i === 0 ? 'selected' : ''}>${esc(s.name || ('Sheet' + (i + 1)))}（${num(s.nRows)}×${num(s.nCols)}）</option>`
        )).join('')
        : '<option value="">（该文件没有工作表）</option>';
      st.sheetIndex = 0;
    } catch (e) {
      srcSheetEl.innerHTML = '<option value="">工作表读取失败</option>';
      toast(`来源工作表读取失败：${e.message || e}`, 'err');
    }
  }

  async function doConvert() {
    if (st.busy) return;
    const file = st.files.find((f) => String(f.id) === String(st.fileId));
    if (!file) { toast('请先选择来源文件', 'err'); return; }
    if (!st.sheetIndex && st.sheets.length && st.sheets[0]) st.sheetIndex = 0;
    if (!st.profileId) { toast('请先选择映射方案（可到「字段映射」页保存）', 'err'); return; }
    if (!st.templateId) { toast('请先选择目标模板', 'err'); return; }
    setBusy(true);
    setTaskError('');
    setBar(0, '已提交，等待任务开始…');
    taskDetailEl.textContent = '正在提交转换任务…';
    resultEl.innerHTML = '<div class="loading">转换中，请稍候…</div>';
    resultStateEl.textContent = '';
    try {
      const sheet = st.sheets[st.sheetIndex] || {};
      const source = {
        fileId: file.id,
        filePath: file.path,
        path: file.path,
        sheetIndex: st.sheetIndex,
        sheet: sheet.name || undefined,
      };
      const body = {
        // 契约 §2.4：sources[]；同时把首个来源摊平到顶层，兼容按单表读取的服务实现
        sources: [source],
        ...source,
        profileId: st.profileId,
        templateId: st.templateId,
        mode: modeEl.value,
        outName: outNameEl.value.trim() || null,
      };
      const res = await api.post('/convert', body);
      if (!host.isConnected) return;
      const d = res.data || {};
      let out;
      if (d.taskId) {
        setBar(0, '任务已开始');
        out = await pollTask(d.taskId);
      } else {
        // 服务未提供长任务时直接使用同步结果
        setBar(100, '完成');
        out = d;
      }
      if (!host.isConnected) return;
      setBar(100, '完成');
      showResult(out || {});
      toast('转换完成', 'ok');
    } catch (e) {
      if (!host.isConnected) return;
      setBar(0, '失败');
      taskDetailEl.textContent = '转换失败。';
      setTaskError(e.message || String(e));
      resultEl.innerHTML = '<div class="empty-state">没有生成文件。</div>';
      toast(`转换失败：${e.message || e}`, 'err');
    } finally {
      if (host.isConnected) setBusy(false);
    }
  }

  async function loadBase() {
    const res = await Promise.all([api.get('/files'), api.get('/profiles'), api.get('/templates')]);
    if (!host.isConnected) return;
    const f = res[0].data || {};
    const p = res[1].data || {};
    const t = res[2].data || {};
    st.files = Array.isArray(f.files) ? f.files : [];
    st.profiles = Array.isArray(p.profiles) ? p.profiles : [];
    st.templates = Array.isArray(t.templates) ? t.templates : [];

    srcFileEl.innerHTML = st.files.length
      ? st.files.map((x) => `<option value="${esc(x.id)}" title="${esc(x.path)}">${esc(x.name || basename(x.path))}（${esc(x.fmt || '?')}）</option>`).join('')
      : '<option value="">（工作区没有文件）</option>';
    if (st.fileId && st.files.some((x) => String(x.id) === String(st.fileId))) srcFileEl.value = st.fileId;
    else st.fileId = (st.files[0] || {}).id || '';
    await loadSheets();

    profileEl.innerHTML = st.profiles.length
      ? '<option value="">请选择映射方案…</option>' + st.profiles.map((x) => (
        `<option value="${esc(x.id)}">${esc(x.name || x.id)}（${(x.rules || []).length} 条规则）</option>`
      )).join('')
      : '<option value="">（还没有保存的映射方案）</option>';
    if (st.profileId && st.profiles.some((x) => String(x.id) === String(st.profileId))) profileEl.value = st.profileId;
    else st.profileId = '';

    tplEl.innerHTML = st.templates.length
      ? '<option value="">请选择目标模板…</option>' + st.templates.map((x) => (
        `<option value="${esc(x.id)}">${esc(x.name || x.id)}（${(x.fields || []).length} 字段）</option>`
      )).join('')
      : '<option value="">（还没有登记模板）</option>';
    if (st.templateId && st.templates.some((x) => String(x.id) === String(st.templateId))) tplEl.value = st.templateId;

    // 根据方案自动带出模板
    const profile = st.profiles.find((x) => String(x.id) === String(st.profileId));
    if (profile && !st.templateId && profile.templateId) {
      st.templateId = profile.templateId;
      if (st.templates.some((x) => String(x.id) === String(st.templateId))) tplEl.value = st.templateId;
    }
    renderSummary();
  }

  // ------------------------------------------------------------ 事件
  srcFileEl.addEventListener('change', async () => {
    st.fileId = srcFileEl.value;
    await loadSheets();
    renderSummary();
  });
  srcSheetEl.addEventListener('change', () => {
    st.sheetIndex = num(srcSheetEl.value);
    renderSummary();
  });
  profileEl.addEventListener('change', () => {
    st.profileId = profileEl.value;
    const p = st.profiles.find((x) => String(x.id) === String(st.profileId));
    if (p && p.templateId && st.templates.some((t) => String(t.id) === String(p.templateId))) {
      st.templateId = p.templateId;
      tplEl.value = p.templateId;
    }
    renderSummary();
  });
  tplEl.addEventListener('change', () => { st.templateId = tplEl.value; renderSummary(); });
  modeEl.addEventListener('change', () => { st.mode = modeEl.value; });
  convertBtn.addEventListener('click', doConvert);
  host.querySelector('#refresh').addEventListener('click', async () => {
    try { await loadBase(); toast('清单已刷新', 'ok'); } catch (e) { toast(`刷新失败：${e.message || e}`, 'err'); }
  });

  // ------------------------------------------------------------ 初始化
  try {
    await loadBase();
  } catch (e) {
    if (!host.isConnected) return;
    host.querySelector('#summary').innerHTML = `<div class="error-state">初始化失败：${esc(e.message || e)}</div>`;
    toast(`转换页加载失败：${e.message || e}`, 'err');
  }
}
