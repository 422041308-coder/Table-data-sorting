// S6 前端 · 字段映射页
// 原生 ES module，无框架、无构建、无 CDN。render(host, ctx) 由 app.js 动态调用。
// 覆盖：来源表/模板选择、双侧字段、自动推荐、映射矩阵（增删改/转置）、逐单元格预演、保存方案。

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

function colLetter(col) {
  let n = Number(col);
  if (!Number.isFinite(n) || n < 1) return '';
  let out = '';
  while (n) { n--; out = String.fromCharCode(65 + (n % 26)) + out; n = Math.floor(n / 26); }
  return out;
}
function letterIndex(letter) {
  let n = 0;
  for (const ch of String(letter || '').trim().toUpperCase()) {
    const c = ch.charCodeAt(0);
    if (c < 65 || c > 90) continue;
    n = n * 26 + (c - 64);
  }
  return n;
}

const MODES = [
  ['cell', '逐格 cell'], ['block', '区域块 block'], ['transpose', '区域转置 transpose'],
  ['filldown', '向下填充 filldown'], ['repeat', '常量重复 repeat'],
];
const TRANSFORMS = [
  ['identity', '原样 identity'], ['number', '数值 number'], ['date', '日期 date'],
  ['concat', '拼接 concat'], ['constant', '常量 constant'],
];
const ON_MISSING = [['blank', '留空 blank'], ['skip', '跳过 skip'], ['error', '报错 error']];
const SEL_KINDS = [
  ['fixed', '固定坐标 fixed'], ['record', '记录行+偏移 record'],
  ['header', '按字段名 header'], ['span', '连续N个 span'],
];
const opts = (list, cur) => list.map(([v, label]) => (
  `<option value="${esc(v)}" ${String(v) === String(cur) ? 'selected' : ''}>${esc(label)}</option>`
)).join('');

const getPath = (obj, path) => path.split('.').reduce((o, k) => (o == null ? undefined : o[k]), obj);
function setPath(obj, path, value) {
  const keys = path.split('.');
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) {
    if (cur[keys[i]] == null || typeof cur[keys[i]] !== 'object') cur[keys[i]] = {};
    cur = cur[keys[i]];
  }
  cur[keys[keys.length - 1]] = value;
}

export async function render(host, ctx) {
  ensureCss();
  const { api, toast } = ctx;
  document.querySelectorAll('.pg-modal-mask').forEach((n) => n.remove());

  const st = {
    files: [], templates: [], profiles: [], sheets: [],
    fileId: '', sheetIndex: 0, templateId: '',
    sourceFields: [], suggestions: [], rules: [],
    sourceSignature: '', suggestMeta: null,
    preview: null, previewOffset: 0, loadedProfileId: '',
  };

  host.innerHTML = `
  <div class="page page-mapping">
    <section class="panel">
      <div class="page-head">
        <div>
          <h2>字段映射</h2>
          <p class="sub">选择来源表与目标模板，自动推荐字段对应，也可用坐标规则自由指定（支持「A 表某行 → B 表某列」转置），预演确认后保存为可复用方案。</p>
        </div>
        <div class="toolbar">
          <button class="btn ghost sm" id="refresh-all">刷新清单</button>
          <button class="btn ghost sm" id="register-tpl">登记模板</button>
        </div>
      </div>
      <div class="toolbar">
        <span class="mini">来源文件</span>
        <select id="src-file" class="pselect"></select>
        <span class="mini">来源工作表</span>
        <select id="src-sheet" class="pselect"></select>
        <span class="spring"></span>
        <span class="mini">目标模板</span>
        <select id="tpl" class="pselect"></select>
      </div>
    </section>

    <div class="grid2">
      <section class="panel">
        <div class="subhead"><h3>来源字段</h3><span class="mini" id="src-count"></span></div>
        <div class="fields-box" id="src-fields"><div class="loading">读取中…</div></div>
      </section>
      <section class="panel">
        <div class="subhead"><h3>模板字段（目标列）</h3><span class="mini" id="tpl-count"></span></div>
        <div class="fields-box" id="tpl-fields"><div class="loading">读取中…</div></div>
      </section>
    </div>

    <section class="panel">
      <div class="subhead">
        <h3>自动推荐</h3>
        <div class="toolbar">
          <button class="btn sm ghost" id="suggest">自动推荐</button>
          <button class="btn sm" id="apply-all" disabled>全部应用为规则</button>
        </div>
      </div>
      <div id="suggest-box" class="muted">尚未生成推荐。点击「自动推荐」按字段名称、类型与别名词典匹配（带置信度）。</div>
    </section>

    <section class="panel">
      <div class="subhead">
        <h3>映射矩阵</h3>
        <div class="toolbar">
          <button class="btn sm ghost" id="add-cell">+ 字段映射（列→列）</button>
          <button class="btn sm ghost" id="add-transpose">+ 区域转置（行→列）</button>
          <button class="btn sm ghost" id="clear-rules">清空规则</button>
        </div>
      </div>
      <div class="hint" style="margin-top:0">
        每行一条规则：源行/源列 → 目标行/目标列，行列均为 <b>1 基</b>。选择器：<b>fixed</b> 固定坐标、<b>record</b> 当前记录行+偏移（数据逐条下移）、<b>header</b> 按字段名、<b>span</b> 区域（起始 + 数量，连续 N 个）。
        典型转置（A 表某行 → B 表某列）：模式 transpose，源行 fixed=3、源列 span 起始=2 数量=32 → 目标行 fixed=4、目标列 fixed=6；预演会逐格显示「来源 A3..AF3 → 目标 F4..F35」。
      </div>
      <div id="rules" class="rules"></div>
      <div id="rules-empty" class="empty-state">还没有映射规则。可用上方「自动推荐」一键生成，或手动新增。</div>
    </section>

    <section class="panel">
      <div class="subhead">
        <h3>预演与保存</h3>
        <div class="toolbar">
          <input id="preview-limit" class="pinput" type="number" min="1" max="2000" value="200" style="width:92px" title="预演记录数">
          <button class="btn" id="preview">预演（逐单元格）</button>
          <button class="btn ghost" id="save">保存方案</button>
        </div>
      </div>
      <div class="toolbar" style="margin-bottom:8px">
        <input id="profile-name" class="pinput" placeholder="方案名称（保存时必填，如：扣分工单 → 台账）" style="min-width:240px">
        <input id="profile-notes" class="pinput grow" placeholder="备注（可选）">
      </div>
      <div id="preview-box"><div class="empty-state">点击「预演」按当前规则逐单元格计算目标坐标与值（不会写任何文件）。</div></div>
    </section>

    <section class="panel">
      <div class="subhead"><h3>已保存方案</h3><button class="btn sm ghost" id="reload-profiles">刷新</button></div>
      <div id="profiles"><div class="loading">读取中…</div></div>
    </section>
  </div>`;

  const srcFileEl = host.querySelector('#src-file');
  const srcSheetEl = host.querySelector('#src-sheet');
  const tplEl = host.querySelector('#tpl');
  const srcFieldsEl = host.querySelector('#src-fields');
  const tplFieldsEl = host.querySelector('#tpl-fields');
  const srcCountEl = host.querySelector('#src-count');
  const tplCountEl = host.querySelector('#tpl-count');
  const suggestBoxEl = host.querySelector('#suggest-box');
  const rulesEl = host.querySelector('#rules');
  const rulesEmptyEl = host.querySelector('#rules-empty');
  const previewBoxEl = host.querySelector('#preview-box');
  const profilesEl = host.querySelector('#profiles');
  const nameEl = host.querySelector('#profile-name');
  const notesEl = host.querySelector('#profile-notes');
  const previewLimitEl = host.querySelector('#preview-limit');
  const applyAllBtn = host.querySelector('#apply-all');
  const suggestBtn = host.querySelector('#suggest');

  const currentTemplate = () => st.templates.find((t) => String(t.id) === String(st.templateId)) || null;

  // ------------------------------------------------------------ 数据加载
  async function loadBase() {
    const res = await Promise.all([api.get('/files'), api.get('/templates'), api.get('/profiles')]);
    if (!host.isConnected) return;
    const f = res[0].data || {};
    const t = res[1].data || {};
    const p = res[2].data || {};
    st.files = Array.isArray(f.files) ? f.files : [];
    st.templates = Array.isArray(t.templates) ? t.templates : [];
    st.profiles = Array.isArray(p.profiles) ? p.profiles : [];

    srcFileEl.innerHTML = st.files.length
      ? st.files.map((x) => `<option value="${esc(x.id)}" title="${esc(x.path)}">${esc(x.name || basename(x.path))}（${esc(x.fmt || '?')}）</option>`).join('')
      : '<option value="">（工作区没有文件）</option>';
    if (st.fileId && st.files.some((x) => String(x.id) === String(st.fileId))) srcFileEl.value = st.fileId;
    else st.fileId = (st.files[0] || {}).id || '';

    tplEl.innerHTML = st.templates.length
      ? st.templates.map((x) => `<option value="${esc(x.id)}" title="${esc(x.filePath || '')}">${esc(x.name || x.id)}（${(x.fields || []).length} 字段）</option>`).join('')
      : '<option value="">（还没有登记模板）</option>';
    if (st.templateId && st.templates.some((x) => String(x.id) === String(st.templateId))) tplEl.value = st.templateId;
    else st.templateId = (st.templates[0] || {}).id || '';

    await loadSheets();
    renderTemplateFields();
    await loadProfiles();
  }

  async function loadSheets() {
    st.sheets = [];
    st.sheetIndex = 0;
    srcSheetEl.innerHTML = '<option value="">读取中…</option>';
    if (!st.fileId) {
      srcSheetEl.innerHTML = '<option value="">（无文件）</option>';
      st.sourceFields = [];
      renderSourceFields();
      return;
    }
    try {
      const res = await api.get(`/files/${encodeURIComponent(st.fileId)}/sheets`);
      if (!host.isConnected) return;
      const d = res.data || {};
      st.sheets = Array.isArray(d.sheets) ? d.sheets : [];
      srcSheetEl.innerHTML = st.sheets.length
        ? st.sheets.map((s, i) => (
          `<option value="${i}" ${i === 0 ? 'selected' : ''}>${esc(s.name || ('Sheet' + (i + 1)))}（${num(s.nRows)} 行 × ${num(s.nCols)} 列）</option>`
        )).join('')
        : '<option value="">（该文件没有工作表）</option>';
      st.sheetIndex = 0;
      await loadSourceFields();
    } catch (e) {
      srcSheetEl.innerHTML = '<option value="">工作表读取失败</option>';
      srcFieldsEl.innerHTML = `<div class="error-state">工作表读取失败：${esc(e.message || e)}</div>`;
    }
  }

  async function loadSourceFields() {
    srcFieldsEl.innerHTML = '<div class="loading">推断来源字段…</div>';
    st.sourceFields = [];
    const sheet = st.sheets[st.sheetIndex];
    if (!sheet) { renderSourceFields(); return; }
    // 1) 服务端直接给了字段列表
    if (Array.isArray(sheet.fields) && sheet.fields.length) {
      st.sourceFields = sheet.fields.map((x, i) => ({
        name: x.name || x.fieldName || `字段${i + 1}`,
        index: num(x.index ?? i),
        colLetter: colLetter(num(x.index ?? i) + 1),
        kindHint: x.kindHint || 'text',
      }));
      renderSourceFields();
      return;
    }
    // 2) 从预览的前几行推断表头（多级表头用 " / " 连接）
    try {
      const res = await api.get(`/files/${encodeURIComponent(st.fileId)}/sheet/${st.sheetIndex}/preview`, { rows: 6 });
      if (!host.isConnected) return;
      const d = res.data || {};
      st.sourceFields = deriveFieldsFromPreview(d, sheet);
    } catch (e) {
      srcFieldsEl.innerHTML = `<div class="error-state">来源字段推断失败（可手动在规则中填写字段名）：${esc(e.message || e)}</div>`;
      return;
    }
    renderSourceFields();
  }

  function deriveFieldsFromPreview(preview, sheet) {
    const rows = Array.isArray(preview.rows) ? preview.rows : [];
    if (!rows.length) return [];
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
      if (vals.length < Math.max(2, Math.floor(maxCount * 0.3))) continue; // 跳过合并大标题
      const isData = vals.some(([, v]) => typeof v === 'number' || /^-?\d+([.,]\d+)?$/.test(String(v).trim()));
      if (isData) break;
      chosen.push(row);
    }
    if (!chosen.length) return [];
    const keys = [...new Set(chosen.flatMap((r) => Object.keys(r.cells || {})))]
      .filter((k) => /^[A-Za-z]+$/.test(k))
      .sort((a, b) => letterIndex(a) - letterIndex(b));
    return keys.map((letter) => {
      const parts = chosen
        .map((r) => r.cells[letter])
        .filter((v) => v != null && String(v).trim() !== '')
        .map((v) => String(v).trim());
      return {
        name: parts.join(' / ') || `第 ${letterIndex(letter)} 列`,
        index: letterIndex(letter) - 1,
        colLetter: letter.toUpperCase(),
        kindHint: 'text',
      };
    });
  }

  function renderSourceFields() {
    srcCountEl.textContent = st.sourceFields.length ? `共 ${st.sourceFields.length} 个字段` : '';
    if (!st.sourceFields.length) {
      srcFieldsEl.innerHTML = '<div class="empty-state">未识别到来源字段：请在下方映射矩阵里直接填写坐标（fixed）或字段名（header）。</div>';
      return;
    }
    srcFieldsEl.innerHTML = st.sourceFields.map((f) => `
      <div class="field-item">
        <span class="fname2">${esc(f.name)}${f.kindHint && f.kindHint !== 'text' ? ` <span class="kindtag">${esc(f.kindHint)}</span>` : ''}</span>
        <span class="coltag">${esc(f.colLetter || colLetter(f.index + 1))} 列（第 ${f.index + 1} 列）</span>
      </div>`).join('');
  }

  function renderTemplateFields() {
    const tpl = currentTemplate();
    tplCountEl.textContent = tpl ? `共 ${(tpl.fields || []).length} 个字段` : '';
    if (!tpl) {
      tplFieldsEl.innerHTML = '<div class="empty-state">还没有登记模板。可点击右上角「登记模板」登记一个 xlsx，或让管理员在 data/templates 里登记。</div>';
      return;
    }
    const fields = Array.isArray(tpl.fields) ? tpl.fields : [];
    if (!fields.length) {
      tplFieldsEl.innerHTML = '<div class="empty-state">该模板没有解析出字段列表。</div>';
      return;
    }
    tplFieldsEl.innerHTML = `
      <div class="field-item"><span class="fname2">模板：${esc(tpl.name || tpl.id)}</span>
        <span class="coltag">数据起始：第 ${esc((tpl.dataStart || [])[0] ?? '?')} 行第 ${esc((tpl.dataStart || [])[1] ?? '?')} 列</span></div>` +
      fields.map((f) => {
        const col = num(f.index) + 1;
        return `<div class="field-item">
          <span class="fname2">${esc(f.name || '')}${f.kindHint && f.kindHint !== 'text' ? ` <span class="kindtag">${esc(f.kindHint)}</span>` : ''}</span>
          <span class="coltag">${colLetter(col)} 列（第 ${col} 列）</span>
        </div>`;
      }).join('');
  }

  async function loadProfiles() {
    try {
      const res = await api.get('/profiles');
      if (!host.isConnected) return;
      const d = res.data || {};
      st.profiles = Array.isArray(d.profiles) ? d.profiles : [];
      renderProfiles();
    } catch (e) {
      profilesEl.innerHTML = `<div class="error-state">方案读取失败：${esc(e.message || e)}</div>`;
    }
  }

  function renderProfiles() {
    if (!st.profiles.length) {
      profilesEl.innerHTML = '<div class="empty-state">还没有保存的映射方案。配置好规则后点「保存方案」即可复用。</div>';
      return;
    }
    profilesEl.innerHTML = `<div class="tablewrap"><table class="grid">
      <thead><tr><th>方案名称</th><th>模板</th><th class="num">规则数</th><th>备注</th><th>操作</th></tr></thead>
      <tbody>${st.profiles.map((p) => `
        <tr>
          <td>${esc(p.name || p.id)}</td>
          <td>${esc((st.templates.find((t) => String(t.id) === String(p.templateId)) || {}).name || p.templateId || '')}</td>
          <td class="num">${(p.rules || []).length}</td>
          <td>${esc(p.notes || '')}</td>
          <td class="nowrap">
            <button class="linkbtn" data-load="${esc(p.id)}">载入</button>
            <button class="linkbtn danger" data-del="${esc(p.id)}">删除</button>
          </td>
        </tr>`).join('')}</tbody>
    </table></div>`;
  }

  // ------------------------------------------------------------ 推荐
  async function doSuggest() {
    if (!st.fileId || !st.templateId) { toast('请先选择来源文件与目标模板', 'err'); return; }
    suggestBtn.disabled = true;
    suggestBtn.textContent = '推荐中…';
    suggestBoxEl.innerHTML = '<div class="loading">正在计算字段匹配…</div>';
    try {
      const res = await api.post('/mapping/suggest', {
        fileId: st.fileId, sheetIndex: st.sheetIndex, templateId: st.templateId,
      });
      if (!host.isConnected) return;
      const d = res.data || {};
      st.suggestions = Array.isArray(d.suggestions) ? d.suggestions
        : Array.isArray(d.matches) ? d.matches : [];
      st.suggestMeta = d;
      st.sourceSignature = d.sourceSignature || '';
      // 服务端返回的真实来源字段（TableModel.fields）优先于前端预览推断
      if (Array.isArray(d.sourceFields) && d.sourceFields.length) {
        st.sourceFields = d.sourceFields.map((f, i) => ({
          name: f.name || `字段${i + 1}`,
          index: num(f.index ?? i),
          colLetter: colLetter(num(f.index ?? i) + 1),
          kindHint: f.kindHint || 'text',
        }));
        renderSourceFields();
      }
      renderSuggestions();
    } catch (e) {
      suggestBoxEl.innerHTML = `<div class="error-state">自动推荐失败：${esc(e.message || e)}</div>`;
    } finally {
      if (host.isConnected) {
        suggestBtn.disabled = false;
        suggestBtn.textContent = '自动推荐';
      }
    }
  }

  function renderSuggestions() {
    applyAllBtn.disabled = !st.suggestions.length;
    const meta = st.suggestMeta || {};
    const metaLine = [
      meta.sourceSheet ? `来源工作表：${meta.sourceSheet}` : '',
      num(meta.recordCount) ? `数据记录 ${num(meta.recordCount)} 行` : '',
      meta.sourceSignature ? `来源指纹：${meta.sourceSignature}` : '',
      meta.matchedProfile ? `已匹配方案：${meta.matchedProfile.name || meta.matchedProfile.id}` : '',
    ].filter(Boolean).join(' · ');
    const metaHtml = metaLine ? `<div class="mini" style="margin:4px 0 6px">${esc(metaLine)}</div>` : '';
    if (!st.suggestions.length) {
      suggestBoxEl.innerHTML = `${metaHtml}<div class="empty-state">没有推荐结果：可手动在映射矩阵中添加规则。</div>`;
      return;
    }
    suggestBoxEl.innerHTML = `${metaHtml}<div class="tablewrap" style="max-height:240px"><table class="grid">
      <thead><tr><th>目标字段（模板）</th><th></th><th>来源字段</th><th class="num">置信度</th><th>依据</th><th>操作</th></tr></thead>
      <tbody>${st.suggestions.map((s, i) => {
        const conf = num(s.confidence);
        const cls = conf >= 0.7 ? '' : conf >= 0.4 ? 'mid' : 'low';
        return `<tr>
          <td>${esc(s.targetField || '')}</td>
          <td class="arrow">→</td>
          <td>${esc(s.sourceField || '')}</td>
          <td class="num"><span class="conf ${cls}">${(conf * 100).toFixed(0)}%</span></td>
          <td class="mini">${esc(s.reason || '')}</td>
          <td class="nowrap"><button class="linkbtn" data-apply="${i}">应用</button></td>
        </tr>`;
      }).join('')}</tbody>
    </table></div>`;
  }

  function ruleFromSuggestion(s) {
    const tpl = currentTemplate();
    const tplField = ((tpl && tpl.fields) || []).find((f) => String(f.name) === String(s.targetField));
    const srcField = st.sourceFields.find((f) => String(f.name) === String(s.sourceField));
    const targetCol = tplField
      ? { kind: 'fixed', value: num(tplField.index) + 1 }
      : { kind: 'header', value: s.targetField || '' };
    const sourceCol = srcField
      ? { kind: 'header', value: srcField.name }
      : { kind: 'header', value: s.sourceField || '' };
    return {
      source: { sheet: null, row: { kind: 'record', value: 0 }, col: sourceCol },
      target: { sheet: null, row: { kind: 'record', value: 0 }, col: targetCol },
      mode: 'cell', transform: 'identity', onMissing: 'blank', constant: null,
      note: `推荐：${s.sourceField || ''} → ${s.targetField || ''}（${((num(s.confidence)) * 100).toFixed(0)}%）`,
    };
  }

  function applySuggestion(i) {
    const s = st.suggestions[i];
    if (!s) return;
    st.rules.push(ruleFromSuggestion(s));
    renderRules();
    toast('已加入 1 条映射规则', 'ok');
  }

  // ------------------------------------------------------------ 映射矩阵
  function blankRule() {
    return {
      source: { sheet: null, row: { kind: 'record', value: 0 }, col: { kind: 'fixed', value: 1 } },
      target: { sheet: null, row: { kind: 'record', value: 0 }, col: { kind: 'fixed', value: 1 } },
      mode: 'cell', transform: 'identity', onMissing: 'blank', constant: null, note: '',
    };
  }

  function spanParts(sel) {
    const v = sel && sel.value;
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      return { start: v.start ?? v.from ?? '', count: v.count ?? v.n ?? '' };
    }
    if (Array.isArray(v) && v.length === 2) return { start: v[0], count: v[1] };
    return { start: '', count: v ?? '' };
  }

  function selInput(label, i, path, sel, kinds) {
    const kind = (sel && sel.kind) || 'fixed';
    const isSpan = kind === 'span';
    const parts = isSpan ? spanParts(sel) : null;
    const val = sel && sel.value != null && !isSpan ? sel.value : '';
    const valueHtml = isSpan
      ? `<span class="selpair">
           <input class="pinput" data-i="${i}" data-k="${path}.value.start" value="${esc(parts.start)}" placeholder="起始">
           <input class="pinput" data-i="${i}" data-k="${path}.value.count" value="${esc(parts.count)}" placeholder="数量">
         </span>`
      : `<input class="pinput" data-i="${i}" data-k="${path}.value" value="${esc(val)}" placeholder="${kind === 'header' ? '字段名' : '1 基坐标/偏移'}">`;
    return `<label>${esc(label)}
      <span class="selpair">
        <select class="pselect" data-i="${i}" data-k="${path}.kind">${opts(kinds, kind)}</select>
        ${valueHtml}
      </span>
    </label>`;
  }

  function ruleHtml(r, i) {
    return `<div class="rule-item" data-i="${i}">
      <div class="rule-head">
        <span>规则 ${i + 1}${r.note ? ` · ${esc(r.note)}` : ''}</span>
        <span>
          <button class="linkbtn" data-up="${i}">上移</button>
          <button class="linkbtn" data-down="${i}">下移</button>
          <button class="linkbtn danger" data-del="${i}">删除</button>
        </span>
      </div>
      <div class="rule-grid">
        ${selInput('源行', i, 'source.row', r.source.row, SEL_KINDS)}
        ${selInput('源列', i, 'source.col', r.source.col, SEL_KINDS)}
        ${selInput('目标行', i, 'target.row', r.target.row, SEL_KINDS)}
        ${selInput('目标列', i, 'target.col', r.target.col, SEL_KINDS)}
        <label>模式<select class="pselect" data-i="${i}" data-k="mode">${opts(MODES, r.mode)}</select></label>
        <label>转换<select class="pselect" data-i="${i}" data-k="transform">${opts(TRANSFORMS, r.transform)}</select></label>
        <label>缺失策略<select class="pselect" data-i="${i}" data-k="onMissing">${opts(ON_MISSING, r.onMissing)}</select></label>
        <label>常量<input class="pinput" data-i="${i}" data-k="constant" value="${esc(r.constant ?? '')}" placeholder="constant 时使用"></label>
        <label>备注<input class="pinput" data-i="${i}" data-k="note" value="${esc(r.note || '')}" placeholder="可选"></label>
      </div>
    </div>`;
  }

  function renderRules() {
    rulesEmptyEl.hidden = st.rules.length > 0;
    rulesEl.innerHTML = st.rules.map(ruleHtml).join('');
  }

  function onRuleChange(e) {
    const el = e.target.closest('[data-i][data-k]');
    if (!el) return;
    const i = num(el.dataset.i);
    const key = el.dataset.k;
    const rule = st.rules[i];
    if (!rule) return;
    const raw = el.value;
    if (key.endsWith('.kind')) {
      const base = key.replace(/\.kind$/, '');
      const old = getPath(rule, base);
      if (raw === 'span') {
        if (!old || typeof old !== 'object' || Array.isArray(old)) {
          setPath(rule, base, { start: null, count: old == null || old === '' ? 1 : old });
        }
      } else if (old && typeof old === 'object' && !Array.isArray(old)) {
        const sp = spanParts({ value: old });
        setPath(rule, base, sp.start != null && sp.start !== '' ? sp.start : (sp.count === '' ? 1 : sp.count));
      }
      setPath(rule, key, raw);
    } else if (key.endsWith('.value.start') || key.endsWith('.value.count')) {
      const base = key.replace(/\.(start|count)$/, '');
      const cur = getPath(rule, base);
      const obj = cur && typeof cur === 'object' && !Array.isArray(cur) ? { ...cur } : {};
      const n = raw === '' ? null : (Number.isFinite(Number(raw)) ? Number(raw) : raw);
      obj[key.endsWith('.start') ? 'start' : 'count'] = n;
      setPath(rule, base, obj);
    } else if (key === 'mode' || key === 'transform' || key === 'onMissing') {
      setPath(rule, key, raw);
    } else if (key === 'note') {
      rule.note = raw;
    } else if (key === 'constant') {
      rule.constant = raw === '' ? null : raw;
    } else {
      const kindPath = key.replace(/\.value$/, '.kind');
      const kind = getPath(rule, kindPath) || '';
      if (kind === 'header') setPath(rule, key, raw);
      else if (raw === '') setPath(rule, key, 0);
      else setPath(rule, key, Number.isFinite(Number(raw)) ? Number(raw) : raw);
    }
    renderRules();
  }

  function onRuleClick(e) {
    const btn = e.target.closest('button');
    if (!btn) return;
    if (btn.dataset.del != null) {
      st.rules.splice(num(btn.dataset.del), 1);
      renderRules();
      return;
    }
    if (btn.dataset.up != null || btn.dataset.down != null) {
      const i = num(btn.dataset.up != null ? btn.dataset.up : btn.dataset.down);
      const j = btn.dataset.up != null ? i - 1 : i + 1;
      if (j < 0 || j >= st.rules.length) return;
      const tmp = st.rules[i];
      st.rules[i] = st.rules[j];
      st.rules[j] = tmp;
      renderRules();
    }
  }

  // ------------------------------------------------------------ 预演
  async function doPreview() {
    if (!st.fileId || !st.templateId) { toast('请先选择来源文件与目标模板', 'err'); return; }
    if (!st.rules.length) { toast('请先添加至少一条映射规则', 'err'); return; }
    const btn = host.querySelector('#preview');
    btn.disabled = true;
    btn.textContent = '预演中…';
    previewBoxEl.innerHTML = '<div class="loading">正在按规则逐单元格计算…</div>';
    try {
      const res = await api.post('/mapping/preview', {
        fileId: st.fileId,
        sheetIndex: st.sheetIndex,
        templateId: st.templateId,
        profileId: st.loadedProfileId || null,
        rules: st.rules.map(clone),
        limit: Math.max(1, Math.min(2000, num(previewLimitEl.value) || 200)),
      });
      if (!host.isConnected) return;
      st.preview = res.data || {};
      st.previewOffset = 0;
      renderPreview();
    } catch (e) {
      previewBoxEl.innerHTML = `<div class="error-state">预演失败：${esc(e.message || e)}</div>`;
    } finally {
      if (host.isConnected) {
        btn.disabled = false;
        btn.textContent = '预演（逐单元格）';
      }
    }
  }

  function previewRef(ref) {
    if (!ref) return '<span class="muted">（常量）</span>';
    return `<div class="mono wrapdot">${esc(ref.display || `${ref.a1 || ''}`)}</div>
      <div class="mini">${esc(basename(ref.filePath || ''))} · ${esc(ref.sheet || '')}</div>`;
  }

  function renderPreview() {
    const d = st.preview || {};
    const cells = Array.isArray(d.cells) ? d.cells : [];
    const warnings = Array.isArray(d.warnings) ? d.warnings : [];
    // 预演响应里带了真实来源字段时，顺便刷新字段面板（比预览推断更准）
    if (Array.isArray(d.sourceFields) && d.sourceFields.length) {
      const names = d.sourceFields.map((f) => f.name).join('|');
      if (names !== st.sourceFields.map((f) => f.name).join('|')) {
        st.sourceFields = d.sourceFields.map((f, i) => ({
          name: f.name || `字段${i + 1}`,
          index: num(f.index ?? i),
          colLetter: colLetter(num(f.index ?? i) + 1),
          kindHint: f.kindHint || 'text',
        }));
        renderSourceFields();
      }
    }
    if (d.sourceSignature) st.sourceSignature = d.sourceSignature;
    if (!cells.length && !warnings.length) {
      previewBoxEl.innerHTML = '<div class="empty-state">没有可预演的数据：请确认来源表有数据行，且规则包含 record 行选择器。</div>';
      return;
    }
    const pageSize = 50;
    const total = cells.length;
    const pages = Math.max(1, Math.ceil(total / pageSize));
    const page = Math.min(pages, Math.floor(st.previewOffset / pageSize) + 1);
    const slice = cells.slice((page - 1) * pageSize, page * pageSize);
    previewBoxEl.innerHTML = `
      ${warnings.length ? `<div class="warn-state">${warnings.map((w) => esc(typeof w === 'string' ? w : JSON.stringify(w))).join('<br>')}</div>` : ''}
      <div class="toolbar between" style="margin:8px 0">
        <div class="muted">共 ${total} 个目标单元格${d.truncated ? '（已截断，仅预演前 N 条记录）' : ''}，逐格展示「来源坐标 → 目标坐标 → 值」</div>
        <div class="muted">第 ${page} / ${pages} 页</div>
      </div>
      <div class="tablewrap" style="max-height:46vh">
        <table class="grid">
          <thead><tr>
            <th class="num">规则</th><th>来源坐标（来源文件 · 工作表 · 行 · 列）</th><th></th>
            <th>目标坐标（模板文件 · 工作表 · 行 · 列）</th><th>值</th>
          </tr></thead>
          <tbody>${slice.map((c) => `
            <tr>
              <td class="num">${num(c.ruleIndex) + 1}</td>
              <td>${previewRef(c.sourceRef)}</td>
              <td class="arrow">→</td>
              <td>${previewRef(c.target)}</td>
              <td class="hitv">${c.value == null || c.value === '' ? '<span class="muted">（空）</span>' : esc(c.value)}</td>
            </tr>`).join('')}</tbody>
        </table>
      </div>
      ${pages > 1 ? `<div class="pager">
        <button class="pgbtn" data-pv-page="${page - 1}" ${page <= 1 ? 'disabled' : ''}>‹</button>
        <span class="pginfo">第 ${page} / ${pages} 页 · 每页 ${pageSize} 格</span>
        <button class="pgbtn" data-pv-page="${page + 1}" ${page >= pages ? 'disabled' : ''}>›</button>
      </div>` : ''}`;
    previewBoxEl.querySelectorAll('[data-pv-page]').forEach((b) => {
      b.addEventListener('click', () => {
        st.previewOffset = (num(b.dataset.pvPage) - 1) * pageSize;
        renderPreview();
      });
    });
  }

  // ------------------------------------------------------------ 保存方案
  async function saveProfile() {
    if (!st.templateId) { toast('请先选择目标模板', 'err'); return; }
    if (!st.rules.length) { toast('请先添加至少一条映射规则', 'err'); return; }
    let name = (nameEl.value || '').trim();
    if (!name) {
      const d = new Date();
      const p = (x) => String(x).padStart(2, '0');
      name = `映射方案_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
      nameEl.value = name;
    }
    const btn = host.querySelector('#save');
    btn.disabled = true;
    btn.textContent = '保存中…';
    try {
      const signature = st.sourceSignature || st.sourceFields.slice(0, 4).map((f) => f.name).join('|');
      const res = await api.post('/profiles', {
        name,
        templateId: st.templateId,
        sourceSignature: signature,
        rules: st.rules.map(clone),
        notes: notesEl.value || '',
      });
      if (!host.isConnected) return;
      const p = (res.data || {}).profile || {};
      st.loadedProfileId = p.id || st.loadedProfileId;
      toast(`方案已保存${p.id ? `（${p.id}）` : ''}，可在「转换」页选用`, 'ok');
      await loadProfiles();
    } catch (e) {
      toast(`保存失败：${e.message || e}`, 'err');
    } finally {
      if (host.isConnected) {
        btn.disabled = false;
        btn.textContent = '保存方案';
      }
    }
  }

  function loadProfileIntoEditor(pid) {
    const p = st.profiles.find((x) => String(x.id) === String(pid));
    if (!p) return;
    st.rules = clone(Array.isArray(p.rules) ? p.rules : []).map((r) => ({
      source: r.source || { sheet: null, row: { kind: 'record', value: 0 }, col: { kind: 'fixed', value: 1 } },
      target: r.target || { sheet: null, row: { kind: 'record', value: 0 }, col: { kind: 'fixed', value: 1 } },
      mode: r.mode || 'cell',
      transform: r.transform || 'identity',
      onMissing: r.onMissing || r.on_missing || 'blank',
      constant: r.constant ?? null,
      note: r.note || '',
    }));
    st.loadedProfileId = p.id || '';
    nameEl.value = p.name || '';
    notesEl.value = p.notes || '';
    if (p.templateId && st.templates.some((t) => String(t.id) === String(p.templateId))) {
      st.templateId = p.templateId;
      tplEl.value = p.templateId;
      renderTemplateFields();
    }
    renderRules();
    toast(`已载入方案：${p.name || p.id}（${st.rules.length} 条规则）`, 'ok');
  }

  // ------------------------------------------------------------ 登记模板弹窗
  function openRegisterTemplate() {
    const mask = document.createElement('div');
    mask.className = 'pg-modal-mask';
    mask.innerHTML = `
      <div class="pg-modal" role="dialog" aria-modal="true">
        <div class="modal-head"><h3>登记模板</h3><button class="modal-x" data-close>×</button></div>
        <div class="modal-body">
          <div class="hint" style="margin-top:0">登记一个 xlsx 作为模板，后端会解析表头与字段；「数据起始单元格」由模板或用户在预览中框选，这里可一并传入（1 基）。</div>
          <div class="formline"><label>文件路径</label><input id="t-path" class="pinput mono" placeholder="C:\\...\\项目台账统计表.xlsx"></div>
          <div class="formline"><label>模板名称</label><input id="t-name" class="pinput" placeholder="如：项目台账统计表（2024）"></div>
          <div class="formline"><label>工作表</label><input id="t-sheet" class="pinput" placeholder="留空 = 第一个工作表"></div>
          <div class="formline"><label>数据起始</label>
            <input id="t-row" class="pinput" type="number" min="1" placeholder="行（1 基）" style="max-width:130px">
            <input id="t-col" class="pinput" type="number" min="1" placeholder="列（1 基）" style="max-width:130px">
          </div>
          <div id="t-err" class="error-state" hidden></div>
        </div>
        <div class="modal-foot"><button class="btn ghost" data-close>取消</button><button class="btn" id="t-submit">登记</button></div>
      </div>`;
    document.body.appendChild(mask);
    const close = () => mask.remove();
    mask.querySelectorAll('[data-close]').forEach((b) => b.addEventListener('click', close));
    mask.addEventListener('click', (e) => { if (e.target === mask) close(); });
    mask.querySelector('#t-submit').addEventListener('click', async () => {
      const path = mask.querySelector('#t-path').value.trim();
      const name = mask.querySelector('#t-name').value.trim();
      const sheet = mask.querySelector('#t-sheet').value.trim();
      const row = num(mask.querySelector('#t-row').value);
      const col = num(mask.querySelector('#t-col').value);
      const err = mask.querySelector('#t-err');
      if (!path) {
        err.hidden = false;
        err.textContent = '请填写模板文件路径。';
        return;
      }
      const body = { filePath: path, path, name: name || basename(path), sheet: sheet || null };
      if (row && col) body.dataStart = [row, col];
      try {
        await api.post('/templates', body);
        toast('模板已登记', 'ok');
        close();
        await loadBase();
      } catch (e) {
        err.hidden = false;
        err.textContent = `登记失败：${e.message || e}`;
      }
    });
  }

  // ------------------------------------------------------------ 事件绑定
  srcFileEl.addEventListener('change', async () => {
    st.fileId = srcFileEl.value;
    st.loadedProfileId = '';
    await loadSheets();
  });
  srcSheetEl.addEventListener('change', async () => {
    st.sheetIndex = num(srcSheetEl.value);
    await loadSourceFields();
  });
  tplEl.addEventListener('change', () => {
    st.templateId = tplEl.value;
    st.loadedProfileId = '';
    renderTemplateFields();
  });
  host.querySelector('#refresh-all').addEventListener('click', async () => {
    try { await loadBase(); toast('清单已刷新', 'ok'); } catch (e) { toast(`刷新失败：${e.message || e}`, 'err'); }
  });
  host.querySelector('#register-tpl').addEventListener('click', openRegisterTemplate);
  suggestBtn.addEventListener('click', doSuggest);
  applyAllBtn.addEventListener('click', () => {
    if (!st.suggestions.length) return;
    const existing = new Set(st.rules.map((r) => `${getPath(r, 'source.col.value')}→${getPath(r, 'target.col.value')}`));
    let added = 0;
    for (const s of st.suggestions) {
      const r = ruleFromSuggestion(s);
      const key = `${getPath(r, 'source.col.value')}→${getPath(r, 'target.col.value')}`;
      if (existing.has(key)) continue;
      existing.add(key);
      st.rules.push(r);
      added++;
    }
    renderRules();
    toast(added ? `已应用 ${added} 条推荐规则` : '推荐规则都已存在', added ? 'ok' : 'info');
  });
  host.querySelector('#add-cell').addEventListener('click', () => {
    st.rules.push(blankRule());
    renderRules();
  });
  host.querySelector('#add-transpose').addEventListener('click', () => {
    st.rules.push({
      source: { sheet: null, row: { kind: 'fixed', value: 3 }, col: { kind: 'span', value: { start: 2, count: 32 } } },
      target: { sheet: null, row: { kind: 'fixed', value: 4 }, col: { kind: 'fixed', value: 6 } },
      mode: 'transpose', transform: 'identity', onMissing: 'blank', constant: null,
      note: '区域转置：A 表第 3 行第 2~33 列（横向）→ B 表 F 列第 4~35 行（竖向）',
    });
    renderRules();
    toast('已添加转置规则示例：源行固定=3 / 源列 span 起始=2 数量=32 → 目标行固定=4 / 目标列固定=6，请按实际列号修改', 'info');
  });
  host.querySelector('#clear-rules').addEventListener('click', () => {
    if (!st.rules.length) return;
    if (!confirm('清空当前所有映射规则？')) return;
    st.rules = [];
    renderRules();
  });
  rulesEl.addEventListener('change', onRuleChange);
  rulesEl.addEventListener('click', onRuleClick);
  host.querySelector('#preview').addEventListener('click', doPreview);
  host.querySelector('#save').addEventListener('click', saveProfile);
  host.querySelector('#reload-profiles').addEventListener('click', loadProfiles);
  profilesEl.addEventListener('click', async (e) => {
    const loadBtn = e.target.closest('[data-load]');
    const delBtn = e.target.closest('[data-del]');
    if (loadBtn) { loadProfileIntoEditor(loadBtn.dataset.load); return; }
    if (delBtn) {
      if (!confirm('删除这个映射方案？')) return;
      try {
        await api.del(`/profiles/${encodeURIComponent(delBtn.dataset.del)}`);
        toast('方案已删除', 'ok');
        await loadProfiles();
      } catch (err) {
        toast(`删除失败：${err.message || err}`, 'err');
      }
    }
  });

  // ------------------------------------------------------------ 初始化
  try {
    await loadBase();
  } catch (e) {
    if (!host.isConnected) return;
    host.querySelector('#src-fields').innerHTML = `<div class="error-state">初始化失败：${esc(e.message || e)}</div>`;
    toast(`映射页加载失败：${e.message || e}`, 'err');
    return;
  }
  renderRules();
}
