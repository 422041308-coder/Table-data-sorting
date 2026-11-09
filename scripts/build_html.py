# -*- coding: utf-8 -*-
"""把 docs/*.md 渲染成一个自包含的 HTML 架构说明书（零依赖，只用于交付阅读）。

用法：
    python scripts/build_html.py docs/01-xxx.md docs/02-xxx.md -o docs/项目架构设计.html
"""
import html
import re
import sys
import os

# ---------------------------------------------------------------- CSS
CSS = r"""
:root{
  --ink:#1f2937; --muted:#64748b; --line:#e2e8f0; --accent:#4f46e5;
  --accent2:#0ea5e9; --warn:#b45309; --bg:#f6f8fb; --card:#ffffff;
  --code-bg:#f1f5f9; --th:#eef2f7;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{
  margin:0; background:var(--bg); color:var(--ink);
  font:15px/1.75 "Microsoft YaHei","PingFang SC","Segoe UI",system-ui,-apple-system,sans-serif;
  -webkit-font-smoothing:antialiased;
}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}

/* layout */
#toc{
  position:fixed;left:0;top:0;bottom:0;width:288px;overflow:auto;
  background:#0f172a;color:#cbd5e1;padding:26px 20px 40px;
}
#toc .brand{font-size:16px;font-weight:700;color:#fff;letter-spacing:.5px}
#toc .brand small{display:block;font-size:11.5px;font-weight:400;color:#7c8aa5;margin-top:5px;letter-spacing:0}
#toc .grp{margin-top:22px}
#toc .grp>a{display:block;font-size:12px;font-weight:700;color:#93a4c4;text-transform:none;
  padding:4px 0 6px;border-bottom:1px solid #1e2a44;margin-bottom:6px}
#toc ul{list-style:none;margin:0;padding:0}
#toc li a{display:block;color:#b6c2d6;font-size:12.8px;line-height:1.55;
  padding:3px 8px;border-radius:6px;border-left:2px solid transparent}
#toc li a:hover{background:#1b2740;color:#fff;text-decoration:none}
#toc li.lv3 a{padding-left:20px;color:#8fa0bd;font-size:12.2px}
#toc .foot{margin-top:28px;font-size:11px;color:#5d6b85;line-height:1.7}

main{margin-left:288px;padding:34px 46px 90px;max-width:1180px}
.hero{
  background:linear-gradient(120deg,#eef2ff 0%,#e0f2fe 100%);
  border:1px solid #dbe4f0;border-radius:16px;padding:26px 30px;margin-bottom:30px;
}
.hero h1{margin:0 0 6px;font-size:25px;letter-spacing:.3px}
.hero p{margin:0;color:#475569;font-size:13.5px}
.kpis{display:flex;flex-wrap:wrap;gap:10px;margin-top:18px}
.kpi{background:#fff;border:1px solid #dbe4f0;border-radius:11px;padding:9px 14px;min-width:118px}
.kpi b{display:block;font-size:19px;color:var(--accent);line-height:1.25}
.kpi span{font-size:11.5px;color:var(--muted)}

h1.doc{font-size:22px;margin:46px 0 10px;padding:8px 0 12px;border-bottom:2px solid var(--accent);letter-spacing:.3px}
h2{font-size:19px;margin:44px 0 14px;padding-left:11px;border-left:4px solid var(--accent);letter-spacing:.2px}
h3{font-size:15.5px;margin:28px 0 10px;color:#334155;padding-left:9px;border-left:3px solid #c7d2fe}
h4{font-size:14px;margin:20px 0 8px;color:#334155}
p{margin:10px 0}
strong{color:#111827}

/* tables */
.tablewrap{overflow-x:auto;margin:14px 0 18px;border:1px solid var(--line);border-radius:10px;background:var(--card)}
table{border-collapse:collapse;width:100%;font-size:13.2px;line-height:1.62}
th,td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);padding:9px 11px;text-align:left;vertical-align:top;overflow-wrap:break-word}
th:last-child,td:last-child{border-right:none}
tr:last-child td{border-bottom:none}
th{background:var(--th);font-weight:700;color:#334155;white-space:nowrap}
tbody tr:nth-child(even){background:#fbfcfe}
td code,th code{white-space:nowrap}

/* code */
pre{background:var(--code-bg);border:1px solid var(--line);border-radius:10px;padding:15px 17px;
  overflow-x:auto;margin:14px 0;font-size:12.8px;line-height:1.72}
pre code{font-family:"Cascadia Mono",Consolas,"Courier New",monospace;white-space:pre;color:#1e293b}
p code,li code,td code,blockquote code{background:#eef2f7;border:1px solid #e2e8f0;border-radius:5px;
  padding:1px 5px;font-family:"Cascadia Mono",Consolas,monospace;font-size:12.3px;color:#0f172a}
blockquote{margin:14px 0;padding:11px 16px;background:#fffbeb;border-left:4px solid #f59e0b;border-radius:0 8px 8px 0;color:#78350f;font-size:13.6px}
blockquote p{margin:0}
ul,ol{margin:10px 0;padding-left:24px}
li{margin:5px 0}
hr{border:none;border-top:1px dashed #cbd5e1;margin:34px 0}
img{max-width:100%;height:auto;border:1px solid var(--line);border-radius:10px;margin:14px 0;display:block;box-shadow:0 2px 10px rgba(15,23,42,.06)}

/* architecture diagram */
.arch{margin:18px 0 8px;border:1px solid var(--line);border-radius:14px;background:var(--card);padding:18px 18px 12px}
.arch .layer{display:flex;gap:14px;align-items:stretch;border-radius:10px;padding:12px 15px;margin-bottom:0}
.arch .layer+.arrow{text-align:center;color:#94a3b8;font-size:15px;line-height:1.4;margin:2px 0}
.arch .lt{flex:0 0 132px;font-weight:700;font-size:13.5px;color:#fff;border-radius:8px;padding:5px 10px;display:flex;align-items:center;justify-content:center;text-align:center}
.arch .lc{flex:1;min-width:0}
.arch .lc .head{font-weight:700;font-size:13.5px;margin-bottom:6px;color:#1e293b}
.arch .chips{display:flex;flex-wrap:wrap;gap:6px}
.arch .chip{background:#fff;border:1px solid #dbe4f0;border-radius:7px;padding:3px 9px;font-size:11.8px;color:#334155;white-space:nowrap}
.arch .chip.key{border-color:#c7d2fe;background:#eef2ff;color:#3730a3;font-weight:700}
.l1 .lt{background:#6d28d9}.l1{background:#f5f3ff;border:1px solid #ede9fe}
.l2 .lt{background:#4f46e5}.l2{background:#eef2ff;border:1px solid #e0e7ff}
.l3 .lt{background:#0284c7}.l3{background:#eff6ff;border:1px solid #dbeafe}
.l4 .lt{background:#0d9488}.l4{background:#f0fdfa;border:1px solid #ccfbf1}
.l5 .lt{background:#b45309}.l5{background:#fffbeb;border:1px solid #fef3c7}
.arch .io{display:flex;gap:10px;flex-wrap:wrap;margin-top:12px;padding-top:12px;border-top:1px dashed #cbd5e1}
.arch .io .box{flex:1;min-width:230px;border:1px dashed #cbd5e1;border-radius:9px;padding:8px 12px;background:#f8fafc;font-size:12.4px;color:#475569}
.arch .io .box b{color:#1e293b}
.legend{font-size:11.8px;color:var(--muted);margin:6px 0 0}
.tag{display:inline-block;background:#eef2ff;color:#3730a3;border:1px solid #c7d2fe;border-radius:20px;
  font-size:11px;padding:1px 9px;margin-right:6px;vertical-align:1px}

@media (max-width:1100px){
  #toc{position:static;width:auto;height:auto}
  main{margin-left:0;padding:22px 18px 70px}
  .arch .layer{flex-direction:column}
  .arch .lt{flex:none}
}
@media print{
  #toc{display:none}
  main{margin:0;max-width:none;padding:0}
  body{background:#fff}
  .hero{background:#fff}
  h2{page-break-after:avoid}
  .tablewrap,pre,.arch{page-break-inside:avoid}
}
"""

ARCH = """
<div class="arch">
  <div class="layer l1">
    <div class="lt">表现层</div>
    <div class="lc">
      <div class="head">原生 HTML + CSS + ES Module 单页（无框架、无构建、无 CDN，离线可用）</div>
      <div class="chips">
        <span class="chip">检索页（文件/行/列 定位）</span>
        <span class="chip">字段映射页（矩阵 + 区域框选 + 预演）</span>
        <span class="chip">模板转换页</span>
        <span class="chip">双表对比页（左右坐标）</span>
        <span class="chip">设置页</span>
      </div>
    </div>
  </div>
  <div class="arrow">▼ &nbsp; HTTP + JSON &nbsp;(127.0.0.1:8765)</div>
  <div class="layer l2">
    <div class="lt">接口层</div>
    <div class="lc">
      <div class="head">app/server.py &nbsp;—&nbsp; 路由 / 静态托管 / 任务与下载令牌</div>
      <div class="chips">
        <span class="chip">/api/search</span><span class="chip">/api/cell</span><span class="chip">/api/templates</span>
        <span class="chip">/api/profiles</span><span class="chip">/api/mapping/preview</span><span class="chip">/api/convert</span>
        <span class="chip">/api/compare</span><span class="chip">/api/audit</span>
      </div>
    </div>
  </div>
  <div class="arrow">▼</div>
  <div class="layer l3">
    <div class="lt">服务层</div>
    <div class="lc">
      <div class="head">app/services/ &nbsp;—&nbsp; 业务编排（薄）</div>
      <div class="chips">
        <span class="chip">search_service</span><span class="chip">mapping_service</span>
        <span class="chip">convert_service</span><span class="chip">compare_service</span>
        <span class="chip">workspace_service</span><span class="chip">edit_service</span>
      </div>
    </div>
  </div>
  <div class="arrow">▼</div>
  <div class="layer l4">
    <div class="lt">领域层</div>
    <div class="lc">
      <div class="head">app/core/ &nbsp;—&nbsp; 纯逻辑，可单测</div>
      <div class="chips">
        <span class="chip key">model.py 契约</span>
        <span class="chip">detect.py 表头/数据区</span>
        <span class="chip">normalize.py 归一化</span>
        <span class="chip">index.py 倒排</span>
        <span class="chip">search.py 查询/打分</span>
        <span class="chip">mapping.py 规则引擎</span>
        <span class="chip">template.py 模板注入</span>
        <span class="chip">compare.py 差异</span>
        <span class="chip">cellref.py 坐标</span>
        <span class="chip">audit.py 备份/审计</span>
        <span class="chip">export.py 导出</span>
      </div>
    </div>
  </div>
  <div class="arrow">▼</div>
  <div class="layer l5">
    <div class="lt">读写层</div>
    <div class="lc">
      <div class="head">统一模型 WorkbookModel：读 → 只读解析；写 → 最小改动回写</div>
      <div class="chips">
        <span class="chip key">app/readers/</span>
        <span class="chip">sniff.py 魔数嗅探</span>
        <span class="chip">ooxml.py (xlsx/xlsm/et-zip)</span>
        <span class="chip">biff.py (xls/et-ole)</span>
        <span class="chip">csvr.py</span>
        <span class="chip key">app/writers/</span>
        <span class="chip">ooxml_writer.py 原地写回</span>
        <span class="chip">biff_writer.py（二期）</span>
      </div>
    </div>
  </div>
  <div class="io">
    <div class="box"><b>读路径：</b>目标文件夹（递归）→ 魔数识别 → 解析 → 表头/数据区识别 → 索引</div>
    <div class="box"><b>写路径：</b>冲突检测 → 备份到 data/backup/ → 定位节点 → 最小改动 → 重解析自检 → 审计 data/audit/</div>
  </div>
</div>
<p class="legend"><span class="tag">零依赖</span>仅 Python 3.10+ 标准库；<span class="tag">只读优先</span>源文件默认不改，写回必须备份+自检+可回滚；<span class="tag">契约冻结</span>model.py 是并行开发唯一同步点。</p>
"""


def esc(s):
    return html.escape(s, quote=False)


def inline(s):
    s = esc(s)
    # 图片：![alt](src) —— 放在链接之前处理
    s = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', r'<img src="\2" alt="\1">', s)
    s = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2">\1</a>', s)
    # 代码片段先切分，粗体只作用于代码之外，避免 `.../**` 这类 glob 被误判
    parts = re.split(r'(`[^`]+`)', s)
    out = []
    for p in parts:
        if len(p) >= 2 and p.startswith('`') and p.endswith('`'):
            out.append('<code>' + p[1:-1] + '</code>')
        else:
            out.append(re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', p))
    return ''.join(out)


def smart_join(parts):
    out = parts[0]
    for nxt in parts[1:]:
        a, b = out[-1:], nxt[:1]
        if a and b and (a.isascii() and a.isalnum()) and (b.isascii() and b.isalnum()):
            out += " " + nxt
        else:
            out += nxt
    return out


def is_special(line):
    s = line.strip()
    return (s.startswith('#') or s.startswith('```') or s.startswith('|')
            or s.startswith('>') or re.match(r'^[-*]\s+', s)
            or re.match(r'^\d+\.\s+', s) or s in ('---', '***'))


def split_cells(row):
    r = row.strip()
    if r.startswith('|'):
        r = r[1:]
    if r.endswith('|'):
        r = r[:-1]
    parts = re.split(r'(?<!\\)\|', r)
    return [p.strip().replace('\\|', '|') for p in parts]


def render_table(rows):
    header = split_cells(rows[0])
    body, aligns = [], []
    for r in rows[1:]:
        cells = split_cells(r)
        if cells and all(re.fullmatch(r':?-{2,}:?', c or '') for c in cells):
            aligns = cells
            continue
        body.append(cells)
    out = ['<div class="tablewrap"><table>', '<thead><tr>']
    for c in header:
        out.append(f'<th>{inline(c)}</th>')
    out.append('</tr></thead><tbody>')
    for cells in body:
        out.append('<tr>')
        for c in cells:
            out.append(f'<td>{inline(c)}</td>')
        out.append('</tr>')
    out.append('</tbody></table></div>')
    return '\n'.join(out)


def render_code(code, lang=''):
    if '表现层' in code and '接口层' in code and '读写层' in code:
        return ARCH
    return f'<pre><code>{esc(code)}</code></pre>'


def convert(md, doc_index=0):
    lines = md.splitlines()
    out, toc = [], []
    sec = [doc_index * 1000]
    i = 0
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s.startswith('```'):
            code = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code.append(lines[i])
                i += 1
            i += 1
            out.append(render_code('\n'.join(code)))
            continue
        if s.startswith('#'):
            level = len(s) - len(s.lstrip('#'))
            text = s[level:].strip()
            if level == 1:
                out.append(f'<h1 class="doc">{inline(text)}</h1>')
                i += 1
                continue
            sec[0] += 1
            sid = f'sec{sec[0]}'
            tag = f'h{min(level, 4)}'
            out.append(f'<{tag} id="{sid}">{inline(text)}</{tag}>')
            if level in (2, 3):
                toc.append((level, text, sid))
            i += 1
            continue
        if s.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append(lines[i].strip())
                i += 1
            out.append(render_table(rows))
            continue
        if re.match(r'^[-*]\s+', s):
            items = []
            while i < len(lines) and re.match(r'^[-*]\s+', lines[i].strip()):
                items.append(re.sub(r'^[-*]\s+', '', lines[i].strip()))
                i += 1
            out.append('<ul>' + ''.join(f'<li>{inline(x)}</li>' for x in items) + '</ul>')
            continue
        if re.match(r'^\d+\.\s+', s):
            items = []
            while i < len(lines) and re.match(r'^\d+\.\s+', lines[i].strip()):
                items.append(re.sub(r'^\d+\.\s+', '', lines[i].strip()))
                i += 1
            out.append('<ol>' + ''.join(f'<li>{inline(x)}</li>' for x in items) + '</ol>')
            continue
        if s.startswith('>'):
            qs = []
            while i < len(lines) and lines[i].strip().startswith('>'):
                qs.append(lines[i].strip().lstrip('>').strip())
                i += 1
            out.append('<blockquote><p>' + inline(smart_join(qs)) + '</p></blockquote>')
            continue
        if s in ('---', '***'):
            out.append('<hr>')
            i += 1
            continue
        para = [s]
        i += 1
        while i < len(lines) and lines[i].strip() and not is_special(lines[i]):
            para.append(lines[i].strip())
            i += 1
        out.append('<p>' + inline(smart_join(para)) + '</p>')
    return '\n'.join(out), toc


def build(md_paths, out_path, title, subtitle, no_kpi=False):
    docs, nav = [], []
    for k, p in enumerate(md_paths):
        with open(p, encoding='utf-8') as f:
            md = f.read()
        body, toc = convert(md, k)
        docs.append(body)
        first = next((l for l in md.splitlines() if l.startswith('# ')), '')
        nav.append((first[2:].strip() if first else os.path.basename(p), toc))

    toc_html = ['<nav id="toc">', f'<div class="brand">{esc(title)}<small>{esc(subtitle)}</small></div>']
    for name, toc in nav:
        toc_html.append('<div class="grp">')
        toc_html.append(f'<a href="#top">{esc(name)}</a><ul>')
        for level, text, sid in toc:
            cls = ' class="lv3"' if level == 3 else ''
            toc_html.append(f'<li{cls}><a href="#{sid}">{esc(text)}</a></li>')
        toc_html.append('</ul></div>')
    toc_html.append('<div class="foot">由 docs/*.md 自动生成<br>零依赖 · 可离线打开 · 可打印</div>')
    toc_html.append('</nav>')

    kpis = "" if no_kpi else """
    <div class="hero" id="top">
      <h1>表格检索 · 模板改造 · 数据核对 —— 需求与架构设计</h1>
      <p>面向"项目台账"场景的本地 Web 应用：全字段模糊检索、可替换模板的字段映射、双表差异核对、按单元格就地修改。</p>
      <div class="kpis">
        <div class="kpi"><b>66</b><span>实测表格文件</span></div>
        <div class="kpi"><b>25.3 万</b><span>单元格语料</span></div>
        <div class="kpi"><b>6</b><span>条需求（R1–R6）</span></div>
        <div class="kpi"><b>7</b><span>个开发切片</span></div>
        <div class="kpi"><b>250</b><span>项测试全绿</span></div>
        <div class="kpi"><b>0</b><span>第三方依赖</span></div>
        <div class="kpi"><b>18.8s</b><span>索引构建</span></div>
      </div>
    </div>"""
    if no_kpi:
        kpis = '<div id="top"></div>'

    html_doc = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(title)}</title>
<style>{CSS}</style>
</head>
<body>
{''.join(toc_html)}
<main>
{kpis}
{''.join(docs)}
</main>
</body>
</html>
"""
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html_doc)
    return out_path


if __name__ == '__main__':
    args = [a for a in sys.argv[1:]]
    out = 'docs/架构设计.html'
    title = '表格检索 / 模板改造 / 数据核对 —— 项目架构设计'
    subtitle = '需求整合稿 v1.0 · 2026-11-09'
    no_kpi = False
    rest = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == '-o' and i + 1 < len(args):
            out = args[i + 1]
            i += 2
            continue
        if a == '--title' and i + 1 < len(args):
            title = args[i + 1]
            i += 2
            continue
        if a == '--subtitle' and i + 1 < len(args):
            subtitle = args[i + 1]
            i += 2
            continue
        if a == '--no-kpi':
            no_kpi = True
            i += 1
            continue
        rest.append(a)
        i += 1
    p = build(rest, out, title, subtitle, no_kpi=no_kpi)
    print(f'OK {p} {os.path.getsize(p)} bytes')
