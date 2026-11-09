# -*- coding: utf-8 -*-
"""交付验收脚本：在真实语料上跑一遍验收清单，产出 docs/验收报告.md。

用法：
    python scripts/acceptance.py            # 全量
    python scripts/acceptance.py --no-index # 跳过索引（约 30s）
"""
from __future__ import annotations

import os
import shutil
import sys
import time
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

SAMPLE = os.path.join(ROOT, "samples", "项目台账统计表.xlsx")
SAMPLE_SHEET = "2024年项目台账（自筹-甲代-施工-日常经营）"
TMP = os.path.join(ROOT, "data", "tmp", "acceptance")
REPORT = os.path.join(ROOT, "docs", "验收报告.md")
CORPUS = os.environ.get("EXCEL_TOOLKIT_CORPUS", os.path.join(ROOT, "samples"))

results: list[dict] = []


def item(name: str, ok: bool, detail: str = "", cost: float = 0.0):
    results.append({"name": name, "ok": bool(ok), "detail": detail, "cost": cost})
    mark = "PASS" if ok else "FAIL"
    print(f"[{mark}] {name}  ({cost:.2f}s)  {detail}")


def fresh(name: str) -> str:
    os.makedirs(TMP, exist_ok=True)
    p = os.path.join(TMP, name)
    if os.path.exists(p):
        try:
            os.chmod(p, 0o666)
        except OSError:
            pass
        try:
            os.remove(p)
        except OSError:
            pass
    shutil.copy2(SAMPLE, p)
    os.chmod(p, 0o666)
    return p


def main(do_index: bool = True):
    # ------------------------------------------------------------------ 1. 读取与识别
    t0 = time.time()
    from app.core.detect import detect_workbook
    from app.core.readers import read_workbook
    from app.core.sniff import detect_format
    from app.services.workspace_service import iter_files

    paths = iter_files(CORPUS)
    ok_files, formats, cells = 0, {}, 0
    for p in paths:
        try:
            fmt = detect_format(p)
            wb = read_workbook(p)
            n = sum(len(s.cells) for s in wb.sheets)
            if n <= 0:
                continue
            ok_files += 1
            formats[fmt] = formats.get(fmt, 0) + 1
            cells += n
        except Exception:  # noqa: BLE001
            pass
    item("R1-a 全格式读取", ok_files >= len(paths) * 0.95,
         f"{ok_files}/{len(paths)} 个文件 · {cells} 个单元格 · " +
         " ".join(f"{k}:{v}" for k, v in sorted(formats.items())), time.time() - t0)

    wb = read_workbook(SAMPLE)
    tables = detect_workbook(wb)
    t = tables[0]
    ok = (len(t.fields) == 58 and t.header_rows == [1, 2] and t.data_start == 3
          and any(r == 7 for r, _ in t.excluded_rows) and len(t.rows) == 4)
    item("R2-a 样本结构识别", ok,
         f"字段 {len(t.fields)} · 表头 {[r + 1 for r in t.header_rows]} · "
         f"数据起始 第{t.data_start + 1}行 · 排除 {len(t.excluded_rows)} 行（含小计）")

    # ------------------------------------------------------------------ 2. 检索
    t0 = time.time()
    from app.services import search_service, workspace_service

    if do_index:
        task = types.SimpleNamespace(progress=0, message="")
        res = search_service.start_refresh(task, {}, root=workspace_service.get_root())
        index_ok = res["files"] >= 60 and res["cells"] > 100000
        item("R1-b 建立索引", index_ok,
             f"{res['files']} 个文件 · {res['rows']} 行 · {res['cells']} 单元格 · "
             f"构建 {res['took']}s", time.time() - t0)
    else:
        search_service.ensure_index(workspace_service.get_root())

    t0 = time.time()
    hit = search_service.search("DEMO-GD-24001", limit=5)
    hits = hit["hits"]
    ok = bool(hits) and hits[0]["ref"]["a1"] == "A2"
    if hits:
        r = hits[0]["ref"]
        detail = (f"命中 {hit['total']} 条 · 首条 {r['filePath'].split(chr(92))[-1]} "
                  f"第{r['row']}行 {r['colLetter']}列 {r['a1']} · {hit['took']}s")
    else:
        detail = "未命中（语料可能已变化）"
    item("R1-c 检索结果带文件/行/列", ok, detail, time.time() - t0)

    t0 = time.time()
    hit2 = search_service.search("合同", limit=50)
    ok2 = hit2["total"] > 0 and hit2["took"] < 1.0 and len(hit2["hits"]) <= 50
    item("R1-d 检索性能与分页", ok2,
         f"「合同」命中 {hit2['total']} 条 · {hit2['took']}s · 返回 {len(hit2['hits'])} 条", time.time() - t0)

    t0 = time.time()
    out = search_service.export({"query": "DEMO-GD-24001"}, out_dir=os.path.join(ROOT, "data", "output"))
    ok3 = os.path.isfile(out["path"]) and out["count"] >= 1
    item("R1-e 检索结果导出", ok3, f"{os.path.basename(out['path'])} · {out['count']} 行", time.time() - t0)

    # ------------------------------------------------------------------ 3. 对比
    t0 = time.time()
    from app.core.model import CellRef
    from app.core.writers.ooxml_writer import write_cell
    from app.services.compare_service import compare

    left = right = None
    found = {}
    for dirpath, _dirs, files in os.walk(CORPUS):
        for name in files:
            if name.startswith("项目台账统计表 -修改") and name.endswith(".xlsx"):
                found.setdefault(name, os.path.join(dirpath, name))
    for name, path in found.items():
        if "(1)" in name:
            right = path
        else:
            left = path
    if left and right:
        same = compare({"left": {"path": left}, "right": {"path": right}, "keys": [32], "options": {}})
        ok = same["summary"]["changed"] == 0 and same["summary"]["added"] == 0
        item("R3-a 同版本对比无差异", ok, f"matched={same['summary'].get('matched')}", time.time() - t0)

        t0 = time.time()
        mod = fresh("对比_改动版.xlsx")
        write_cell(mod, CellRef(mod, SAMPLE_SHEET, 4, 29), 99999)
        diff = compare({"left": {"path": SAMPLE}, "right": {"path": mod}, "keys": [32], "options": {}})
        changed = [x for x in diff["items"] if x["kind"] == "changed"]
        ok = len(changed) == 1 and changed[0]["leftRef"]["a1"] == "AC4" and changed[0]["rightRef"]["a1"] == "AC4"
        item("R3-b 差异带左右坐标", ok,
             f"changed={len(changed)} · 左 {changed[0]['leftRef']['display'] if changed else '-'}", time.time() - t0)

        t0 = time.time()
        rep = None
        try:
            rep = __import__("app.services.compare_service", fromlist=["export"]).export(
                {"left": {"path": SAMPLE}, "right": {"path": mod}, "keys": [32], "options": {}},
                root=workspace_service.get_root(), out_dir=os.path.join(ROOT, "data", "output"))
        except Exception as e:  # noqa: BLE001
            detail = f"导出失败：{e}"
        item("R3-c 对比报告导出", bool(rep and os.path.isfile(rep["path"])),
             os.path.basename(rep["path"]) if rep else detail, time.time() - t0)
    else:
        item("R3-* 真实两版本对比", False, "语料中未找到成对台账文件")

    # ------------------------------------------------------------------ 4. 模板转换（含转置）
    t0 = time.time()
    from app.services import mapping_service

    reg = mapping_service.add_template({"filePath": SAMPLE, "name": "验收模板"})
    tid = reg["template"]["id"]
    rules = [
        {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "项目名称"}},
         "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 2}},
         "mode": "cell", "transform": "identity", "onMissing": "blank"},
        {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "合同编码"}},
         "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 32}},
         "mode": "cell", "transform": "identity", "onMissing": "blank"},
        {"source": {"row": {"kind": "fixed", "value": 4}, "col": {"kind": "span", "value": {"start": 2, "count": 3}}},
         "target": {"row": {"kind": "span", "value": {"start": 4, "count": 3}}, "col": {"kind": "fixed", "value": 10}},
         "mode": "transpose", "transform": "identity", "onMissing": "blank"},
    ]
    pv = mapping_service.preview_by_ids({"filePath": SAMPLE, "sheet": SAMPLE_SHEET,
                                         "templateId": tid, "rules": rules, "limit": 30})
    af = [c for c in pv["cells"] if c["target"]["a1"] == "AF4"]
    ok_pv = bool(af) and af[0]["value"] == "DEMO-ZS-24001"
    item("R2-b 映射预演（逐单元格）", ok_pv,
         f"{len(pv['cells'])} 格 · AF4={af[0]['value'] if af else '-'}")

    task = types.SimpleNamespace(progress=0, message="")
    conv = mapping_service.start_convert(task, {
        "filePath": SAMPLE, "sheet": SAMPLE_SHEET, "templateId": tid,
        "rules": rules, "outName": "验收_台账输出.xlsx"})
    wb2 = read_workbook(conv["path"])
    ws2 = wb2.sheets[0]
    t2 = [x for x in detect_workbook(wb2) if x.sheet == SAMPLE_SHEET][0]
    src_ws = wb.sheets[0]
    expect = [(src_ws.get(3, c).text if src_ws.get(3, c) else None) for c in (1, 2, 3)]
    got = [(ws2.get(r, 9).text if ws2.get(r, 9) else None) for r in (3, 4, 5)]
    ok_conv = (len(t2.fields) == 58 and t2.header_rows == [1, 2] and len(wb2.sheets[0].merges) >= 20
               and ws2.get(3, 31) is not None and ws2.get(3, 31).text == "DEMO-ZS-24001"
               and expect == got)
    item("R2-c 模板转换（保样式 + 行→列转置）", ok_conv,
         f"{conv['count']} 行 · 字段 {len(t2.fields)} · 合并 {len(wb2.sheets[0].merges)} · "
         f"转置 {'一致' if expect == got else '不一致'}", time.time() - t0)

    # ------------------------------------------------------------------ 5. 编辑写回
    t0 = time.time()
    from app.core.model import EditOp
    from app.services import edit_service

    edit_copy = fresh("编辑_验收副本.xlsx")
    before = edit_service.get_cell(edit_copy, SAMPLE_SHEET, 4, 29)
    wrote = edit_service.write(EditOp(CellRef(edit_copy, SAMPLE_SHEET, 4, 29), 88888, "value",
                                      before["current"]["mtime"], before["current"]["size"]))
    after = edit_service.get_cell(edit_copy, SAMPLE_SHEET, 4, 29)
    entries = edit_service.audit(edit_copy, 10)

    def _num(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return None

    ok_edit = wrote["written"] == 1 and _num(after["value"]) == 88888.0 and bool(entries)
    item("R6-a 单元格写回 + 审计", ok_edit,
         f"AC4 {before['value']} → {after['value']} · 审计 {len(entries)} 条", time.time() - t0)

    t0 = time.time()
    undone = edit_service.undo(entries[0]["id"]) if entries else {}
    restored = edit_service.get_cell(edit_copy, SAMPLE_SHEET, 4, 29)
    ok_undo = bool(undone) and _num(restored["value"]) == _num(before["value"])
    item("R6-b 撤销回滚", ok_undo, f"恢复为 {restored['value']}", time.time() - t0)

    # ------------------------------------------------------------------ 6. 汇总
    total_ok = sum(1 for r in results if r["ok"])
    print(f"\n{total_ok}/{len(results)} 项通过")
    write_report(total_ok)
    return total_ok == len(results)


def write_report(passed: int):
    lines = [
        "# 验收报告",
        "",
        f"- 生成时间：{time.strftime('%Y-%m-%d %H:%M:%S')}",
        f"- 语料目录：`{CORPUS}`",
        f"- 样本：`samples/项目台账统计表.xlsx`",
        f"- 结论：**{passed}/{len(results)} 项通过**",
        "",
        "| # | 验收项 | 结果 | 说明 | 耗时(s) |",
        "|---|---|---|---|---|",
    ]
    for i, r in enumerate(results, 1):
        lines.append(f"| {i} | {r['name']} | {'✅ PASS' if r['ok'] else '❌ FAIL'} | {r['detail']} | {r['cost']:.2f} |")
    lines += [
        "",
        "## 说明",
        "",
        "- 验收脚本：`scripts/acceptance.py`（可重复执行）。",
        "- 全量单元/集成测试：`python -m unittest discover -s tests -v`。",
        "- 报告中的“文件/行/列”均为用户可见坐标（1 基，与 Excel 行标一致）。",
    ]
    os.makedirs(os.path.dirname(REPORT), exist_ok=True)
    with open(REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"报告已写入：{REPORT}")


if __name__ == "__main__":
    do_index = "--no-index" not in sys.argv
    ok = main(do_index)
    sys.exit(0 if ok else 1)
