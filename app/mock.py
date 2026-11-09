# -*- coding: utf-8 -*-
"""示例数据（前端先行开发用）。

真实 service 就绪后，``app/api.py`` 会自动改走真实实现，本模块不再参与。
所有结构必须与 ``CONTRACT.md`` 一致。
"""
from __future__ import annotations

import os
import time

ROOT = os.path.join(os.path.expanduser("~"), "Documents")

FILES = [
    {"id": "f1", "path": ROOT + r"\示例台账-修改.xlsx", "name": "示例台账-修改.xlsx",
     "fmt": "xlsx", "size": 26191, "mtime": 1762652123.0, "sheetsCount": 2, "status": "ok",
     "error": "", "indexedAt": "2026-11-09 15:40:00"},
    {"id": "f2", "path": ROOT + r"\示例预警.xls", "name": "示例预警.xls",
     "fmt": "xls", "size": 82432, "mtime": 1762502271.0, "sheetsCount": 2, "status": "ok",
     "error": "", "indexedAt": "2026-11-09 15:40:00"},
    {"id": "f3", "path": ROOT + r"\示例表格C.xls", "name": "示例表格C.xls",
     "fmt": "xls", "size": 33280, "mtime": 1762652200.0, "sheetsCount": 1, "status": "ok",
     "error": "", "indexedAt": "2026-11-09 15:40:00"},
    {"id": "f4", "path": ROOT + r"\示例清单.xlsx", "name": "示例清单.xlsx",
     "fmt": "xlsx", "size": 17821, "mtime": 1762623575.0, "sheetsCount": 1, "status": "ok",
     "error": "", "indexedAt": "2026-11-09 15:40:00"},
]

SHEETS = [
    {"index": 0, "name": "2024年项目台账（自筹-甲代-施工-日常经营）", "nRows": 63, "nCols": 58,
     "headerRows": [2, 3], "dataStart": 4, "fields": 58, "hidden": False},
    {"index": 1, "name": "Sheet1", "nRows": 1, "nCols": 1, "headerRows": [], "dataStart": 1, "fields": 0,
     "hidden": False},
]

TEMPLATE = {
    "id": "t1",
    "name": "项目台账统计表（2024）",
    "filePath": r"D:\表格筛选与数据核对（示例作者）\samples\项目台账统计表.xlsx",
    "sheet": "2024年项目台账（自筹-甲代-施工-日常经营）",
    "titleRange": [1, 1, 1, 58],
    "headerRange": [2, 3, 1, 58],
    "dataStart": [4, 1],
    "fields": [
        {"index": 0, "name": "序号", "level": 1, "kindHint": "number", "colSpan": [0, 0]},
        {"index": 1, "name": "项目名称", "level": 1, "kindHint": "text", "colSpan": [1, 1]},
        {"index": 2, "name": "资金来源（自筹/代建/施工/日常经营）", "level": 1, "kindHint": "text", "colSpan": [2, 2]},
        {"index": 3, "name": "合同名称", "level": 1, "kindHint": "text", "colSpan": [3, 3]},
        {"index": 18, "name": "乙方中标单位 / 中标单位名称", "level": 2, "kindHint": "text", "colSpan": [23, 23]},
        {"index": 28, "name": "乙方中标金额 / 含税金额（元）", "level": 2, "kindHint": "money", "colSpan": [28, 28]},
        {"index": 31, "name": "合同签订情况 / 合同编码", "level": 2, "kindHint": "code", "colSpan": [31, 31]},
        {"index": 45, "name": "乙方已付款金额 / 金额小计", "level": 2, "kindHint": "money", "colSpan": [45, 45]},
        {"index": 56, "name": "实施部门", "level": 1, "kindHint": "text", "colSpan": [56, 56]},
        {"index": 57, "name": "备注", "level": 1, "kindHint": "text", "colSpan": [57, 57]},
    ],
}

PROFILE = {
    "id": "p1",
    "name": "扣分工单 → 台账",
    "templateId": "t1",
    "sourceSignature": "工单编号|扣分原因|扣分时间",
    "rules": [
        {"source": {"sheet": None, "row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "工单编号"}},
         "target": {"sheet": None, "row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 32}},
         "mode": "cell", "transform": "identity", "onMissing": "blank", "constant": None,
         "note": "工单编号 → 合同编码列"},
        {"source": {"sheet": None, "row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "扣分原因"}},
         "target": {"sheet": None, "row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 58}},
         "mode": "cell", "transform": "identity", "onMissing": "blank", "constant": None,
         "note": "扣分原因 → 备注列"},
    ],
    "notes": "示例方案",
}

HITS = [
    {"ref": {"filePath": FILES[0]["path"], "sheet": "2024年项目台账（自筹-甲代-施工-日常经营）",
             "row": 4, "col": 32, "colLetter": "AF", "a1": "AF4",
             "display": "第4行 · 第AF列(第32列) · AF4"},
     "fieldName": "合同签订情况 / 合同编码", "valueText": "DEMO-ZS-24001",
     "snippet": "DEMO-ZS-24001", "highlight": [0, 17], "score": 9.2},
    {"ref": {"filePath": FILES[1]["path"], "sheet": "Sheet0", "row": 3, "col": 2,
             "colLetter": "B", "a1": "B3", "display": "第3行 · 第B列(第2列) · B3"},
     "fieldName": "工单标题", "valueText": "施工噪音扰民问题（示例区住建局）",
     "snippet": "施工噪音扰民问题", "highlight": [0, 8], "score": 7.4},
    {"ref": {"filePath": FILES[2]["path"], "sheet": "Sheet0", "row": 12, "col": 3,
             "colLetter": "C", "a1": "C12", "display": "第12行 · 第C列(第3列) · C12"},
     "fieldName": "扣分类型", "valueText": "超期未办结", "snippet": "超期未办结",
     "highlight": [0, 5], "score": 6.1},
]


def status(root=None) -> dict:
    return {
        "root": root or ROOT,
        "files": len(FILES),
        "sheets": 4,
        "cells": 12580,
        "indexed": True,
        "indexedAt": "2026-11-09 15:40:00",
        "took": 3.2,
        "notes": ["mock 数据，真实索引就绪后自动替换"],
    }


def index_status(root=None) -> dict:
    return {"indexed": True, "files": len(FILES), "cells": 12580,
            "indexedAt": "2026-11-09 15:40:00", "stale": False}


def files() -> list[dict]:
    return FILES


def sheets(file_id: str) -> list[dict]:
    return SHEETS if file_id == "f1" else SHEETS[:1]


def preview(file_id: str, sheet_index: int, rows: int = 50) -> dict:
    header = ["序号", "项目名称", "资金来源", "合同名称", "委托单位/甲方"]
    body = [
        [1, "示例项目甲—园区改造工程", "企业自筹",
         "示例项目甲—园区改造工程设计进工程建设工程设计合同", ""],
        ["", "", "", "示例项目甲—园区改造工程", ""],
        ["", "", "", "示例项目甲—园区改造工程建设工程监理合同", ""],
        ["", "小计", "", "", ""],
    ]
    out_rows = [{"r": 1, "cells": {"A": "2024年示例区示例公司项目台账（单位：元）"}}]
    out_rows.append({"r": 2, "cells": {chr(65 + i): v for i, v in enumerate(header)}})
    for i, row in enumerate(body, start=3):
        out_rows.append({"r": i, "cells": {chr(65 + j): v for j, v in enumerate(row) if v != ""}})
    return {
        "fileId": file_id, "sheetIndex": sheet_index, "name": SHEETS[0]["name"],
        "nRows": 63, "nCols": 58,
        "merges": [[1, 1, 1, 58], [2, 2, 2, 3], [4, 7, 2, 2]],
        "headerRows": [2, 3], "dataStart": 4,
        "rows": out_rows[:rows],
    }


def search(q: str, **kwargs) -> dict:
    kw = (q or "").strip()
    hits = [h for h in HITS if not kw or kw in h["valueText"]]
    limit = int(kwargs.get("limit") or 50)
    offset = int(kwargs.get("offset") or 0)
    return {"total": len(hits), "took": 0.004, "hits": hits[offset:offset + limit], "query": q}


def export(body: dict, root=None, out_dir=None) -> dict:
    from .core.xlsxwriter import write_rows_xlsx

    out_dir = out_dir or os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "output")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"搜索导出_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    rows = [[h["ref"]["filePath"], h["ref"]["sheet"], h["ref"]["row"], h["ref"]["colLetter"],
             h["ref"]["a1"], h["fieldName"], h["valueText"]] for h in HITS]
    write_rows_xlsx(path, rows, header=["文件地址", "工作表", "第几行", "列", "A1", "字段名", "命中内容"], sheet_name="检索结果")
    return {"path": path, "count": len(rows)}


def get_cell(path: str, sheet: str, row: int, col: int) -> dict:
    from .core.model import CellRef

    ref = CellRef(path, sheet, row, col)
    return {
        "ref": ref.to_dict(),
        "value": "DEMO-ZS-24001",
        "kind": "text",
        "formula": None,
        "display": "DEMO-ZS-24001",
        "fieldName": "合同签订情况 / 合同编码",
        "isSubtotal": False,
        "mergeOwner": None,
        "editable": True,
        "current": {"mtime": 1762652123.0, "size": 26191},
    }


def write(op) -> dict:
    return {"written": 1, "results": [{"ref": op.ref.to_dict(), "result": "ok"}], "mock": True}


def write_batch(ops) -> dict:
    return {"written": len(ops), "results": [{"ref": o.ref.to_dict(), "result": "ok"} for o in ops], "mock": True}


def replace_preview(body: dict) -> dict:
    return {"matches": [{"ref": h["ref"], "oldValue": h["valueText"], "newValue": "（替换后）"} for h in HITS],
            "count": len(HITS), "mock": True}


def replace_apply(body: dict) -> dict:
    return {"written": len(HITS), "mock": True}


def undo(op_id: str) -> dict:
    return {"undone": False, "message": "mock 模式下不可撤销", "mock": True}


def audit(file=None, limit: int = 200) -> list[dict]:
    return [{
        "ts": "2026-11-09 15:41:02",
        "ref": HITS[0]["ref"],
        "oldValue": "DEMO-JQSC-QT-24000",
        "newValue": "DEMO-ZS-24001",
        "backupPath": r"D:\表格筛选与数据核对（示例作者）\data\backup\20261109_154102\项目台账统计表.xlsx",
        "result": "ok",
    }]


def templates() -> list[dict]:
    return [TEMPLATE]


def add_template(body: dict, root=None) -> dict:
    t = dict(TEMPLATE)
    t["id"] = "t" + str(int(time.time() * 1000))
    t["name"] = body.get("name") or "新模板"
    return {"template": t, "mock": True}


def delete_template(tid: str) -> dict:
    return {"deleted": tid, "mock": True}


def set_anchor(tid: str, body: dict) -> dict:
    t = dict(TEMPLATE)
    t["dataStart"] = body.get("dataStart") or t["dataStart"]
    t["headerRange"] = body.get("headerRange") or t["headerRange"]
    return {"template": t, "mock": True}


def profiles() -> list[dict]:
    return [PROFILE]


def save_profile(body: dict) -> dict:
    p = dict(PROFILE)
    p.update({k: v for k, v in body.items() if k in ("name", "rules", "notes", "templateId", "sourceSignature")})
    p["id"] = body.get("id") or "p" + str(int(time.time() * 1000))
    return {"profile": p, "mock": True}


def delete_profile(pid: str) -> dict:
    return {"deleted": pid, "mock": True}


def suggest_by_ids(body: dict, root=None) -> dict:
    return {"suggestions": [
        {"targetField": "合同名称", "sourceField": "扣分原因", "confidence": 0.42, "reason": "文本相似度"},
        {"targetField": "合同签订情况 / 合同编码", "sourceField": "工单编号", "confidence": 0.88, "reason": "别名词典：编号"},
        {"targetField": "备注", "sourceField": "扣分类型", "confidence": 0.31, "reason": "字段类型一致"},
    ], "mock": True}


def preview_by_ids(body: dict, root=None) -> dict:
    n = int(body.get("limit") or 5)
    cells = []
    for i in range(n):
        for c in (0, 1):
            cells.append({
                "target": {"filePath": TEMPLATE["filePath"], "sheet": TEMPLATE["sheet"],
                           "row": 4 + i, "col": 32 + c, "colLetter": "AF", "a1": f"AF{4 + i}",
                           "display": f"第{4 + i}行 · AF列"},
                "sourceRef": {"filePath": FILES[2]["path"], "sheet": "Sheet0", "row": 2 + i,
                              "col": 1 + c, "colLetter": "B", "a1": f"B{2 + i}", "display": "来源"},
                "value": ["DEMO-GD-24001", "超期未办结"][c],
                "ruleIndex": c,
            })
    return {"cells": cells, "truncated": False, "mock": True}


def convert(body: dict, root=None) -> dict:
    from .core.xlsxwriter import write_rows_xlsx

    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "output")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"台账_mock_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
    write_rows_xlsx(path, [["样例", "数据"]], header=["项目名称", "备注"], sheet_name="台账")
    return {"path": path, "count": 1, "mock": True}


def compare(body: dict, root=None) -> dict:
    from .core.model import CellRef

    left = body.get("left") or {}
    right = body.get("right") or {}
    items = [
        {"kind": "changed", "key": "DEMO-ZS-24001", "fieldName": "乙方中标金额 / 含税金额（元）",
         "leftRef": CellRef(left.get("path", ""), left.get("sheet", ""), 4, 29).to_dict(),
         "rightRef": CellRef(right.get("path", ""), right.get("sheet", ""), 4, 29).to_dict(),
         "leftValue": "52800", "rightValue": "56800"},
        {"kind": "added", "key": "DEMO-JS-24009", "fieldName": "",
         "leftRef": None,
         "rightRef": CellRef(right.get("path", ""), right.get("sheet", ""), 9, 1).to_dict(),
         "leftValue": "", "rightValue": "新增项目"},
        {"kind": "removed", "key": "DEMO-JL-24003", "fieldName": "",
         "leftRef": CellRef(left.get("path", ""), left.get("sheet", ""), 6, 1).to_dict(),
         "rightRef": None,
         "leftValue": "监理合同", "rightValue": ""},
    ]
    return {
        "summary": {"matched": 12, "added": 1, "removed": 1, "changed": 1, "leftRows": 14, "rightRows": 14},
        "items": items, "keys": ["合同编码"], "mock": True,
    }
