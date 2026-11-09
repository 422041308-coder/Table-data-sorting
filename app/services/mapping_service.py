# -*- coding: utf-8 -*-
"""S4 服务层：模板登记 / 映射方案 CRUD / 自动推荐 / 预演 / 转换（长任务）。

数据落盘（UTF-8 JSON，原子写）::

    data/templates/<templateId>.json   模板登记（TemplateDef + createdAt/updatedAt）
    data/templates/files/*.xlsx        可选：上传后复制保存的模板文件
    data/profiles/<profileId>.json     映射方案（MappingProfile）
    data/output/*.xlsx                 转换产物（保样式克隆输出）

文件定位统一走 ``app/services/workspace_service.py``（``file_id`` / ``resolve_file``），
读表优先 S1/S2（``read_workbook`` + ``detect_table``）；S1/S2 未就绪时退回本模块的
临时兜底（基于 ``app/core/template.py`` 的 zip/XML 读取，仅支持 xlsx）。

坐标：进来的 body 字段是 1 基（与 Excel 一致）；返回给前端的 CellRef 也是 1 基。
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import time
import uuid
from typing import Any

from ..core import mapping as mapping_core
from ..core import template as template_core
from ..core.model import (
    KIND_NUMBER,
    KIND_TEXT,
    Cell,
    Field,
    FileRef,
    MappingProfile,
    TableModel,
    TemplateDef,
)

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJECT_ROOT = os.path.dirname(APP_DIR)
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
TEMPLATES_DIR = os.path.join(DATA_DIR, "templates")
PROFILES_DIR = os.path.join(DATA_DIR, "profiles")
OUTPUT_DIR = os.path.join(DATA_DIR, "output")
TEMPLATE_FILES_DIR = os.path.join(TEMPLATES_DIR, "files")

_BAD_NAME = re.compile(r'[\\/:*?"<>|\r\n\t]+')


# ==========================================================================
# 基础工具
# ==========================================================================
def _workspace():
    try:
        from . import workspace_service
    except ImportError:
        return None
    return workspace_service


def _file_id(path: str) -> str:
    ws = _workspace()
    if ws is not None:
        try:
            return ws.file_id(path)
        except Exception:  # noqa: BLE001
            pass
    return hashlib.sha1(os.path.abspath(path).encode("utf-8")).hexdigest()[:12]


def _ensure_dirs() -> None:
    for d in (TEMPLATES_DIR, PROFILES_DIR, OUTPUT_DIR):
        os.makedirs(d, exist_ok=True)


def _atomic_json(path: str, data: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")
    os.replace(tmp, path)


def _read_json(path: str) -> dict | None:
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def _resolve_path(raw: str, root=None) -> str:
    """文件 id 或路径 → 绝对路径（优先 workspace_service）。"""
    if isinstance(raw, dict):
        raw = raw.get("filePath") or raw.get("path") or raw.get("file") or ""
    raw = str(raw or "").strip()
    if not raw:
        raise ValueError("缺少文件参数（fileId / filePath）")
    ws = _workspace()
    if ws is not None:
        try:
            return os.path.abspath(ws.resolve_file(raw))
        except Exception:  # noqa: BLE001 - 交给下面的兜底
            pass
    p = os.path.expanduser(raw)
    if not os.path.isabs(p):
        p = os.path.join(str(root) if root else PROJECT_ROOT, p)
    return os.path.abspath(p)


def _pick_sheet(wb, body: dict, who: str = "来源"):
    sheets = list(getattr(wb, "sheets", []) or [])
    if not sheets:
        raise ValueError(f"{who}文件没有工作表：{getattr(wb, 'path', '')}")
    spec = body.get("sheet")
    idx = body.get("sheetIndex")
    if isinstance(spec, str) and spec.strip():
        for s in sheets:
            if s.name == spec:
                return s
        raise ValueError(f"{who}工作表不存在：{spec!r}")
    if idx is not None and not isinstance(idx, bool):
        try:
            i = int(idx)
        except (TypeError, ValueError):
            raise ValueError(f"sheetIndex 必须是整数：{idx!r}")
        if i < 0 or i >= len(sheets):
            raise ValueError(f"{who}工作表序号越界：{i}（共 {len(sheets)} 个）")
        return sheets[i]
    # 默认取第一个有内容的工作表
    for s in sheets:
        if getattr(s, "cells", None):
            return s
    return sheets[0]


def _file_ref(path: str, wb, fid: str | None = None) -> FileRef:
    try:
        st = os.stat(path)
        size, mtime = st.st_size, st.st_mtime
    except OSError:
        size, mtime = 0, 0.0
    return FileRef(
        id=fid or _file_id(path),
        path=os.path.abspath(path),
        name=os.path.basename(path),
        fmt=getattr(wb, "fmt", "unknown") if wb is not None else "unknown",
        size=size,
        mtime=mtime,
        sheets_count=len(getattr(wb, "sheets", []) or []) if wb is not None else 0,
        status="ok",
    )


# ==========================================================================
# 读表：优先 S1/S2，未就绪走兜底
# ==========================================================================
def _read_workbook_compat(path: str):
    try:
        from ..core.readers import read_workbook  # type: ignore

        return read_workbook(path)
    except ImportError:
        pass
    fmt = ""
    try:
        from ..core.sniff import detect_format

        fmt = detect_format(path)
    except Exception:  # noqa: BLE001
        pass
    if fmt in ("xlsx", "xlsm", "xlsb", "et-zip"):
        from ..core.readers.ooxml import read_ooxml

        return read_ooxml(path)
    if fmt in ("xls", "et-ole"):
        try:
            from ..core.readers.biff import read_biff

            return read_biff(path)
        except ImportError:
            pass
    # 最后再试一次 ooxml（兼容 sniff 尚未就绪的情况）
    from ..core.readers.ooxml import read_ooxml

    return read_ooxml(path)


def _derive_table_from_sheet(wb, sheet, path: str, fid: str | None) -> TableModel:
    """用 S2 的 detect_table 识别；不可用时退回单行表头兜底。"""
    ref = _file_ref(path, wb, fid)
    try:
        from ..core.detect import detect_table

        try:
            table = detect_table(wb, sheet, file_ref=ref)
        except TypeError:
            table = detect_table(wb, sheet)
        if table is not None:
            if table.source is None:
                table.source = ref
            if not table.sheet:
                table.sheet = sheet.name
            return table
    except ImportError:
        pass
    return _fallback_table(path, sheet.name, ref)


def _load_table(body: dict, root=None, who: str = "来源") -> TableModel:
    raw = (
        body.get("filePath")
        or body.get("path")
        or body.get("file")
        or body.get("fileId")
        or body.get("sourceFile")
        or body.get("source")
    )
    if isinstance(raw, dict):
        raw = raw.get("filePath") or raw.get("path") or raw.get("file")
    if not raw and isinstance(body.get("sources"), list) and body["sources"]:
        raw = body["sources"][0]
    if not raw:
        raise ValueError(f"{who}表缺少文件参数（fileId 或 filePath）")
    path = _resolve_path(str(raw), root)
    if not os.path.isfile(path):
        raise ValueError(f"{who}表文件不存在：{path}")
    fid = _file_id(path)
    ws = _workspace()
    if ws is not None:
        try:
            wb, tables = ws.load(fid)
            chosen = _pick_sheet(wb, body, who)
            for t in tables or []:
                if t.sheet == chosen.name:
                    return t
            return _derive_table_from_sheet(wb, chosen, path, fid)
        except Exception:  # noqa: BLE001 - 文件不在工作区/读表失败 → 直接读
            pass
    wb = _read_workbook_compat(path)
    chosen = _pick_sheet(wb, body, who)
    return _derive_table_from_sheet(wb, chosen, path, fid)


def _load_sources(body: dict, root=None) -> list[TableModel]:
    """支持单来源（fileId/filePath）与批量来源（body.sources[]，append 语义）。"""
    specs = body.get("sources")
    if not isinstance(specs, list) or not specs:
        return [_load_table(body, root, "来源")]
    out: list[TableModel] = []
    for i, spec in enumerate(specs):
        sub = dict(body)
        sub.pop("sources", None)
        if isinstance(spec, dict):
            sub.update(spec)
        else:
            sub["filePath"] = spec
        try:
            out.append(_load_table(sub, root, f"来源{i + 1}"))
        except ValueError as e:
            raise ValueError(f"来源{i + 1}读取失败：{e}") from e
    return out


# ---------------------------- 兜底识别（xlsx） ----------------------------
_FALLBACK_TOTAL = ("小计", "合计", "总计", "其中", "累计")


def _fallback_table(path: str, sheet_name: str, ref: FileRef | None = None) -> TableModel:
    tpl = template_core.parse_template(path, sheet_name)
    grid = template_core.read_grid(path, tpl.sheet)
    cells = grid["cells"]
    header_rows = list(range(tpl.header_range[0] - 1, tpl.header_range[1]))
    data_start = tpl.data_start[0] - 1
    max_r = max((r for (r, _c) in cells), default=data_start - 1)
    fields = list(tpl.fields)
    rows: list[dict[int, Cell]] = []
    row_numbers: list[int] = []
    excluded: list[tuple[int, str]] = []
    for r in range(data_start, max_r + 1):
        row: dict[int, Cell] = {}
        texts: list[str] = []
        nums = 0
        for c in range(grid["nCols"]):
            v = cells.get((r, c))
            if v is None or v == "":
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                row[c] = Cell(raw=v, kind=KIND_NUMBER, text=str(v), num=float(v))
                nums += 1
            elif isinstance(v, bool):
                row[c] = Cell(raw=v, kind="bool", text=str(v))
            else:
                row[c] = Cell(raw=v, kind=KIND_TEXT, text=str(v))
                texts.append(str(v))
        if not row:
            excluded.append((r, "空行"))
            continue
        first = texts[0].strip() if texts else ""
        if any(first.startswith(k) for k in _FALLBACK_TOTAL):
            excluded.append((r, "小计/合计行"))
            continue
        if len(row) <= 2 and nums == 0:
            excluded.append((r, "占位/空行"))
            continue
        rows.append(row)
        row_numbers.append(r)
    return TableModel(
        source=ref,
        sheet=grid["sheet"],
        header_rows=header_rows,
        data_start=data_start,
        data_end=(row_numbers[-1] if row_numbers else max(data_start - 1, 0)),
        excluded_rows=excluded,
        fields=fields,
        rows=rows,
        row_numbers=row_numbers,
        parse_notes=["S1/S2 未就绪，mapping_service 临时兜底读取（xlsx）"],
    )


# ==========================================================================
# 模板存取
# ==========================================================================
def _template_path(tid: str) -> str:
    return os.path.join(TEMPLATES_DIR, f"{_safe_id(tid)}.json")


def _safe_id(v: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]", "_", str(v or ""))


def _template_dict(tid: str) -> dict:
    if not tid:
        raise ValueError("缺少 templateId")
    data = _read_json(_template_path(tid))
    if data is None:
        raise ValueError(f"模板未登记：{tid}")
    return data


def _template_def_from_dict(data: dict) -> TemplateDef:
    path = str(data.get("filePath") or "")
    if not path or not os.path.isfile(path):
        raise ValueError(f"模板文件不存在或未登记：{path or data.get('id')}")
    fields = [
        Field(
            index=int(f.get("index", i)),
            name=str(f.get("name") or ""),
            path=str(f.get("name") or ""),
            level=int(f.get("level") or 1),
            kind_hint=str(f.get("kindHint") or KIND_TEXT),
            col_span=tuple(f.get("colSpan") or (f.get("index", i), f.get("index", i))),
        )
        for i, f in enumerate(data.get("fields") or [])
    ]
    title = data.get("titleRange")
    return TemplateDef(
        id=str(data.get("id") or ""),
        name=str(data.get("name") or os.path.basename(path)),
        file_path=os.path.abspath(path),
        sheet=str(data.get("sheet") or ""),
        title_range=tuple(int(x) for x in title) if title else None,
        header_range=tuple(int(x) for x in (data.get("headerRange") or (1, 1, 1, 1))),
        data_start=tuple(int(x) for x in (data.get("dataStart") or (2, 1))),
        fields=fields,
    )


def _template_from_body(body: dict, root=None) -> TemplateDef:
    tid = body.get("templateId") or body.get("template_id")
    if tid:
        return _template_def_from_dict(_template_dict(str(tid)))
    tpl = body.get("template")
    if isinstance(tpl, dict):
        path = tpl.get("filePath") or tpl.get("path")
        if not path:
            raise ValueError("body.template 缺少 filePath")
        path = _resolve_path(str(path), root)
        return template_core.parse_template(
            path,
            tpl.get("sheet"),
            header_range=_parse_range(tpl.get("headerRange")),
            data_start=_parse_cell(tpl.get("dataStart")),
            title_range=_parse_range(tpl.get("titleRange")),
        )
    raise ValueError("缺少 templateId（或 body.template）")


def _template_for_body(body: dict, root=None) -> TemplateDef:
    """取模板：优先 body.templateId；否则由 body.profileId 关联的方案反查。"""
    if body.get("templateId") or body.get("template_id") or isinstance(body.get("template"), dict):
        return _template_from_body(body, root)
    pid = body.get("profileId") or body.get("profile_id")
    if pid:
        data = _profile_dict(str(pid))
        tid = str(data.get("templateId") or "")
        if tid:
            return _template_def_from_dict(_template_dict(tid))
    return _template_from_body(body, root)


def _parse_cell(v) -> tuple[int, int] | None:
    """接受 [row, col] / "F4" / {"row":4,"col":6} / "row,col"。"""
    if v is None:
        return None
    if isinstance(v, (list, tuple)) and len(v) == 2:
        return int(v[0]), int(v[1])
    if isinstance(v, dict):
        row = v.get("row", v.get("r"))
        col = v.get("col", v.get("c") or v.get("column"))
        if row is None or col is None:
            raise ValueError(f"坐标对象缺少 row/col：{v!r}")
        if isinstance(col, str) and col.isalpha():
            from ..core.model import col_index

            col = col_index(col)
        return int(row), int(col)
    if isinstance(v, str):
        s = v.strip().upper().replace("$", "")
        if "," in s:
            a, b = s.split(",", 1)
            return int(a), int(b)
        from ..core.model import split_a1

        row, col = split_a1(s)
        return row, col
    raise ValueError(f"无法解析单元格坐标：{v!r}")


def _parse_range(v):
    """接受 [r1,r2,c1,c2] / "A2:BF3" / {"row1","row2","col1","col2"}。"""
    if v is None:
        return None
    if isinstance(v, (list, tuple)) and len(v) == 4:
        return int(v[0]), int(v[1]), int(v[2]), int(v[3])
    if isinstance(v, str):
        s = v.strip().upper().replace("$", "")
        if ":" not in s:
            raise ValueError(f"区域需要 A1:B2 形式：{v!r}")
        from ..core.model import split_a1

        a, b = s.split(":", 1)
        r1, c1 = split_a1(a)
        r2, c2 = split_a1(b)
        return min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)
    if isinstance(v, dict):
        r1 = v.get("row1", v.get("r1", v.get("startRow")))
        r2 = v.get("row2", v.get("r2", v.get("endRow")))
        c1 = v.get("col1", v.get("c1", v.get("startCol")))
        c2 = v.get("col2", v.get("c2", v.get("endCol")))
        if None in (r1, r2, c1, c2):
            raise ValueError(f"区域对象缺少 row1/row2/col1/col2：{v!r}")
        return int(r1), int(r2), int(c1), int(c2)
    raise ValueError(f"无法解析区域：{v!r}")


def _save_template(tpl: TemplateDef, extra: dict | None = None, created_at: str | None = None) -> dict:
    _ensure_dirs()
    data = tpl.to_dict()
    data.update(extra or {})
    data["createdAt"] = created_at or _now()
    data["updatedAt"] = _now()
    _atomic_json(_template_path(tpl.id), data)
    return data


# ==========================================================================
# 模板接口（CONTRACT §4 S4）
# ==========================================================================
def templates() -> list[dict]:
    """已登记模板清单。"""
    _ensure_dirs()
    out: list[dict] = []
    for name in sorted(os.listdir(TEMPLATES_DIR)):
        if not name.endswith(".json"):
            continue
        data = _read_json(os.path.join(TEMPLATES_DIR, name))
        if data:
            out.append(data)
    out.sort(key=lambda d: str(d.get("name") or ""))
    return out


def add_template(body: dict, root=None) -> dict:
    """登记模板：``{name?, filePath?/path?, sheet?, dataStart?, headerRange?, titleRange?}``。

    也支持直接上传：body 含 ``__raw__``(bytes) + ``__filename__`` 时先保存到
    ``data/templates/files/`` 再解析。
    """
    if not isinstance(body, dict):
        raise ValueError("模板登记参数必须是对象")
    raw = body.get("__raw__")
    if isinstance(raw, (bytes, bytearray)):
        _ensure_dirs()
        os.makedirs(TEMPLATE_FILES_DIR, exist_ok=True)
        filename = _BAD_NAME.sub("_", str(body.get("__filename__") or "template.xlsx"))
        if not filename.lower().endswith((".xlsx", ".xlsm")):
            raise ValueError("只支持 xlsx/xlsm 模板（.xls 请先另存为 xlsx）")
        stamp = time.strftime("%Y%m%d_%H%M%S")
        dest = os.path.join(TEMPLATE_FILES_DIR, f"{stamp}_{uuid.uuid4().hex[:6]}_{filename}")
        with open(dest, "wb") as f:
            f.write(bytes(raw))
        file_path = dest
    else:
        raw_path = body.get("filePath") or body.get("path") or body.get("file")
        if not raw_path:
            raise ValueError("缺少模板文件路径（filePath 或上传文件）")
        file_path = _resolve_path(str(raw_path), root)
    if not os.path.isfile(file_path):
        raise ValueError(f"模板文件不存在：{file_path}")
    if not file_path.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("模板必须是 xlsx/xlsm（.xls 请先另存为 xlsx）")

    header_range = _parse_range(body.get("headerRange"))
    data_start = _parse_cell(body.get("dataStart") or body.get("dataStartCell"))
    title_range = _parse_range(body.get("titleRange"))
    sheet = body.get("sheet") if isinstance(body.get("sheet"), str) and body.get("sheet") else None
    tpl = template_core.parse_template(
        file_path,
        sheet,
        header_range=header_range,
        data_start=data_start,
        title_range=title_range,
    )
    old = _read_json(_template_path(tpl.id)) or {}
    tpl.name = str(body.get("name") or tpl.name or os.path.basename(file_path))
    data = _save_template(
        tpl,
        extra={"originPath": file_path},
        created_at=old.get("createdAt"),
    )
    return {"template": data}


def delete_template(tid: str) -> dict:
    """删除模板登记（不删除用户的原始模板文件；仅删除登记时复制进 data 的副本）。"""
    data = _read_json(_template_path(tid))
    if data is None:
        raise ValueError(f"模板未登记：{tid}")
    path = str(data.get("filePath") or "")
    try:
        os.remove(_template_path(tid))
    except OSError as e:
        raise ValueError(f"删除模板登记失败：{e}")
    # 若模板文件是我们复制保存的，一并删除
    try:
        if path and os.path.abspath(path).startswith(os.path.abspath(TEMPLATE_FILES_DIR) + os.sep):
            os.remove(path)
    except OSError:
        pass
    affected = []
    for p in profiles():
        if str(p.get("templateId") or "") == tid:
            affected.append(p.get("id"))
    return {"deleted": tid, "affectedProfiles": affected}


def set_anchor(tid: str, body: dict) -> dict:
    """重新框选模板锚点：数据起始格 / 表头区 / 标题区，并重新解析字段。"""
    data = _template_dict(tid)
    old_ds = tuple(int(x) for x in (data.get("dataStart") or (2, 1)))
    old_hr = tuple(int(x) for x in (data.get("headerRange") or (1, 1, 1, 1)))
    old_tr = data.get("titleRange")
    ds = _parse_cell(body.get("dataStart") or body.get("dataStartCell")) or old_ds
    hr = _parse_range(body.get("headerRange")) or old_hr
    tr = _parse_range(body.get("titleRange")) if body.get("titleRange") is not None else (
        tuple(int(x) for x in old_tr) if old_tr else None
    )
    # 支持 {"headerRows":[2,3]} 这种前端框选返回值
    rows = body.get("headerRows")
    if rows and not body.get("headerRange"):
        rs = sorted(int(x) for x in rows)
        hr = (rs[0], rs[-1], hr[2], hr[3])
    tpl = template_core.parse_template(
        data["filePath"],
        data.get("sheet") or None,
        header_range=tuple(hr),
        data_start=tuple(ds),
        title_range=tr,
        tpl_id=tid,
    )
    if body.get("name"):
        tpl.name = str(body["name"])
    else:
        tpl.name = str(data.get("name") or tpl.name)
    saved = _save_template(tpl, created_at=data.get("createdAt"))
    saved.update({k: v for k, v in data.items() if k not in saved})
    _atomic_json(_template_path(tid), saved)
    return {"template": saved}


# ==========================================================================
# 方案存取
# ==========================================================================
def _profile_path(pid: str) -> str:
    return os.path.join(PROFILES_DIR, f"{_safe_id(pid)}.json")


def _profile_dict(pid: str) -> dict:
    data = _read_json(_profile_path(pid))
    if data is None:
        raise ValueError(f"映射方案不存在：{pid}")
    return data


def profiles() -> list[dict]:
    """映射方案清单。"""
    _ensure_dirs()
    out: list[dict] = []
    for name in sorted(os.listdir(PROFILES_DIR)):
        if not name.endswith(".json"):
            continue
        data = _read_json(os.path.join(PROFILES_DIR, name))
        if data:
            out.append(data)
    out.sort(key=lambda d: str(d.get("name") or ""))
    return out


def save_profile(body: dict) -> dict:
    """新增/更新映射方案：``{id?, name, templateId, rules[], sourceSignature?, notes?}``。"""
    if not isinstance(body, dict):
        raise ValueError("映射方案必须是对象")
    pid = str(body.get("id") or "").strip()
    old = _read_json(_profile_path(pid)) if pid else None
    if pid and old is None:
        pid = ""  # 传了不存在的 id：按新增处理
    if not pid:
        pid = "p-" + uuid.uuid4().hex[:10]
    template_id = str(body.get("templateId") or body.get("template_id") or (old or {}).get("templateId") or "")
    if not template_id:
        raise ValueError("缺少 templateId：方案必须关联一个已登记模板")
    _template_dict(template_id)  # 校验模板存在
    rules_raw = body.get("rules")
    if rules_raw is None:
        rules_raw = (old or {}).get("rules") or []
    if not isinstance(rules_raw, list):
        raise ValueError("rules 必须是数组")
    rules = []
    for i, r in enumerate(rules_raw):
        try:
            rules.append(mapping_core.rule_from_dict(r).to_dict())
        except ValueError as e:
            raise ValueError(f"第 {i + 1} 条规则非法：{e}") from e
    data = {
        "id": pid,
        "name": str(body.get("name") or (old or {}).get("name") or "未命名方案"),
        "templateId": template_id,
        "sourceSignature": str(
            body.get("sourceSignature")
            or body.get("source_signature")
            or (old or {}).get("sourceSignature")
            or ""
        ),
        "rules": rules,
        "notes": str(body.get("notes") if body.get("notes") is not None else (old or {}).get("notes") or ""),
        "createdAt": (old or {}).get("createdAt") or _now(),
        "updatedAt": _now(),
    }
    _atomic_json(_profile_path(pid), data)
    return {"profile": data}


def delete_profile(pid: str) -> dict:
    if _read_json(_profile_path(pid)) is None:
        raise ValueError(f"映射方案不存在：{pid}")
    try:
        os.remove(_profile_path(pid))
    except OSError as e:
        raise ValueError(f"删除方案失败：{e}")
    return {"deleted": pid}


def _load_profile(pid: str) -> tuple[MappingProfile, dict]:
    data = _profile_dict(pid)
    prof = mapping_core.profile_from_dict(data)
    prof.id = prof.id or pid
    return prof, data


# ==========================================================================
# 自动推荐 / 预演
# ==========================================================================
def _source_fields_payload(source: TableModel) -> list[dict]:
    return [
        {"index": f.index, "name": f.name, "kindHint": f.kind_hint, "level": f.level}
        for f in source.fields
    ]


def _template_fields_payload(template: TemplateDef) -> list[dict]:
    return [
        {"index": f.index, "name": f.name, "kindHint": f.kind_hint, "level": f.level}
        for f in template.fields
    ]


def suggest_by_ids(body: dict, root=None) -> dict:
    """自动推荐字段映射；若来源表头指纹命中已有方案，一并返回匹配方案。"""
    source = _load_table(body, root, "来源")
    template = _template_from_body(body, root)
    suggestions = mapping_core.suggest(source, template)
    signature = mapping_core.source_signature(source)
    matched = None
    for p in profiles():
        if str(p.get("sourceSignature") or "") == signature and signature:
            if str(p.get("templateId") or "") in ("", template.id):
                matched = p
                break
    return {
        "suggestions": suggestions,
        "sourceSignature": signature,
        "templateSignature": mapping_core.template_signature(template),
        "templateId": template.id,
        "sourceSheet": source.sheet,
        "sourceFields": _source_fields_payload(source),
        "templateFields": _template_fields_payload(template),
        "recordCount": len(source.rows),
        "matchedProfile": matched,
    }


def _profile_from_body(body: dict, template: TemplateDef) -> tuple[MappingProfile, str]:
    """优先 body.rules（前端实时规则），否则 body.profileId 读已保存方案。"""
    rules_raw = body.get("rules")
    if rules_raw is None and isinstance(body.get("profile"), dict):
        rules_raw = body["profile"].get("rules")
    if rules_raw is not None:
        if isinstance(rules_raw, (str, bytes)):
            try:
                rules_raw = json.loads(rules_raw)
            except ValueError as e:
                raise ValueError(f"rules 不是合法 JSON：{e}") from e
        if not isinstance(rules_raw, list):
            raise ValueError("rules 必须是数组")
        rules = [mapping_core.rule_from_dict(r) for r in rules_raw]
        prof = MappingProfile(
            id="__preview__",
            name=str(body.get("name") or "未保存方案"),
            template_id=template.id,
            source_signature=str(body.get("sourceSignature") or ""),
            rules=rules,
        )
        return prof, ""
    pid = body.get("profileId") or body.get("profile_id")
    if not pid:
        raise ValueError("缺少 rules 或 profileId：请先在映射页配置规则")
    prof, _data = _load_profile(str(pid))
    return prof, str(pid)


def preview_by_ids(body: dict, root=None) -> dict:
    """映射预演：逐单元格返回"来源坐标→目标坐标→值"（不落盘）。"""
    sources = _load_sources(body, root)
    source = sources[0]
    template = _template_for_body(body, root)
    profile, pid = _profile_from_body(body, template)
    try:
        limit = int(body.get("limit") if body.get("limit") is not None else 50)
    except (TypeError, ValueError):
        limit = 50
    result = mapping_core.preview(profile, source, template, limit=limit)
    if len(sources) > 1:
        result["warnings"] = list(result.get("warnings") or []) + [
            f"共 {len(sources)} 个来源，预演只展示第 1 个来源 {source.path or source.sheet}；"
            "生成时会按追加方式写入全部来源"
        ]
    result.update(
        {
            "sourceSignature": mapping_core.source_signature(source),
            "templateId": template.id,
            "profileId": pid,
            "sourceSheet": source.sheet,
            "sourceFields": _source_fields_payload(source),
            "templateFields": _template_fields_payload(template),
            "dataStart": list(template.data_start),
        }
    )
    return result


# ==========================================================================
# 转换（长任务）
# ==========================================================================
def _progress(task, pct: int, message: str) -> None:
    if task is None:
        return
    try:
        task.progress = int(pct)
        task.message = str(message)
    except Exception:  # noqa: BLE001
        pass


def _safe_filename(name: str, fallback: str = "台账") -> str:
    name = _BAD_NAME.sub("_", str(name or "")).strip(" ._")
    if not name:
        name = fallback
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return name


def _unique_output(path: str) -> str:
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    for i in range(1, 1000):
        cand = f"{stem}_{i}{ext}"
        if not os.path.exists(cand):
            return cand
    return f"{stem}_{uuid.uuid4().hex[:6]}{ext}"


def _build_inject_rows(
    cells: dict, data_start: tuple[int, int]
) -> tuple[list[dict[int, Any]], int, list[str]]:
    """把 ``{(r0,c0): value}`` 转成 ``inject(rows=[{1基列: 值}])`` 需要的行列表。"""
    ds_row, _ds_col = int(data_start[0]), int(data_start[1])
    warnings: list[str] = []
    by_row: dict[int, dict[int, Any]] = {}
    skipped = 0
    for (r0, c0), value in cells.items():
        try:
            r1 = int(r0) + 1
            c1 = int(c0) + 1
        except (TypeError, ValueError):
            continue
        if r1 < ds_row:
            skipped += 1
            continue
        by_row.setdefault(r1, {})[c1] = value
    if skipped:
        warnings.append(
            f"{skipped} 个目标格早于模板数据起始行（第 {ds_row} 行），已忽略；请检查规则的目标行选择器"
        )
    if not by_row:
        return [], 0, warnings
    last = max(by_row)
    rows: list[dict[int, Any]] = []
    for r1 in range(ds_row, last + 1):
        rows.append(by_row.get(r1, {}))
    return rows, skipped, warnings


def start_convert(task, body: dict, root=None) -> dict:
    """执行转换（长任务）：映射 → 保样式克隆注入 → 自检。返回 ``{"path","count"}``。

    支持单来源（filePath/fileId）与批量来源（``sources: [...]``）。
    多来源按**追加**语义：第 k 个来源的记录接在第 k-1 个来源之后写入同一模板。
    """
    if not isinstance(body, dict):
        raise ValueError("转换参数必须是对象")
    _progress(task, 5, "读取来源表…")
    sources = _load_sources(body, root)
    total_records = sum(len(s.rows) for s in sources)
    _progress(task, 15, "解析模板…")
    template = _template_for_body(body, root)
    profile, pid = _profile_from_body(body, template)
    ds_row = int(template.data_start[0])

    _progress(task, 30, f"计算映射（{total_records} 行 × {len(profile.rules)} 条规则）…")
    all_cells: dict[tuple[int, int], Any] = {}
    offset = 0
    for idx, source in enumerate(sources):
        cells = mapping_core.apply_rules(profile, source, template)
        for (r0, c0), value in cells.items():
            r1 = r0 + 1
            if r1 < ds_row:
                continue
            if idx > 0:
                r1 += offset  # 追加：后续来源整体下移
            all_cells[(r1 - 1, c0)] = value
        offset += len(source.rows)
    rows, skipped, warnings = _build_inject_rows(all_cells, template.data_start)
    if len(sources) > 1:
        warnings.append(f"批量转换：{len(sources)} 个来源按追加方式写入同一模板，共 {total_records} 条记录")

    _progress(task, 55, f"注入模板（{len(rows)} 行数据）…")
    _ensure_dirs()
    out_name = body.get("outName") or body.get("name")
    if not out_name:
        out_name = f"{_safe_filename(template.name or '台账')[:-5]}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_path = _unique_output(os.path.join(OUTPUT_DIR, _safe_filename(out_name)))
    template_path = template.file_path
    if not os.path.isfile(template_path):
        raise ValueError(f"模板文件不存在：{template_path}")
    if not template_path.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("模板必须是 xlsx/xlsm，才能保样式克隆注入")

    path = template_core.inject(
        template_path, out_path, template.data_start, rows, sheet=template.sheet
    )

    _progress(task, 85, "重新解析自检…")
    problems = template_core.verify_injected(path, template.data_start, rows, template.sheet)
    if problems:
        raise ValueError("生成文件自检失败：" + "；".join(problems[:3]))
    _progress(task, 100, f"完成：{os.path.basename(path)}（{len(rows)} 行）")
    result = {
        "path": path,
        "name": os.path.basename(path),
        "count": len(rows),
        "records": total_records,
        "sourceCount": len(sources),
        "profileId": pid,
        "templateId": template.id,
        "skippedCells": skipped,
        "warnings": warnings,
    }
    try:
        if task is not None:
            task.result = result
    except Exception:  # noqa: BLE001
        pass
    return result


# 兼容 api.py 里"若没有 start_convert 则回退 convert"的分支
def convert(body: dict, root=None) -> dict:
    return start_convert(None, body, root=root)
