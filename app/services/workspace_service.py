# -*- coding: utf-8 -*-
"""工作区服务：目录扫描、文件清单、工作表/预览（Lead 维护的集成层）。

* 文件 id 约定：``sha1(绝对路径)[:12]``，全应用统一；
* 工作表/预览返回给前端的行号一律 **1 基**（与 Excel 一致），内部模型是 0 基。
"""
from __future__ import annotations

import hashlib
import json
import os
import threading
import time
from collections import OrderedDict

from ..core.model import FMT_UNKNOWN, FileRef, WorkbookModel

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 打包成单文件 EXE 时用 EXCEL_TOOLKIT_HOME 指定 data/ 的位置
PROJECT_ROOT = os.environ.get("EXCEL_TOOLKIT_HOME") or os.path.dirname(APP_DIR)
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
SETTINGS_PATH = os.path.join(DATA_DIR, "settings.json")

SUPPORTED_EXTS = {".xlsx", ".xlsm", ".xlsb", ".xls", ".csv", ".et"}

DEFAULT_ROOT = os.environ.get(
    "EXCEL_TOOLKIT_ROOT",
    os.path.join(os.path.expanduser("~"), "Documents"),
)

_EXTS_FMT = {
    ".xlsx": "xlsx", ".xlsm": "xlsm", ".xlsb": "xlsb", ".xls": "xls",
    ".csv": "csv", ".et": "et-zip",
}

_state = {
    "root": None,
    "files": OrderedDict(),   # fid -> FileRef
    "order": [],
    "scan_at": "",
    "errors": [],
}
_cache: "OrderedDict[str, tuple]" = OrderedDict()  # fid -> (WorkbookModel, tables)
_lock = threading.RLock()


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------
def file_id(path: str) -> str:
    return hashlib.sha1(os.path.abspath(path).encode("utf-8")).hexdigest()[:12]


def _load_settings() -> dict:
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:  # noqa: BLE001
        return {}


def _save_settings(data: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = SETTINGS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, SETTINGS_PATH)


def _detect_fmt(path: str) -> str:
    try:
        from ..core.sniff import detect_format

        fmt = detect_format(path)
        if fmt and fmt != FMT_UNKNOWN:
            return fmt
    except Exception:  # noqa: BLE001
        pass
    return _EXTS_FMT.get(os.path.splitext(path)[1].lower(), FMT_UNKNOWN)


def iter_files(root: str):
    """递归枚举受支持的表格文件（跳过 ~$ 临时文件、隐藏目录与大文件）。"""
    root = os.path.abspath(root or "")
    if not os.path.isdir(root):
        return []
    out = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for name in sorted(filenames):
            if name.startswith("~$") or name.startswith("."):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext not in SUPPORTED_EXTS:
                continue
            full = os.path.join(dirpath, name)
            try:
                size = os.path.getsize(full)
            except OSError:
                continue
            if size <= 0 or size > 80 * 1024 * 1024:
                continue
            out.append(full)
    return out


# --------------------------------------------------------------------------
# 根目录
# --------------------------------------------------------------------------
def get_root() -> str:
    with _lock:
        if _state["root"]:
            return _state["root"]
        root = _load_settings().get("root") or DEFAULT_ROOT
        _state["root"] = os.path.abspath(root)
        return _state["root"]


def set_root(path: str) -> str:
    path = os.path.abspath(path)
    if not os.path.isdir(path):
        raise ValueError(f"目录不存在：{path}")
    with _lock:
        _state["root"] = path
        _state["files"].clear()
        _state["order"].clear()
        _cache.clear()
        _save_settings({"root": path})
    scan()
    return path


# --------------------------------------------------------------------------
# 扫描与清单
# --------------------------------------------------------------------------
def scan(progress=None) -> dict:
    root = get_root()
    paths = iter_files(root)
    files = OrderedDict()
    errors = []
    for i, p in enumerate(paths):
        try:
            st = os.stat(p)
            files[file_id(p)] = FileRef(
                id=file_id(p),
                path=p,
                name=os.path.basename(p),
                fmt=_detect_fmt(p),
                size=st.st_size,
                mtime=st.st_mtime,
                status="ok",
                indexed_at=time.strftime("%Y-%m-%d %H:%M:%S"),
            )
        except OSError as e:
            errors.append({"path": p, "error": str(e)})
        if progress:
            progress(i + 1, len(paths))
    with _lock:
        _state["files"] = files
        _state["order"] = list(files.keys())
        _state["scan_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _state["errors"] = errors
    return {"root": root, "files": len(files), "errors": errors}


def _ensure_scanned():
    if not _state["order"]:
        scan()


def files() -> list[dict]:
    _ensure_scanned()
    with _lock:
        return [f.to_dict() for f in _state["files"].values()]


def path_of(fid: str) -> str:
    _ensure_scanned()
    with _lock:
        f = _state["files"].get(fid)
    if f is None:
        raise ValueError(f"未知文件 id：{fid}")
    return f.path


def file_ref(fid: str) -> FileRef:
    _ensure_scanned()
    with _lock:
        f = _state["files"].get(fid)
    if f is None:
        raise ValueError(f"未知文件 id：{fid}")
    return f


def resolve_file(arg: str) -> str:
    """接受文件 id 或绝对路径，返回绝对路径。"""
    if not arg:
        raise ValueError("缺少文件参数")
    if os.path.isabs(arg) or os.sep in arg or "/" in arg or ":" in arg:
        return os.path.abspath(arg)
    return path_of(arg)


# --------------------------------------------------------------------------
# 读取与缓存
# --------------------------------------------------------------------------
def load(fid: str):
    """读取并识别一个文件，返回 (WorkbookModel, list[TableModel])。"""
    with _lock:
        if fid in _cache:
            _cache.move_to_end(fid)
            return _cache[fid]
    from ..core.detect import detect_workbook
    from ..core.readers import read_workbook

    ref = file_ref(fid)
    wb = read_workbook(ref.path)
    tables = detect_workbook(wb, ref)
    with _lock:
        _cache[fid] = (wb, tables)
        while len(_cache) > 3:
            _cache.popitem(last=False)
    return wb, tables


def invalidate(fid: str | None = None) -> None:
    with _lock:
        if fid is None:
            _cache.clear()
        else:
            _cache.pop(fid, None)


# --------------------------------------------------------------------------
# 工作表 / 预览
# --------------------------------------------------------------------------
def sheets(fid: str) -> list[dict]:
    ref = file_ref(fid)
    wb, tables = load(fid)
    by_sheet = {t.sheet: t for t in tables}
    out = []
    for idx, ws in enumerate(wb.sheets):
        t = by_sheet.get(ws.name)
        out.append({
            "index": idx,
            "name": ws.name,
            "nRows": ws.n_rows,
            "nCols": ws.n_cols,
            "headerRows": [r + 1 for r in (t.header_rows if t else [])],
            "dataStart": (t.data_start + 1) if t else 1,
            "fields": len(t.fields) if t else 0,
            "hidden": ws.hidden,
        })
    return out


def preview(fid: str, sheet_index: int, rows: int = 50) -> dict:
    from ..core.model import col_letter

    ref = file_ref(fid)
    wb, tables = load(fid)
    if sheet_index < 0 or sheet_index >= len(wb.sheets):
        raise ValueError(f"工作表序号越界：{sheet_index}")
    ws = wb.sheets[sheet_index]
    t = next((x for x in tables if x.sheet == ws.name), None)
    max_row = min(ws.n_rows, max(0, rows))
    out_rows = []
    for r0 in range(max_row):
        cells = {}
        for c0 in range(ws.n_cols):
            cell = ws.get(r0, c0)
            if cell is None or cell.raw is None:
                continue
            from ..core.normalize import cell_display

            try:
                text = cell_display(cell)
            except Exception:  # noqa: BLE001
                text = cell.text or str(cell.raw)
            if text:
                cells[col_letter(c0 + 1)] = text
        if cells:
            out_rows.append({"r": r0 + 1, "cells": cells})
    return {
        "fileId": fid,
        "filePath": ref.path,
        "sheetIndex": sheet_index,
        "name": ws.name,
        "nRows": ws.n_rows,
        "nCols": ws.n_cols,
        "merges": [[r1 + 1, r2 + 1, c1 + 1, c2 + 1] for (r1, r2, c1, c2) in ws.merges],
        "headerRows": [r + 1 for r in (t.header_rows if t else [])],
        "dataStart": (t.data_start + 1) if t else 1,
        "rows": out_rows,
    }


# --------------------------------------------------------------------------
# 状态
# --------------------------------------------------------------------------
def status() -> dict:
    _ensure_scanned()
    info = {
        "root": get_root(),
        "files": len(_state["order"]),
        "scannedAt": _state["scan_at"],
        "errors": _state["errors"][:20],
        "indexed": False,
        "cells": 0,
        "indexedAt": "",
    }
    try:
        from . import search_service

        st = search_service.status(root=get_root())
        info["indexed"] = bool(st.get("indexed"))
        info["cells"] = st.get("cells", 0)
        info["indexedAt"] = st.get("indexedAt", "")
        info["sheets"] = st.get("sheets", 0)
    except Exception:  # noqa: BLE001
        pass
    return info


def index_status() -> dict:
    try:
        from . import search_service

        return search_service.status(root=get_root())
    except Exception as e:  # noqa: BLE001
        return {"indexed": False, "error": str(e)}
