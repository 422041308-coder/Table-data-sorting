# -*- coding: utf-8 -*-
"""检索服务：索引生命周期 + 查询 + 导出（Lead 维护）。"""
from __future__ import annotations

import os
import threading
import time

from ..core.index import SearchIndex
from ..core.model import FileRef

_state = {
    "root": None,
    "index": None,
    "indexed_at": "",
    "stats": {},
    "errors": [],
}
_lock = threading.RLock()


# --------------------------------------------------------------------------
# 构建
# --------------------------------------------------------------------------
def _build(root: str, progress=None) -> dict:
    from ..core.detect import detect_workbook
    from ..core.readers import read_workbook
    from .workspace_service import file_id, iter_files

    t0 = time.time()
    index = SearchIndex()
    errors: list[dict] = []
    paths = iter_files(root)
    total = len(paths)
    ok = 0
    for i, path in enumerate(paths):
        try:
            wb = read_workbook(path)
            st = os.stat(path)
            ref = FileRef(id=file_id(path), path=path, name=os.path.basename(path),
                          fmt=wb.fmt, size=st.st_size, mtime=st.st_mtime)
            tables = detect_workbook(wb, ref)
            for t in tables:
                if t.rows:
                    index.add(t)
            ok += 1
        except Exception as e:  # noqa: BLE001
            errors.append({"path": path, "error": f"{type(e).__name__}: {e}"})
        if progress:
            progress(i + 1, total)
    took = round(time.time() - t0, 2)
    stats = index.size()
    with _lock:
        _state["root"] = os.path.abspath(root)
        _state["index"] = index
        _state["indexed_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _state["stats"] = {"files": ok, "failed": len(errors), "took": took, **stats}
        _state["errors"] = errors
    return {"root": os.path.abspath(root), "indexedAt": _state["indexed_at"],
            "files": ok, "failed": len(errors), "took": took, "errors": errors[:50], **stats}


def ensure_index(root: str | None = None) -> SearchIndex:
    from .workspace_service import get_root

    root = os.path.abspath(root or get_root())
    with _lock:
        if _state["index"] is not None and _state["root"] == root:
            return _state["index"]
    _build(root)
    with _lock:
        return _state["index"]


def invalidate() -> None:
    with _lock:
        _state["index"] = None
        _state["root"] = None
        _state["stats"] = {}
        _state["errors"] = []
        _state["indexed_at"] = ""


def start_refresh(task, body: dict | None = None, root: str | None = None) -> dict:
    """长任务入口：重建索引，实时回报进度。"""
    from .workspace_service import get_root

    root = os.path.abspath(root or get_root())

    def progress(done: int, total: int):
        task.progress = int(done * 100 / total) if total else 100
        task.message = f"已解析 {done}/{total} 个文件"

    task.message = "开始扫描…"
    return _build(root, progress)


def status(root: str | None = None) -> dict:
    from .workspace_service import get_root

    root = os.path.abspath(root or get_root())
    with _lock:
        indexed = _state["index"] is not None and _state["root"] == root
        stats = dict(_state["stats"]) if indexed else {}
        indexed_at = _state["indexed_at"] if indexed else ""
        errors = list(_state["errors"]) if indexed else []
    return {
        "root": root,
        "indexed": indexed,
        "files": stats.get("files", 0),
        "rows": stats.get("rows", 0),
        "cells": stats.get("cells", 0),
        "sheets": stats.get("tables", 0),
        "tokens": stats.get("tokens", 0),
        "took": stats.get("took", 0),
        "indexedAt": indexed_at,
        "failed": stats.get("failed", 0),
        "errors": errors[:20],
    }


# --------------------------------------------------------------------------
# 查询
# --------------------------------------------------------------------------
def search(q: str, file_id: str | None = None, field: str | None = None,
           mode: str = "all", limit: int = 50, offset: int = 0,
           root: str | None = None) -> dict:
    from .workspace_service import resolve_file

    t0 = time.time()
    index = ensure_index(root)
    path_filter = None
    if file_id:
        try:
            path_filter = resolve_file(file_id)
        except Exception:  # noqa: BLE001
            path_filter = None
    hits, total = index.search(q, limit=limit, offset=offset, mode=mode or "all",
                               field=field or None, path_filter=path_filter)
    return {
        "total": total,
        "took": round(time.time() - t0, 4),
        "query": q,
        "mode": mode,
        "hits": [h.to_dict() for h in hits],
        "index": index.size(),
    }


def update_cell(ref, text: str, field_name: str = "") -> bool:
    with _lock:
        index = _state["index"]
    if index is None:
        return False
    return index.update_cell(ref, text, field_name)


# --------------------------------------------------------------------------
# 导出
# --------------------------------------------------------------------------
def export(body: dict, root: str | None = None, out_dir: str | None = None) -> dict:
    from ..core.xlsxwriter import auto_widths, write_rows_xlsx
    from .workspace_service import PROJECT_ROOT

    body = body or {}
    q = body.get("query") or body.get("q") or ""
    svc = search(q, file_id=body.get("file"), field=body.get("field"),
                 mode=body.get("mode") or "all", limit=100000, offset=0, root=root)
    hits = svc.get("hits", [])
    rows = []
    for h in hits:
        ref = h["ref"]
        rows.append([
            ref["filePath"], ref["sheet"], ref["row"], ref["col"],
            ref["colLetter"], ref["a1"], h.get("fieldName", ""), h.get("valueText", ""),
        ])
    out_dir = out_dir or os.path.join(PROJECT_ROOT, "data", "output")
    os.makedirs(out_dir, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    safe = "".join(ch for ch in (q or "全部") if ch not in r'\/:*?"<>|')[:24] or "全部"
    path = os.path.join(out_dir, f"检索结果_{safe}_{stamp}.xlsx")
    write_rows_xlsx(
        path, rows,
        header=["文件地址", "工作表", "第几行", "第几列", "列字母", "A1", "字段名", "命中内容"],
        sheet_name="检索结果",
        widths=auto_widths(rows, ["文件地址", "工作表", "第几行", "第几列", "列字母", "A1", "字段名", "命中内容"]),
    )
    return {"path": path, "count": len(rows), "query": q}
