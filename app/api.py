# -*- coding: utf-8 -*-
"""HTTP 接口层（Lead 维护，契约冻结）。

每个 handler 都是"薄壳"：解析参数 → 调 service → 包装响应。
service 尚未实现时回退到 ``app/mock.py`` 的示例数据，保证前端可以先行开发。
"""
from __future__ import annotations

import importlib
import os
import subprocess
import sys
import traceback

from . import mock
from .core.model import (
    CONTRACT_VERSION,
    CellRef,
    EditOp,
    MappingProfile,
    MappingRule,
    Sel,
    TemplateDef,
)
from .server import ApiError

DEFAULT_ROOT = os.environ.get(
    "EXCEL_TOOLKIT_ROOT",
    os.path.join(os.path.expanduser("~"), "Documents"),
)


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------
def _service(name: str):
    """按名字取 service 模块；不存在返回 None。"""
    try:
        return importlib.import_module(f"app.services.{name}")
    except ImportError as e:  # 服务未实现 → 回退 mock
        if f"app.services.{name}" in str(e):
            return None
        raise


def _fallback(module, fn: str, *args, **kwargs):
    """优先真服务；未实现则用 mock。"""
    if module is not None and hasattr(module, fn):
        return getattr(module, fn)(*args, **kwargs)
    fn_mock = getattr(mock, fn, None)
    if fn_mock is None:
        raise ApiError("NOT_IMPLEMENTED", f"功能尚未实现：{fn}", 501)
    return fn_mock(*args, **kwargs)


def ok(data=None, **extra):
    out = {"ok": True}
    if data is not None:
        out["data"] = data
    out.update(extra)
    return out


def _root(req) -> str:
    ws = _service("workspace_service")
    if ws is not None and hasattr(ws, "get_root"):
        try:
            return ws.get_root()
        except Exception:  # noqa: BLE001
            pass
    return req.app.root or DEFAULT_ROOT


def _safe_local_path(app, raw: str) -> str:
    """只允许操作工作区、项目目录与 data 目录内的文件（/api/open 用）。"""
    p = os.path.abspath(raw or "")
    roots = [
        os.path.abspath(app.root),
        os.path.abspath(os.path.join(app.root, "data")),
        os.path.abspath(DEFAULT_ROOT),
    ]
    try:
        from .services import workspace_service

        roots.append(os.path.abspath(workspace_service.get_root()))
    except Exception:  # noqa: BLE001
        pass
    if not any(p == r or p.startswith(r + os.sep) for r in roots):
        raise ApiError("FORBIDDEN", "只允许打开工作区内的文件", 403)
    return p


def _task_payload(app, task) -> dict:
    """任务结果里若带生成文件，顺带补一个下载 token（convert/compare 用）。"""
    data = task.to_dict()
    if task.status == "done" and isinstance(task.result, dict):
        result = dict(task.result)
        path = result.get("path")
        if isinstance(path, str) and os.path.isfile(path):
            token = app.tokens_by_task.get(task.id)
            if not token:
                token = app.add_download(path)
                app.tokens_by_task[task.id] = token
            result["token"] = token
            result["name"] = os.path.basename(path)
        data["result"] = result
    return data


# --------------------------------------------------------------------------
# 路由注册
# --------------------------------------------------------------------------
def register(app):
    @app.get(r"/api/health")
    def health(req):
        return ok({
            "status": "up",
            "contract": CONTRACT_VERSION,
            "version": "0.1.0",
            "root": _root(req),
            "uptime": round(__import__("time").time() - req.app.started_at, 1),
        })

    # ---------------- workspace ----------------
    @app.get(r"/api/workspace")
    def get_workspace(req):
        ws = _service("workspace_service")
        data = _fallback(ws, "status")
        data["root"] = _root(req)
        return ok(data)

    @app.post(r"/api/workspace")
    def set_workspace(req):
        body = req.json()
        path = body.get("dir") or body.get("root") or ""
        if not path or not os.path.isdir(path):
            raise ApiError("BAD_DIR", f"目录不存在：{path}", 400)
        ws = _service("workspace_service")
        if ws is None or not hasattr(ws, "set_root"):
            return ok({"root": os.path.abspath(path), "mock": True})
        ws.set_root(path)
        return ok({"root": ws.get_root()})

    @app.get(r"/api/files")
    def list_files(req):
        ws = _service("workspace_service")
        return ok({"files": _fallback(ws, "files"), "root": _root(req)})

    @app.get(r"/api/files/([^/]+)/sheets")
    def file_sheets(req):
        ws = _service("workspace_service")
        return ok({"sheets": _fallback(ws, "sheets", req.group(1))})

    @app.get(r"/api/files/([^/]+)/sheet/(\d+)/preview")
    def sheet_preview(req):
        ws = _service("workspace_service")
        rows = req.q_int("rows", 50)
        return ok(_fallback(ws, "preview", req.group(1), int(req.group(2)), rows))

    @app.get(r"/api/tasks/([^/]+)")
    def task_status(req):
        task = req.app.get_task(req.group(1))
        if task is None:
            raise ApiError("NOT_FOUND", "任务不存在", 404)
        return ok(_task_payload(req.app, task))

    # ---------------- index / search ----------------
    @app.get(r"/api/index/status")
    def index_status(req):
        svc = _service("search_service")
        return ok(_fallback(svc, "status", root=_root(req)))

    @app.post(r"/api/index/refresh")
    def index_refresh(req):
        svc = _service("search_service")
        if svc is not None and hasattr(svc, "start_refresh"):
            task = req.app.start_task("index", svc.start_refresh, req.json() or {}, root=_root(req))
            return ok({"taskId": task.id})
        return ok({"taskId": None, "mock": True})

    @app.get(r"/api/search")
    def search(req):
        q = req.param("q", "") or ""
        svc = _service("search_service")
        result = _fallback(
            svc, "search", q,
            file_id=req.param("file") or None,
            field=req.param("field") or None,
            mode=req.param("mode", "all"),
            limit=req.q_int("limit", 50),
            offset=req.q_int("offset", 0),
            root=_root(req),
        )
        return ok(result)

    @app.post(r"/api/search/export")
    def search_export(req):
        body = req.json()
        svc = _service("search_service")
        out = _fallback(svc, "export", body, root=_root(req), out_dir=os.path.join(req.app.root, "data", "output"))
        token = req.app.add_download(out["path"])
        return ok({"file": out["path"], "name": os.path.basename(out["path"]), "token": token,
                   "count": out.get("count", 0)})

    # ---------------- cell / edit ----------------
    @app.get(r"/api/cell")
    def get_cell(req):
        path = req.param("file", "")
        sheet = req.param("sheet", "")
        row = req.q_int("row")
        col = req.q_int("col")
        if not path or not sheet or not row or not col:
            raise ApiError("BAD_PARAM", "需要 file/sheet/row/col", 400)
        svc = _service("edit_service")
        return ok(_fallback(svc, "get_cell", path, sheet, row, col))

    @app.put(r"/api/cell")
    def put_cell(req):
        body = req.json()
        ref = body.get("ref") or {}
        try:
            cell_ref = CellRef(ref["filePath"], ref["sheet"], int(ref["row"]), int(ref["col"]))
        except (KeyError, TypeError, ValueError):
            raise ApiError("BAD_REF", "缺少合法坐标 ref{filePath,sheet,row,col}", 400)
        op = EditOp(
            ref=cell_ref,
            new_value=body.get("newValue"),
            kind=body.get("kind", "value"),
            expect_mtime=float(body.get("expectMtime") or 0),
            expect_size=int(body.get("expectSize") or 0),
        )
        svc = _service("edit_service")
        return ok(_fallback(svc, "write", op))

    @app.post(r"/api/cells/batch")
    def put_cells_batch(req):
        body = req.json()
        ops = []
        for item in body.get("ops") or []:
            ref = item.get("ref") or {}
            try:
                r = CellRef(ref["filePath"], ref["sheet"], int(ref["row"]), int(ref["col"]))
            except (KeyError, TypeError, ValueError):
                raise ApiError("BAD_REF", "批量写回中存在非法坐标", 400)
            ops.append(EditOp(r, item.get("newValue"), item.get("kind", "value"),
                              float(item.get("expectMtime") or 0), int(item.get("expectSize") or 0)))
        if not ops:
            raise ApiError("BAD_BODY", "ops 不能为空", 400)
        svc = _service("edit_service")
        return ok(_fallback(svc, "write_batch", ops))

    @app.post(r"/api/replace/preview")
    def replace_preview(req):
        svc = _service("edit_service")
        return ok(_fallback(svc, "replace_preview", req.json()))

    @app.post(r"/api/replace/apply")
    def replace_apply(req):
        svc = _service("edit_service")
        return ok(_fallback(svc, "replace_apply", req.json()))

    @app.post(r"/api/undo/([^/]+)")
    def undo(req):
        svc = _service("edit_service")
        return ok(_fallback(svc, "undo", req.group(1)))

    @app.get(r"/api/audit")
    def audit(req):
        svc = _service("edit_service")
        return ok({"entries": _fallback(svc, "audit", req.param("file") or None, req.q_int("limit", 200))})

    @app.post(r"/api/open")
    def open_path(req):
        body = req.json()
        path = _safe_local_path(req.app, body.get("path", ""))
        mode = body.get("mode", "file")
        if not os.path.exists(path):
            raise ApiError("NOT_FOUND", "文件不存在", 404)
        try:
            if mode in ("folder", "select"):
                # 用 ShellExecute 打开资源管理器：比 CreateProcess 更稳，
                # 可避免 explorer.exe 0xc0000142（DLL 初始化失败）
                import ctypes

                if os.path.isdir(path):
                    ctypes.windll.shell32.ShellExecuteW(
                        None, "open", "explorer.exe", f'"{os.path.normpath(path)}"', None, 1)
                else:
                    ctypes.windll.shell32.ShellExecuteW(
                        None, "open", "explorer.exe",
                        f'/select,"{os.path.normpath(path)}"', None, 1)
            elif sys.platform == "win32":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:  # noqa: BLE001
            raise ApiError("OPEN_FAILED", f"无法打开：{e}", 500)
        return ok({"opened": path, "mode": mode})

    @app.post(r"/api/browse")
    def browse_folder(req):
        """弹出系统"选择文件夹"对话框（本地服务所在机器），返回绝对路径。"""
        body = req.json() if isinstance(req.body, dict) else {}
        try:
            from .core.winpick import pick_folder

            path = pick_folder(
                title=body.get("title") or "选择表格所在文件夹",
                initial=body.get("initial") or _root(req),
            )
        except Exception as e:  # noqa: BLE001
            raise ApiError("BROWSE_FAILED", f"无法打开文件夹选择框：{e}", 500)
        return ok({"path": path, "canceled": not path})

    # ---------------- templates / profiles ----------------
    @app.get(r"/api/templates")
    def templates(req):
        svc = _service("mapping_service")
        return ok({"templates": _fallback(svc, "templates")})

    @app.post(r"/api/templates")
    def add_template(req):
        body = req.json()
        svc = _service("mapping_service")
        return ok(_fallback(svc, "add_template", body, root=_root(req)))

    @app.delete(r"/api/templates/([^/]+)")
    def delete_template(req):
        svc = _service("mapping_service")
        return ok(_fallback(svc, "delete_template", req.group(1)))

    @app.post(r"/api/templates/([^/]+)/anchor")
    def template_anchor(req):
        svc = _service("mapping_service")
        return ok(_fallback(svc, "set_anchor", req.group(1), req.json()))

    @app.get(r"/api/profiles")
    def profiles(req):
        svc = _service("mapping_service")
        return ok({"profiles": _fallback(svc, "profiles")})

    @app.post(r"/api/profiles")
    def save_profile(req):
        svc = _service("mapping_service")
        return ok(_fallback(svc, "save_profile", req.json()))

    @app.delete(r"/api/profiles/([^/]+)")
    def delete_profile(req):
        svc = _service("mapping_service")
        return ok(_fallback(svc, "delete_profile", req.group(1)))

    @app.post(r"/api/mapping/suggest")
    def mapping_suggest(req):
        body = req.json()
        svc = _service("mapping_service")
        return ok(_fallback(svc, "suggest_by_ids", body, root=_root(req)))

    @app.post(r"/api/mapping/preview")
    def mapping_preview(req):
        body = req.json()
        svc = _service("mapping_service")
        return ok(_fallback(svc, "preview_by_ids", body, root=_root(req)))

    # ---------------- convert / compare ----------------
    @app.post(r"/api/convert")
    def convert(req):
        body = req.json()
        svc = _service("mapping_service")
        if svc is not None and hasattr(svc, "start_convert"):
            task = req.app.start_task("convert", svc.start_convert, body, root=_root(req))
            return ok({"taskId": task.id})
        return ok(_fallback(svc, "convert", body, root=_root(req)))

    @app.post(r"/api/compare")
    def compare(req):
        body = req.json()
        svc = _service("compare_service")
        data = _fallback(svc, "compare", body, root=_root(req))
        return ok(data)

    @app.post(r"/api/compare/export")
    def compare_export(req):
        body = req.json()
        svc = _service("compare_service")
        out = _fallback(svc, "export", body, root=_root(req),
                        out_dir=os.path.join(req.app.root, "data", "output"))
        token = req.app.add_download(out["path"])
        return ok({"file": out["path"], "name": os.path.basename(out["path"]),
                   "token": token, "count": out.get("count", 0)})

    # ---------------- download ----------------
    @app.get(r"/api/download/([^/]+)")
    def download(req):
        path = req.app.downloads.get(req.group(1))
        if not path or not os.path.isfile(path):
            raise ApiError("NOT_FOUND", "下载已失效，请重新生成", 404)
        name = os.path.basename(path)
        with open(path, "rb") as f:
            data = f.read()
        from .server import FileResponse

        raise FileResponse(data, name)
