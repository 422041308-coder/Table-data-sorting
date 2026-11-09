# -*- coding: utf-8 -*-
"""编辑写回服务（S8）：get_cell / write / write_batch / replace_preview / replace_apply / undo / audit。

安全链路（顺序固定，见 docs/CONTRACT.md 第 4 节 S8）：
    冲突检测（磁盘 mtime/size vs expect_*）
        → 备份 data/backup/{时间戳}/<相对路径>
        → 最小改动写回（app/core/writers/ooxml_writer.py）
        → 重新解析自检（目标格=新值、工作表数不变、非目标条目字节不变）
        → 失败自动回滚并抛错
        → 审计追加 data/audit/*.jsonl

补充约定（CONTRACT 5.1）：
* 文件既可用绝对路径（filePath），也可用 ``fileId``（sha1(路径)[:12]，由 workspace_service 统一生成）；
  ``fileId`` 通过 ``workspace_service.resolve_file / files`` 反查，反查不到时要求传 filePath；
* replace/preview 返回 ``{matches, count, truncated}``；replace/apply 只应用前端确认的 ``items``，
  返回 ``{written, results}``（内部复用 write_batch：同文件一次备份、一次落盘）。
"""
from __future__ import annotations

import math
import os
import re
import threading

from ..core import audit as audit_mod
from ..core.model import AuditEntry, CellRef, EditOp, split_a1
from ..core.writers import ooxml_writer as writer
from ..server import ApiError

# app/services/edit_service.py → services → app → 项目根
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

_LOCK = threading.RLock()
_VALID_KINDS = ("value", "formula", "clear")


# --------------------------------------------------------------------------
# 工具
# --------------------------------------------------------------------------
def _now() -> str:
    import datetime as _dt

    return _dt.datetime.now().isoformat(timespec="seconds")


def _norm(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _format_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if math.isfinite(value) and value.is_integer() and abs(value) < 1e16:
            return str(int(value))
        return repr(value)
    return str(value)


def _jsonable(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _resolve_path(value) -> str:
    """把绝对路径 / fileId / 工作区内相对路径解析成存在的绝对路径；失败返回空串。"""
    v = str(value or "").strip()
    if not v:
        return ""
    if os.path.isfile(v):
        return os.path.abspath(v)

    # 1) 交给 workspace_service 统一解析（CONTRACT 5.1）
    try:
        from . import workspace_service  # type: ignore

        fn = getattr(workspace_service, "resolve_file", None)
        if callable(fn):
            try:
                p = fn(v)
            except Exception:  # noqa: BLE001 - 服务内部异常不影响后续兜底
                p = None
            if p and os.path.isfile(p):
                return os.path.abspath(p)
        files_fn = getattr(workspace_service, "files", None)
        if callable(files_fn):
            try:
                items = files_fn() or []
            except Exception:  # noqa: BLE001
                items = []
            for item in items:
                if not isinstance(item, dict):
                    continue
                if str(item.get("id") or "") == v:
                    p = item.get("path") or ""
                    if p and os.path.isfile(p):
                        return os.path.abspath(p)
                if item.get("path") and _norm(item["path"]) == _norm(v):
                    return os.path.abspath(item["path"])
        root_fn = getattr(workspace_service, "get_root", None)
        if callable(root_fn):
            try:
                root = root_fn() or ""
            except Exception:  # noqa: BLE001
                root = ""
            if root:
                cand = os.path.join(root, v)
                if os.path.isfile(cand):
                    return os.path.abspath(cand)
    except Exception:  # noqa: BLE001 - S1/S7 服务未就绪
        pass
    return ""


def _require_file(value) -> str:
    p = _resolve_path(value)
    if not p:
        raw = str(value or "")
        if raw and os.path.isabs(raw):
            raise ApiError("NOT_FOUND", f"文件不存在：{raw}", 404)
        raise ApiError(
            "BAD_PARAM",
            "无法定位文件：请传 filePath（绝对路径）或可解析的 fileId",
            400,
            detail={"value": raw},
        )
    return p


def _ensure_supported(path: str) -> None:
    try:
        writer.ensure_supported(path)
    except ApiError:
        raise
    except Exception as exc:  # 后备 UnsupportedFormatError（app.server 不可导入时）
        code = getattr(exc, "code", None)
        if code == "UNSUPPORTED_FORMAT":
            raise ApiError("UNSUPPORTED_FORMAT", str(exc), 400) from exc
        raise


def _stat_current(path: str) -> dict:
    st = os.stat(path)
    return {"mtime": st.st_mtime, "size": st.st_size}


def _check_conflicts(ops: list[EditOp], st) -> None:
    for op in ops:
        if op.expect_mtime:
            if abs(st.st_mtime - float(op.expect_mtime)) > 1e-6:
                raise ApiError(
                    "CONFLICT",
                    "文件已被修改（修改时间与打开时不一致），请刷新后重试",
                    409,
                    detail={
                        "path": op.ref.file_path,
                        "expectedMtime": float(op.expect_mtime),
                        "actualMtime": st.st_mtime,
                    },
                )
        if op.expect_size:
            if int(st.st_size) != int(op.expect_size):
                raise ApiError(
                    "CONFLICT",
                    "文件已被修改（大小与打开时不一致），请刷新后重试",
                    409,
                    detail={
                        "path": op.ref.file_path,
                        "expectedSize": int(op.expect_size),
                        "actualSize": int(st.st_size),
                    },
                )


def _new_value_text(op: EditOp) -> str:
    if op.kind == "clear" or op.new_value is None:
        return ""
    if op.kind == "formula":
        body = str(op.new_value).strip()
        if body.startswith("="):
            body = body[1:]
        return "=" + body
    return _format_value(op.new_value)


def _ensure_op(op) -> EditOp:
    if isinstance(op, EditOp):
        return op
    if isinstance(op, dict):
        ref = op.get("ref") or {}
        if not isinstance(ref, dict):
            raise ApiError("BAD_REF", "写回项 ref 非法", 400)
        try:
            file_path = ref.get("filePath") or ref.get("file_path") or ref.get("path") or ""
            row, col = ref.get("row"), ref.get("col")
            if (row in (None, "") or col in (None, "")) and ref.get("a1"):
                row, col = split_a1(ref["a1"])
            cell_ref = CellRef(
                os.path.abspath(file_path) if file_path else "",
                str(ref.get("sheet") or ""),
                int(row),
                int(col),
            )
        except (TypeError, ValueError, KeyError) as exc:
            raise ApiError("BAD_REF", "缺少合法坐标 ref{filePath,sheet,row,col}", 400) from exc
        return EditOp(
            cell_ref,
            op.get("newValue"),
            str(op.get("kind") or "value"),
            float(op.get("expectMtime") or 0),
            int(op.get("expectSize") or 0),
        )
    raise ApiError("BAD_BODY", "写回项必须是 EditOp 或对象", 400)


def _read_old_cells(path: str, ops: list[EditOp]) -> dict:
    try:
        return writer.read_cells(path, [op.ref for op in ops])
    except ValueError as exc:
        # read_cells 对“工作表不存在”抛 ValueError；合并区校验留给下面单独抛 ValueError
        if "工作表不存在" in str(exc):
            raise ApiError("NOT_FOUND", str(exc), 404) from exc
        raise


def _append_audits(ops: list[EditOp], old_map: dict, backup_path: str, result: str) -> list[str]:
    ids: list[str] = []
    for op in ops:
        old_info = old_map.get((op.ref.sheet, op.ref.a1)) or {}
        entry = AuditEntry(
            ts=_now(),
            ref=op.ref,
            old_value=str(old_info.get("text") or ""),
            new_value=_new_value_text(op),
            backup_path=backup_path,
            result=result,
        )
        try:
            ids.append(audit_mod.append_audit(entry))
        except Exception:  # noqa: BLE001 - 审计失败不应把已成功的写回变成失败
            ids.append("")
    return ids


# --------------------------------------------------------------------------
# 冲突异常 → ApiError
# --------------------------------------------------------------------------
def _conflict_from(exc: writer.ConflictError) -> ApiError:
    return ApiError(
        "CONFLICT",
        str(exc) or "文件已被修改，请刷新后重试",
        409,
        detail={
            "path": exc.path,
            "expectedMtime": exc.expected_mtime,
            "actualMtime": exc.actual_mtime,
            "expectedSize": exc.expected_size,
            "actualSize": exc.actual_size,
        },
    )


# --------------------------------------------------------------------------
# get_cell
# --------------------------------------------------------------------------
def _detect_meta(file_path: str, sheet: str, row: int, col: int) -> tuple[str, bool]:
    """字段名 / 是否小计行：优先取 S2（app.core.detect）的识别结果。

    S2 可能尚未就绪（或读文件失败），此时返回 ("", False) —— 不影响编辑功能。
    实现方式：read_workbook + detect_workbook，找到包含该坐标的表，
    取 ``TableModel.field_at(col-1).name`` 与 ``excluded_rows`` 中的“小计/合计”行。
    """
    try:
        from ..core.readers import read_workbook
        from ..core.detect import detect_workbook
    except Exception:  # noqa: BLE001 - S2 未就绪
        return "", False
    try:
        wb = read_workbook(file_path)
        tables = detect_workbook(wb) or []
    except Exception:  # noqa: BLE001 - 识别失败不应阻塞编辑
        return "", False

    r0, c0 = int(row) - 1, int(col) - 1
    for table in tables:
        if getattr(table, "sheet", None) != sheet:
            continue
        header_rows = set(getattr(table, "header_rows", []) or [])
        data_start = getattr(table, "data_start", 0)
        data_end = getattr(table, "data_end", -1)
        in_data = data_start <= r0 <= (data_end if data_end is not None and data_end >= 0 else r0)
        if not in_data and r0 not in header_rows:
            continue
        field = table.field_at(c0) if hasattr(table, "field_at") else None
        name = getattr(field, "name", "") if field is not None else ""
        is_subtotal = False
        for item in getattr(table, "excluded_rows", []) or []:
            try:
                rr, reason = item
            except (TypeError, ValueError):
                continue
            if int(rr) == r0 and ("小计" in str(reason) or "合计" in str(reason)):
                is_subtotal = True
                break
        return name, is_subtotal
    return "", False


def get_cell(file_path, sheet, row, col) -> dict:
    """读取单元格详情（GET /api/cell，形状见 CONTRACT 第 5 节）。"""
    p = _require_file(file_path)
    _ensure_supported(p)
    try:
        row_i, col_i = int(row), int(col)
    except (TypeError, ValueError):
        raise ApiError("BAD_PARAM", "row/col 需要整数", 400)
    if row_i < 1 or col_i < 1:
        raise ApiError("BAD_PARAM", "row/col 是 1 基，必须 >= 1", 400)
    try:
        info = writer.read_cell(p, str(sheet), row_i, col_i)
    except ValueError as exc:
        if "工作表不存在" in str(exc):
            raise ApiError("NOT_FOUND", str(exc), 404) from exc
        raise
    field_name, is_subtotal = _detect_meta(p, str(sheet), row_i, col_i)
    ref = CellRef(p, str(sheet), row_i, col_i)
    return {
        "ref": ref.to_dict(),
        "value": _jsonable(info["value"]),
        "kind": info["kind"],
        "formula": info["formula"],
        "display": info["text"],
        "fieldName": field_name,
        "isSubtotal": is_subtotal,
        "mergeOwner": info["mergeOwner"],
        "editable": info["editable"],
        "current": _stat_current(p),
    }


# --------------------------------------------------------------------------
# 写回
# --------------------------------------------------------------------------
def _apply_ops(ops: list[EditOp]) -> dict:
    """同一文件的一批写回：冲突 → 备份 → 写 → 自检/回滚 → 审计。"""
    if not ops:
        raise ApiError("BAD_BODY", "ops 不能为空", 400)

    # 1) 参数校验
    normalized: list[EditOp] = []
    for op in ops:
        op = _ensure_op(op)
        op.kind = str(op.kind or "value").strip().lower()
        if op.kind not in _VALID_KINDS:
            raise ApiError("BAD_PARAM", f"未知写回类型：{op.kind!r}（只支持 value/formula/clear）", 400)
        if not op.ref.sheet:
            raise ApiError("BAD_REF", "缺少工作表名 sheet", 400)
        if int(op.ref.row) < 1 or int(op.ref.col) < 1:
            raise ApiError("BAD_REF", "坐标是 1 基，row/col 必须 >= 1", 400)
        normalized.append(op)

    # 2) 同一文件
    paths = {_norm(op.ref.file_path) for op in normalized if op.ref.file_path}
    if len(paths) > 1:
        raise ApiError("BAD_PARAM", "批量写回只支持同一文件；请按文件分组", 400)
    p = _require_file(normalized[0].ref.file_path)
    _ensure_supported(p)
    for op in normalized:
        op.ref.file_path = p

    with _LOCK:
        # 3) 冲突检测（磁盘 mtime/size）
        st = os.stat(p)
        _check_conflicts(normalized, st)

        # 4) 读旧值 + 合并区校验（写回只允许合并区左上角）
        old_map = _read_old_cells(p, normalized)
        for op in normalized:
            info = old_map.get((op.ref.sheet, op.ref.a1))
            if info is None:
                raise ApiError("BAD_REF", f"无法定位单元格：{op.ref.a1}", 400)
            if not info.get("editable", True):
                raise ValueError("合并区，请编辑左上角单元格")

        # 5) 备份
        try:
            backup = audit_mod.backup_file(p, PROJECT_ROOT)
        except OSError as exc:
            raise ApiError("BACKUP_FAILED", f"备份失败，已取消写回：{exc}", 500) from exc

        # 6) 最小改动写回（内部还会做一次冲突复核 + 自检 + 临时文件原子替换）
        try:
            writer.write_cells(p, normalized, backup_path=backup, root=PROJECT_ROOT)
        except writer.ConflictError as exc:
            _append_audits(normalized, old_map, backup, "conflict")
            raise _conflict_from(exc) from exc
        except writer.WriteError as exc:
            _append_audits(normalized, old_map, backup, "rolled_back")
            raise ApiError(
                "WRITE_FAILED",
                f"写回失败，已自动回滚：{exc}",
                500,
                detail={"backupPath": backup},
            ) from exc

        # 7) 成功审计
        ids = _append_audits(normalized, old_map, backup, "ok")
        st2 = os.stat(p)
        results = []
        for idx, op in enumerate(normalized):
            old_info = old_map.get((op.ref.sheet, op.ref.a1)) or {}
            results.append({
                "ref": op.ref.to_dict(),
                "result": "ok",
                "opId": ids[idx],
                "oldValue": str(old_info.get("text") or ""),
                "newValue": _new_value_text(op),
            })
        return {
            "result": "ok",
            "count": len(normalized),
            "written": len(normalized),
            "backupPath": backup,
            "opIds": ids,
            "results": results,
            "items": results,
            "current": {"mtime": st2.st_mtime, "size": st2.st_size},
        }


def write(op: EditOp) -> dict:
    """写一个单元格。冲突 → ``ApiError("CONFLICT", ..., 409)``；合并区非左上角 → ValueError。"""
    res = _apply_ops([op])
    item = res["results"][0]
    return {
        "ref": item["ref"],
        "oldValue": item["oldValue"],
        "newValue": item["newValue"],
        "result": "ok",
        "opId": item["opId"],
        "written": 1,
        "count": 1,
        "backupPath": res["backupPath"],
        "current": res["current"],
    }


def write_batch(ops: list[EditOp]) -> dict:
    """同文件多格写回：一次备份、一次落盘；跨文件时按文件分组各写一次。"""
    if not ops:
        raise ApiError("BAD_BODY", "ops 不能为空", 400)
    normalized = [_ensure_op(op) for op in ops]
    # 先把 fileId / 相对路径统一解析成绝对路径，再分组，避免同一文件被拆成多次备份
    resolved: dict[str, str] = {}
    for op in normalized:
        raw = op.ref.file_path
        key = _norm(raw) if raw else ""
        if key not in resolved:
            p = _require_file(raw)
            _ensure_supported(p)
            resolved[key] = p
        op.ref.file_path = resolved[key]

    groups: dict[str, list[EditOp]] = {}
    order: list[str] = []
    for op in normalized:
        key = _norm(op.ref.file_path)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(op)

    all_results: list[dict] = []
    all_ids: list[str] = []
    backups: list[str] = []
    total = 0
    current = None
    for key in order:
        res = _apply_ops(groups[key])
        total += int(res["written"])
        all_results.extend(res["results"])
        all_ids.extend(res["opIds"])
        if res.get("backupPath"):
            backups.append(res["backupPath"])
        current = res.get("current") or current
    return {
        "result": "ok",
        "count": total,
        "written": total,
        "results": all_results,
        "items": all_results,
        "opIds": all_ids,
        "backupPath": backups[0] if len(backups) == 1 else "",
        "backups": backups,
        "current": current,
    }


# --------------------------------------------------------------------------
# 查找替换（CONTRACT 5.1）
# --------------------------------------------------------------------------
def _replace_ci(text: str, find: str, replace: str) -> str:
    return re.sub(re.escape(find), lambda _m: replace, text, flags=re.IGNORECASE)


def _replace_in(text: str, find: str, replace: str, match_case: bool) -> str:
    if match_case:
        return text.replace(find, replace)
    return _replace_ci(text, find, replace)


def _scope_files(body: dict) -> tuple[list[str], str]:
    """查找替换的作用域：指定文件（filePath/fileId）或整个工作区。

    返回 ``(可写文件绝对路径列表, 单文件路径或空串)``。
    """
    raw_path = body.get("filePath") or body.get("file") or body.get("path") or ""
    if raw_path:
        p = _resolve_path(raw_path)
        if not p:
            raise ApiError("NOT_FOUND", f"文件不存在：{raw_path}", 404)
        return [p], p
    file_id = body.get("fileId") or body.get("file_id") or ""
    if file_id:
        p = _resolve_path(file_id)
        if not p:
            raise ApiError("NOT_FOUND", f"无法解析 fileId：{file_id}", 404)
        return [p], p
    # 未指定文件：按前端提示的“范围：工作区全部文件”扫描（跳过不支持/读不了的文件）
    try:
        from . import workspace_service  # type: ignore

        entries = workspace_service.files() or []
    except Exception:  # noqa: BLE001 - 工作区服务未就绪
        entries = []
    paths: list[str] = []
    seen: set[str] = set()
    for item in entries:
        if not isinstance(item, dict):
            continue
        if item.get("status") not in (None, "ok"):
            continue
        p = item.get("path") or ""
        if not p or not os.path.isfile(p):
            continue
        key = _norm(p)
        if key in seen:
            continue
        seen.add(key)
        paths.append(os.path.abspath(p))
    return paths, ""


def replace_preview(body: dict) -> dict:
    """查找替换预览（只读，不写盘）：返回 {matches, count, truncated}。"""
    body = body or {}
    find = str(body.get("find") or body.get("query") or "")
    if not find:
        raise ApiError("BAD_PARAM", "find 不能为空", 400)
    replace = body.get("replace")
    replace = "" if replace is None else str(replace)
    match_case = bool(body.get("matchCase") if body.get("matchCase") is not None
                      else body.get("caseSensitive"))
    try:
        limit = int(body.get("limit") or 200)
    except (TypeError, ValueError):
        limit = 200
    limit = max(1, min(limit, 2000))

    files, single = _scope_files(body)
    if not files:
        raise ApiError("BAD_PARAM", "没有可替换的表格文件：请先选择文件或刷新工作区", 400)
    sheet = body.get("sheet") or None

    needle = find if match_case else find.casefold()
    matches: list[dict] = []
    truncated = False
    scanned_files = 0
    scanned_cells = 0
    for p in files:
        try:
            _ensure_supported(p)
            cells = writer.scan_cells(p, sheet)
        except (ApiError, ValueError, OSError, writer.WriteError):
            # 全工作区模式下跳过 xls/损坏/无权限文件；指定单文件时仍提示错误
            if single:
                raise
            continue
        scanned_files += 1
        scanned_cells += len(cells)
        for info in cells:
            text = str(info.get("text") or "")
            if not text:
                continue
            hay = text if match_case else text.casefold()
            if needle not in hay:
                continue
            if not info.get("editable", True):
                continue  # 合并区非左上角不可编辑
            formula = info.get("formula")
            # 公式文本本身命中时替换公式；否则替换展示文本
            if formula and needle in ((formula if match_case else formula.casefold())):
                old_value = formula
                new_value = _replace_in(formula, find, replace, match_case)
                kind = "formula"
            else:
                old_value = text
                new_value = _replace_in(text, find, replace, match_case)
                kind = "value"
            ref = CellRef(p, info["sheet"], int(info["row"]), int(info["col"]))
            matches.append({
                "ref": ref.to_dict(),
                "oldValue": old_value,
                "newValue": new_value,
                "sheet": info["sheet"],
                "kind": kind,
            })
            if len(matches) >= limit:
                truncated = True
                break
        if truncated:
            break
    return {
        "matches": matches,
        "items": matches,
        "count": len(matches),
        "truncated": truncated,
        "filePath": single or None,
        "files": scanned_files,
        "scanned": scanned_cells,
        "find": find,
        "replace": replace,
    }


def _cellref_from_item(ref: dict, fallback_path: str = "") -> CellRef:
    if not isinstance(ref, dict):
        raise ApiError("BAD_REF", "items[].ref 必须是对象", 400)
    file_path = (
        ref.get("filePath") or ref.get("file_path") or ref.get("path") or fallback_path
    )
    sheet = ref.get("sheet") or ref.get("sheetName") or ""
    row, col = ref.get("row"), ref.get("col")
    if (row in (None, "") or col in (None, "")) and ref.get("a1"):
        try:
            row, col = split_a1(ref["a1"])
        except ValueError as exc:
            raise ApiError("BAD_REF", f"非法 A1 引用：{ref.get('a1')}", 400) from exc
    if not file_path or not sheet:
        raise ApiError("BAD_REF", "items[].ref 缺少 filePath 或 sheet", 400)
    try:
        return CellRef(os.path.abspath(file_path), str(sheet), int(row), int(col))
    except (TypeError, ValueError) as exc:
        raise ApiError("BAD_REF", "items[].ref 缺少合法 row/col", 400) from exc


def replace_apply(body: dict) -> dict:
    """应用查找替换（只写用户确认过的 items）；返回 {written, results}。"""
    body = body or {}
    items = body.get("items")
    try:
        expect_mtime = float(body.get("expectMtime") or 0)
        expect_size = int(body.get("expectSize") or 0)
    except (TypeError, ValueError):
        raise ApiError("BAD_PARAM", "expectMtime/expectSize 类型错误", 400)

    raw_path = body.get("filePath") or body.get("file") or body.get("path") or ""
    fallback_path = _resolve_path(raw_path) if raw_path else ""
    if not fallback_path:
        fallback_path = _resolve_path(body.get("fileId") or body.get("file_id") or "")

    ops: list[EditOp] = []
    if isinstance(items, list):
        # 传了 items 就只应用用户确认过的条目（空数组表示"没有要应用的"，直接返回 0）
        for item in items:
            if not isinstance(item, dict):
                raise ApiError("BAD_BODY", "items 元素必须是对象", 400)
            ref = _cellref_from_item(item.get("ref") or {}, fallback_path)
            kind = str(item.get("kind") or "value").strip().lower()
            if kind not in ("value", "formula"):
                kind = "value"
            new_value = item.get("newValue", item.get("value"))
            ops.append(EditOp(ref, new_value, kind, expect_mtime, expect_size))
    else:
        # 兜底：没传 items 时按 find/replace 重新扫描（便于脚本/测试调用）
        scanned = replace_preview(body)
        for m in scanned["matches"]:
            ref = CellRef(
                str(m["ref"]["filePath"]),
                str(m["ref"]["sheet"]),
                int(m["ref"]["row"]),
                int(m["ref"]["col"]),
            )
            kind = "formula" if m.get("kind") == "formula" else "value"
            ops.append(EditOp(ref, m["newValue"], kind, expect_mtime, expect_size))

    if not ops:
        return {"result": "ok", "written": 0, "count": 0, "results": [], "opIds": []}

    res = write_batch(ops)
    results = [
        {
            "ref": item["ref"],
            "result": item["result"],
            "opId": item["opId"],
            "oldValue": item["oldValue"],
            "newValue": item["newValue"],
        }
        for item in res["results"]
    ]
    return {
        "result": "ok",
        "written": int(res["written"]),
        "count": int(res["count"]),
        "results": results,
        "opIds": res["opIds"],
        "backupPath": res.get("backupPath", ""),
        "backups": res.get("backups", []),
        "current": res.get("current"),
    }


# --------------------------------------------------------------------------
# 撤销 / 审计
# --------------------------------------------------------------------------
def undo(op_id: str) -> dict:
    try:
        return audit_mod.undo(op_id)
    except audit_mod.UndoError as exc:
        raise ApiError("UNDO_FAILED", str(exc), 404) from exc
    except OSError as exc:
        raise ApiError("UNDO_FAILED", f"撤销失败：{exc}", 500) from exc


def audit(file=None, limit: int = 200) -> list[dict]:
    """审计记录（新的在前）。"""
    return audit_mod.list_audit(file, limit)


__all__ = [
    "get_cell",
    "write",
    "write_batch",
    "replace_preview",
    "replace_apply",
    "undo",
    "audit",
]
