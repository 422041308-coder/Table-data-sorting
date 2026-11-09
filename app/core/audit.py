# -*- coding: utf-8 -*-
"""写回审计与撤销（S8，零依赖）。

契约（CONTRACT v1.0 / 第 4 节 S8）：
    backup_file(path, root) -> str          备份到 data/backup/{时间戳}/<相对路径>，返回备份绝对路径
    append_audit(entry) -> ...              追加 data/audit/*.jsonl
    list_audit(file_path=None, limit=200)   读审计（新的在前）
    undo(op_id) -> dict                     按备份整文件回滚

实现要点：
* 备份目录固定在工作区根目录的 ``data/backup/`` 下，一次写回一个时间戳子目录；
* 审计按天分文件 ``data/audit/audit_YYYYMMDD.jsonl``，每行一个 JSON 对象；
* 审计行在 ``AuditEntry.to_dict()`` 基础上额外写 ``id`` / ``filePath`` / ``action`` 字段
  （契约只冻结了 AuditEntry 的字段，追加键不破坏兼容；``undo`` 依赖它们定位备份）；
* ``append_audit`` 返回本条记录的 ``id``（契约标注返回 None，但写回服务需要把 opId 回给前端，
  额外的返回值对原调用方无影响）；
* 所有文件操作都使用 UTF-8；Windows 下路径比较做 ``normcase``。
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import shutil
import threading
import time
import uuid

from .model import AuditEntry

# --------------------------------------------------------------------------
# 路径
# --------------------------------------------------------------------------
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA_DIR = os.path.join(PROJECT_ROOT, "data")
BACKUP_DIR = os.path.join(DATA_DIR, "backup")
AUDIT_DIR = os.path.join(DATA_DIR, "audit")

_LOCK = threading.RLock()


class UndoError(ValueError):
    """撤销失败（找不到记录 / 备份丢失）。带稳定 code 供上层翻译成 ApiError。"""

    code = "UNDO_FAILED"


def _now_ts() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def _norm(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _backup_root(root: str | None) -> str:
    base = os.path.abspath(root or PROJECT_ROOT)
    return os.path.join(base, "data", "backup")


# --------------------------------------------------------------------------
# 备份
# --------------------------------------------------------------------------
def backup_file(path: str, root: str | None = None) -> str:
    """把 ``path`` 备份到 ``{root}/data/backup/{时间戳}/<相对 root 的路径>``，返回备份绝对路径。

    * 文件不存在 → ``FileNotFoundError``；
    * 文件在工作区外 → 放到 ``_external/<随机前缀>_<文件名>``，避免路径穿越；
    * 保留原文件时间戳（``shutil.copy2``），便于审计/比对；
    * 同一秒内的多次备份靠时间戳后的随机后缀区分，不会互相覆盖。
    """
    src = os.path.abspath(path)
    if not os.path.isfile(src):
        raise FileNotFoundError(f"备份失败，文件不存在：{src}")

    base = os.path.abspath(root or PROJECT_ROOT)
    stamp = time.strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    bdir = os.path.join(base, "data", "backup", stamp)

    rel = os.path.relpath(src, base)
    if rel.startswith(".."):  # 工作区外：不保留原目录结构，避免 ../ 穿越
        dest = os.path.join(bdir, "_external", uuid.uuid4().hex[:8] + "_" + os.path.basename(src))
    else:
        dest = os.path.join(bdir, rel)

    os.makedirs(os.path.dirname(dest), exist_ok=True)
    shutil.copy2(src, dest)
    return os.path.abspath(dest)


# --------------------------------------------------------------------------
# 审计
# --------------------------------------------------------------------------
def _append_record(record: dict) -> str:
    """追加一条审计记录（内部用，record 已是可 JSON 化的 dict），返回 id。"""
    rec = dict(record)
    rid = rec.get("id") or uuid.uuid4().hex
    rec["id"] = rid
    os.makedirs(AUDIT_DIR, exist_ok=True)
    fname = "audit_" + _dt.datetime.now().strftime("%Y%m%d") + ".jsonl"
    line = json.dumps(rec, ensure_ascii=False, default=str)
    with _LOCK:
        with open(os.path.join(AUDIT_DIR, fname), "a", encoding="utf-8", newline="\n") as f:
            f.write(line + "\n")
            f.flush()
            try:
                os.fsync(f.fileno())
            except OSError:  # 某些网络盘不支持 fsync，忽略
                pass
    return rid


def append_audit(entry: AuditEntry) -> str:
    """追加一条 AuditEntry，返回记录 id。"""
    if not isinstance(entry, AuditEntry):
        raise TypeError("append_audit 只接受 AuditEntry")
    rec = entry.to_dict()
    fp = ""
    try:
        fp = os.path.abspath(entry.ref.file_path)
    except Exception:  # noqa: BLE001
        fp = entry.ref.file_path or ""
    rec["filePath"] = fp
    return _append_record(rec)


def _iter_records():
    """按“新的在前”遍历所有审计记录（跨天文件按文件名升序再反转）。"""
    if not os.path.isdir(AUDIT_DIR):
        return
    files = sorted(
        (os.path.join(AUDIT_DIR, n) for n in os.listdir(AUDIT_DIR) if n.endswith(".jsonl")),
        reverse=True,
    )
    for fp in files:
        try:
            with open(fp, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except OSError:
            continue
        for line in reversed(lines):
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(rec, dict):
                yield rec


def list_audit(file_path: str | None = None, limit: int = 200) -> list[dict]:
    """读取审计记录（新的在前）。``file_path`` 非空时只返回该文件的记录。"""
    try:
        limit = max(1, int(limit))
    except (TypeError, ValueError):
        limit = 200
    want = _norm(file_path) if file_path else None
    out: list[dict] = []
    for rec in _iter_records():
        if want is not None:
            fp = rec.get("filePath") or (rec.get("ref") or {}).get("filePath") or ""
            if not fp or _norm(fp) != want:
                continue
        out.append(rec)
        if len(out) >= limit:
            break
    return out


# --------------------------------------------------------------------------
# 撤销（按备份整文件回滚）
# --------------------------------------------------------------------------
def undo(op_id: str) -> dict:
    """按操作 id 找到备份并整文件还原，同时再追加一条"撤销"审计。

    返回 ``{"opId", "undoOf", "ref": {...}, "filePath", "backupPath", "result": "ok"}``；
    找不到记录/备份 → ``UndoError``。
    """
    op_id = str(op_id or "").strip()
    if not op_id:
        raise UndoError("缺少操作 id，无法撤销")

    for rec in _iter_records():
        rid = str(rec.get("id") or "")
        if not rid:
            continue
        # 支持完整 id 与唯一前缀（前端 token 截断场景）
        if not (rid == op_id or (len(op_id) >= 8 and rid.startswith(op_id))):
            continue
        if rec.get("action") == "undo":
            raise UndoError("该记录是撤销操作本身，不能再撤销")
        if rec.get("result") not in (None, "ok"):
            raise UndoError("该操作未成功完成，不能撤销")

        backup_path = rec.get("backupPath") or rec.get("backup_path") or ""
        if not backup_path or not os.path.isfile(backup_path):
            raise UndoError(f"备份文件不存在，无法撤销：{backup_path or '（未记录）'}")

        ref_data = rec.get("ref") or {}
        file_path = rec.get("filePath") or ref_data.get("filePath") or ""
        if not file_path:
            raise UndoError("审计记录缺少原文件路径，无法撤销")
        file_path = os.path.abspath(file_path)

        # 目标目录可能存在也可能被删；还原时补齐目录
        os.makedirs(os.path.dirname(file_path) or ".", exist_ok=True)
        tmp = file_path + f".undo-{uuid.uuid4().hex[:8]}"
        mode = None
        try:
            shutil.copy2(backup_path, tmp)
            # 目标若只读，替换会失败：临时放开写权限，替换后恢复
            if os.path.exists(file_path):
                mode = os.stat(file_path).st_mode
                if not (mode & 0o200):
                    os.chmod(file_path, mode | 0o200)
            os.replace(tmp, file_path)
        finally:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            if mode is not None and not (mode & 0o200):
                try:
                    os.chmod(file_path, mode)
                except OSError:
                    pass

        # 再记一条撤销审计（旧值=被撤销后的值，新值=恢复后的值）
        try:
            from .model import CellRef

            ref = CellRef(
                file_path,
                str(ref_data.get("sheet") or ""),
                int(ref_data.get("row") or 0),
                int(ref_data.get("col") or 0),
            )
            entry = AuditEntry(
                ts=_now_ts(),
                ref=ref,
                old_value=str(rec.get("newValue") or ""),
                new_value=str(rec.get("oldValue") or ""),
                backup_path=backup_path,
                result="ok",
            )
            new_rec = entry.to_dict()
            new_rec["filePath"] = file_path
            new_rec["action"] = "undo"
            new_rec["undoOf"] = rid
            new_id = _append_record(new_rec)
        except Exception:  # noqa: BLE001 - 审计失败不应让已完成的还原"看起来失败"
            new_id = ""
            ref = None

        return {
            "opId": new_id,
            "undoOf": rid,
            "ref": ref.to_dict() if ref is not None else {},
            "filePath": file_path,
            "backupPath": backup_path,
            "result": "ok",
        }

    raise UndoError(f"未找到可撤销的操作记录：{op_id}")


__all__ = [
    "PROJECT_ROOT",
    "DATA_DIR",
    "BACKUP_DIR",
    "AUDIT_DIR",
    "UndoError",
    "backup_file",
    "append_audit",
    "list_audit",
    "undo",
]
