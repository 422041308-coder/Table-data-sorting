# -*- coding: utf-8 -*-
"""格式嗅探：按魔数 + 容器内容判断表格文件真实格式。

返回值为 ``app.core.model`` 里的格式常量：

* ``"xlsx"`` / ``"xlsm"`` —— OOXML zip（宏工作簿）
* ``"xlsb"``              —— OOXML zip 里的二进制工作簿（``xl/workbook.bin``）
* ``"xls"``               —— OLE2 复合文档里的 BIFF 工作簿
* ``"csv"``               —— 文本（UTF-8 / GBK / UTF-16 等编码嗅探）
* ``"et-zip"``            —— WPS 表格的 zip 版（OOXML 结构，扩展名 .et）
* ``"et-ole"``            —— WPS 表格的 OLE 版（BIFF 结构，扩展名 .et）
* ``"unknown"``           —— 不是本工具支持的表格
"""
from __future__ import annotations

import os
import zipfile

from .model import (
    FMT_CSV,
    FMT_ET_OLE,
    FMT_ET_ZIP,
    FMT_UNKNOWN,
    FMT_XLS,
    FMT_XLSB,
    FMT_XLSM,
    FMT_XLSX,
)

# OLE2 / CFB 复合文档魔数
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
# zip 魔数（普通条目 / 空归档 / 分卷结束）
ZIP_MAGICS = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")

# 走"文本 → CSV"嗅探的扩展名
_TEXT_EXTS = {".csv", ".tsv", ".txt", ""}

# 压缩包内可作为主工作簿的成员
_WORKBOOK_XML = "xl/workbook.xml"
_WORKBOOK_BIN = "xl/workbook.bin"
_CONTENT_TYPES = "[content_types].xml"


# --------------------------------------------------------------------------
# 对外入口
# --------------------------------------------------------------------------
def detect_format(path: str) -> str:
    """识别文件真实格式；无法读取文件时抛 ``ValueError``（中文原因）。"""
    path = os.fspath(path)
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError as exc:
        raise ValueError(f"无法读取文件：{path}（{exc.strerror or exc}）") from exc

    ext = os.path.splitext(path)[1].lower()

    if head.startswith(OLE_MAGIC):
        return _detect_ole(path, ext)
    if head.startswith(ZIP_MAGICS):
        return _detect_zip(path, ext)
    return _detect_text(path, ext, head)


# --------------------------------------------------------------------------
# OLE2：xls / et-ole
# --------------------------------------------------------------------------
def _detect_ole(path: str, ext: str) -> str:
    """OLE2 容器里只可能是 BIFF 系工作簿（xls 或 WPS 的 et-ole）。

    两者内部结构一致（``Workbook`` 流 + BIFF8 记录），区别只在原生格式：
    WPS 表格另存的 ``.et`` 文件即 et-ole，其余按 xls 处理。扩展名不作为
    唯一依据时不猜：非表格扩展名返回 unknown，交给上层忽略。
    """
    if ext == ".et":
        return FMT_ET_OLE
    if ext in {".xls", ".xlsx", ".xlsm", ".xlsb", ".csv"}:
        # 魔数是 OLE2 而扩展名说 xlsx/csv 等，说明文件被改名：以内容为准
        return FMT_XLS
    # 其他扩展名（.doc/.wps/.dat…）不属于本工具支持的 OLE 表格
    return FMT_UNKNOWN


# --------------------------------------------------------------------------
# ZIP：xlsx / xlsm / xlsb / et-zip
# --------------------------------------------------------------------------
def _detect_zip(path: str, ext: str) -> str:
    try:
        with zipfile.ZipFile(path) as zf:
            actual = {}
            for name in zf.namelist():
                actual.setdefault(name.lower(), name)
            names = set(actual)
            if _WORKBOOK_BIN in names:
                return FMT_XLSB
            if _WORKBOOK_XML not in names:
                # docx / pptx / 普通 zip 等
                return FMT_UNKNOWN
            if ext == ".et":
                return FMT_ET_ZIP
            if ext == ".xlsb":
                return FMT_XLSB
            # 宏判定：内容类型声明 macroEnabled，或含 VBA 工程
            content_types = ""
            if _CONTENT_TYPES in names:
                try:
                    content_types = zf.read(actual[_CONTENT_TYPES]).decode("utf-8", "replace")
                except (KeyError, OSError, zipfile.BadZipFile):
                    content_types = ""
            lowered = content_types.lower()
            if "macroenabled" in lowered or "vbaProject.bin" in {n.lower() for n in names}:
                return FMT_XLSM
            if not content_types.strip() and ext == ".xlsm":
                return FMT_XLSM
            return FMT_XLSX
    except (zipfile.BadZipFile, OSError, RuntimeError):
        # 魔数像 zip 但打不开：无法判定
        return FMT_UNKNOWN


# --------------------------------------------------------------------------
# 文本：csv（编码嗅探）
# --------------------------------------------------------------------------
def _detect_text(path: str, ext: str, head: bytes) -> str:
    if ext not in _TEXT_EXTS:
        return FMT_UNKNOWN
    try:
        with open(path, "rb") as f:
            raw = f.read(256 * 1024)
    except OSError:
        return FMT_UNKNOWN
    decoded = decode_text_bytes(raw)
    if decoded is None:
        return FMT_UNKNOWN
    text = decoded[1]
    if not _looks_like_text(text):
        return FMT_UNKNOWN
    if ext in (".csv", ".tsv"):
        return FMT_CSV
    # .txt / 无扩展名：要求存在明显分隔符才认定为 CSV
    return FMT_CSV if any(ch in text for ch in ",\t;|") else FMT_UNKNOWN


# --------------------------------------------------------------------------
# 文本编码嗅探（csvr 复用）
# --------------------------------------------------------------------------
_BOMS = (
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32"),
    (b"\x00\x00\xfe\xff", "utf-32"),
    (b"\xff\xfe", "utf-16"),
    (b"\xfe\xff", "utf-16"),
)
_TEXT_ENCODINGS = ("utf-8", "gbk", "gb18030")


def decode_text_bytes(data: bytes) -> tuple[str, str] | None:
    """嗅探文本编码并解码，返回 ``(编码名, 文本)``；不是文本返回 None。

    依次尝试：UTF-8 BOM → UTF-16/32 BOM → UTF-8 → GBK → GB18030。
    传入的 ``data`` 可能是文件头部片段（末尾可能切断多字节字符），
    解码失败时最多回退 3 个字节再试。
    """
    if not data:
        return None
    for bom, enc in _BOMS:
        if data.startswith(bom):
            text = _try_decode(data, enc)
            return None if text is None else (enc, text)
    for enc in _TEXT_ENCODINGS:
        text = _try_decode(data, enc)
        if text is not None:
            return enc, text
    return None


def _try_decode(data: bytes, enc: str) -> str | None:
    # 截断的头部可能切断多字节字符：退 1~3 字节再解
    for cut in (0, 1, 2, 3):
        if len(data) <= cut:
            break
        try:
            return data[: len(data) - cut].decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return None


def _looks_like_text(text: str) -> bool:
    """粗略判断解码结果像不像文本（排除二进制被误判）。"""
    if "\x00" in text:
        return False
    if not text:
        return False
    bad = 0
    for ch in text:
        if ch < " " and ch not in "\t\r\n":
            bad += 1
    return bad <= max(2, len(text) // 200)
