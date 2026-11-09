# -*- coding: utf-8 -*-
"""CSV 读取器：自动嗅探编码（utf-8-sig / utf-8 / gbk / gb18030 / utf-16）与分隔符。

约定：
* 首行就是数据行（第一行不特殊处理，表头由 S2 结构识别层判断）；
* 空字段不生成 Cell，值一律按文本（``kind="text"``）保留原始内容；
* 工作表名取文件名（不含扩展名）。
"""
from __future__ import annotations

import csv
import io
import os

from ..model import KIND_TEXT, Cell, SheetModel, WorkbookModel
from ..sniff import decode_text_bytes

_MAX_FIELD_SIZE = 32 * 1024 * 1024  # 长文本单元格（备注/JSON 串）不超过 32MB


def read_csv(path: str) -> WorkbookModel:
    """读取 CSV/TSV；失败抛 ValueError（中文原因）。"""
    path = os.fspath(path)
    file_name = os.path.basename(path)
    try:
        with open(path, "rb") as fp:
            raw = fp.read()
    except OSError as exc:
        raise ValueError(f"无法读取文件：{path}（{exc.strerror or exc}）") from exc
    if not raw:
        raise ValueError(f"CSV 文件为空：{file_name}")

    decoded = decode_text_bytes(raw)
    if decoded is None:
        raise ValueError(
            f"无法识别 CSV 编码（已尝试 UTF-8 / GBK / GB18030 / UTF-16）：{file_name}")
    encoding, text = decoded
    if "\x00" in text:
        raise ValueError(f"文件内容不是文本，无法按 CSV 读取：{file_name}")

    delimiter = _detect_delimiter(text)
    try:
        csv.field_size_limit(_MAX_FIELD_SIZE)
    except (OverflowError, ValueError):
        pass

    rows: list[list[str]] = []
    try:
        reader = csv.reader(io.StringIO(text, newline=""),
                            delimiter=delimiter, quotechar='"')
        for row in reader:
            rows.append(row)
    except csv.Error as exc:
        raise ValueError(f"CSV 内容格式错误（{file_name}）：{exc}") from exc

    while rows and not any(value.strip() for value in rows[-1]):
        rows.pop()
    if not rows:
        raise ValueError(f"CSV 文件没有数据行：{file_name}")

    n_rows = len(rows)
    n_cols = max((len(row) for row in rows), default=0)
    cells: dict[tuple[int, int], Cell] = {}
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            if value == "":
                continue
            cells[(row_index, col_index)] = Cell(
                raw=value, kind=KIND_TEXT, text=value)

    sheet_name = os.path.splitext(file_name)[0].strip() or "CSV"
    sheet = SheetModel(name=sheet_name, n_rows=n_rows, n_cols=n_cols, cells=cells)
    wb = WorkbookModel(path=path)
    wb.sheets.append(sheet)
    wb.notes.append(f"CSV 编码：{encoding}；分隔符：{delimiter!r}")
    return wb


def _detect_delimiter(text: str) -> str:
    """优先用 csv.Sniffer，失败时按前若干行的字符出现次数投票。"""
    sample = text[:65536]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
        return dialect.delimiter
    except csv.Error:
        pass
    lines = [line for line in sample.splitlines()[:50] if line.strip()]
    best, best_count = ",", 0
    for candidate in (",", "\t", ";", "|"):
        count = sum(line.count(candidate) for line in lines)
        if count > best_count:
            best, best_count = candidate, count
    return best
