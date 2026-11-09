# -*- coding: utf-8 -*-
"""写回层（S8）：OOXML 最小改动写回。

约定（CONTRACT v1.0 / 第 4 节 S8）：
    from app.core.writers.ooxml_writer import write_cell, write_cells
    write_cell(path, ref, value, kind="value") -> dict
    write_cells(path, edits) -> int
"""
from .ooxml_writer import (  # noqa: F401
    ConflictError,
    UnsupportedFormatError,
    WriteError,
    detect_format,
    ensure_supported,
    read_cell,
    read_cells,
    scan_cells,
    write_cell,
    write_cells,
)

__all__ = [
    "write_cell",
    "write_cells",
    "read_cell",
    "read_cells",
    "scan_cells",
    "detect_format",
    "ensure_supported",
    "WriteError",
    "ConflictError",
    "UnsupportedFormatError",
]
