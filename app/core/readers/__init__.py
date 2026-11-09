# -*- coding: utf-8 -*-
"""读取层入口：``read_workbook`` 按格式嗅探结果分发到具体解析器。

依赖方向：``readers/* → sniff → model``；本模块不依赖 services。
"""
from __future__ import annotations

import os

from ..model import (
    FMT_CSV,
    FMT_ET_OLE,
    FMT_ET_ZIP,
    FMT_XLS,
    FMT_XLSB,
    FMT_XLSM,
    FMT_XLSX,
    WorkbookModel,
)
from ..sniff import detect_format
from .biff import read_biff
from .csvr import read_csv
from .ooxml import read_ooxml

__all__ = [
    "read_workbook",
    "read_ooxml",
    "read_biff",
    "read_csv",
    "detect_format",
]


def read_workbook(path: str) -> WorkbookModel:
    """读取任意受支持格式的表格，返回统一的 ``WorkbookModel``。

    * xlsx / xlsm / et-zip → ``read_ooxml``
    * xls / et-ole          → ``read_biff``
    * csv                   → ``read_csv``
    * xlsb / unknown        → ``ValueError``（中文原因，不静默返回空表）
    """
    path = os.fspath(path)
    if not os.path.isfile(path):
        raise ValueError(f"文件不存在：{path}")

    fmt = detect_format(path)
    try:
        if fmt in (FMT_XLSX, FMT_XLSM, FMT_ET_ZIP):
            workbook = read_ooxml(path)
        elif fmt in (FMT_XLS, FMT_ET_OLE):
            workbook = read_biff(path)
        elif fmt == FMT_CSV:
            workbook = read_csv(path)
        elif fmt == FMT_XLSB:
            raise ValueError(
                f"xlsb 二进制工作簿暂不支持：{os.path.basename(path)}，"
                "请另存为 xlsx 后再读取")
        else:
            raise ValueError(f"无法识别的文件格式：{os.path.basename(path)}")
    except ValueError:
        raise
    except Exception as exc:  # noqa: BLE001  # 兜底转成带中文原因的 ValueError
        raise ValueError(
            f"读取 {os.path.basename(path)} 失败：{type(exc).__name__}: {exc}") from exc

    workbook.fmt = fmt
    if not workbook.sheets:
        raise ValueError(f"未能从文件中解析出任何工作表：{os.path.basename(path)}")
    return workbook
