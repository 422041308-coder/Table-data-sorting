# -*- coding: utf-8 -*-
"""文本归一化与单元格取值工具（S2，CONTRACT §4）。

搜索、对比、映射共用的"单元格怎么取值/怎么比较"逻辑集中在这里，别处不要各写一套：

* :func:`normalize_text` —— 比较用（全角→半角、大小写、空白、引号统一）
* :func:`cell_display`   —— 展示用（数字去尾 ``.0``、日期 ``2024.3.9`` 风格、公式取缓存值）
* :func:`guess_kind`     —— 一列值的类型猜测 ``text|number|date|phone|code|money``
* :func:`to_number` / :func:`to_date_text` —— 宽容解析（失败返回 ``None``，不抛异常）

约定：
* 日期统一输出 ``"2024.3.9"``（无前导 0，点分隔），与台账原文风格一致；
* ``to_number("6%") == 0.06``（百分号按比例换算）；``"(123)" == -123``（会计负数）；
* 纯数字字符串**不**当 Excel 日期序列号解析，避免把手机号/编号误判成日期。
"""
from __future__ import annotations

import datetime as _dt
import math
import re
import unicodedata
from typing import Any

from .model import (
    KIND_BLANK,
    KIND_BOOL,
    KIND_DATE,
    KIND_FORMULA,
    KIND_NUMBER,
)

__all__ = [
    "normalize_text",
    "cell_display",
    "format_number",
    "is_blank",
    "guess_kind",
    "to_number",
    "to_date_text",
]

# 空白（含零宽字符 / BOM）
_WS_RE = re.compile(r"[\s\u200b\ufeff\u00a0]+")
# 引号统一（NFKC 不管这些弯引号）
_QUOTE_MAP = str.maketrans({
    "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u201f": '"',
    "\u2018": "'", "\u2019": "'", "\u201a": "'", "\u201b": "'",
})
_CURRENCY_CHARS = "¥￥$€£"
_CURRENCY_UNITS = ("元人民币", "人民币", "元", "RMB", "rmb", "CNY", "cny")
# 日期文本：2024.3.9 / 2024-03-09 / 2024年3月9日 / 2024/3/9 12:00:00
_DATE_TEXT_RE = re.compile(
    r"^(\d{4})\s*[-./年]\s*(\d{1,2})\s*[-./月]\s*(\d{1,2})\s*日?(?:[ T].*)?$"
)
_DATE_COMPACT_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
# 电话：手机 / 固话 / 400·800 热线
_PHONE_CLEAN_RE = re.compile(r"[\s\-()（）—－—.]+")
_PHONE_MOBILE_RE = re.compile(r"^(?:86)?(1[3-9]\d{9})$")
_PHONE_LAND_RE = re.compile(r"^0\d{2,3}\d{7,8}$")
_PHONE_SERVICE_RE = re.compile(r"^(?:400|800)\d{7}$")
_CJK_RE = re.compile(r"[\u3400-\u9fff]")
_TWO_DECIMAL_RE = re.compile(r"\.\d{2}$")


# --------------------------------------------------------------------------
# 文本
# --------------------------------------------------------------------------
def normalize_text(s: Any) -> str:
    """归一化文本：全角→半角、去多余空白、统一引号、大小写归一。

    供搜索/对比做"宽松相等"判断用；不要拿它当展示文本。
    """
    if s is None:
        return ""
    if not isinstance(s, str):
        s = cell_display(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_QUOTE_MAP)
    s = _WS_RE.sub(" ", s).strip()
    return s.lower()


# --------------------------------------------------------------------------
# 展示
# --------------------------------------------------------------------------
def format_number(v: Any) -> str:
    """数字展示：整数去掉 ``.0``，小数保留可读最短形式。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if not math.isfinite(f):
        return str(v)
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return repr(f)


def _bool_display(value: bool) -> str:
    return "TRUE" if value else "FALSE"


def cell_display(cell: Any) -> str:
    """单元格展示文本。

    * ``Cell``：number 去 ``.0``、date 用 ``2024.3.9``、公式优先取缓存值、bool 用 TRUE/FALSE；
    * 也容忍直接传 ``str`` / ``int`` / ``float`` / ``datetime``（映射与测试会混着传）。
    """
    if cell is None:
        return ""
    if isinstance(cell, str):
        return cell
    if isinstance(cell, bool):
        return _bool_display(cell)
    if isinstance(cell, (int, float)):
        return format_number(cell)
    if isinstance(cell, (_dt.datetime, _dt.date)):
        return to_date_text(cell) or str(cell)

    kind = getattr(cell, "kind", "") or ""
    text = getattr(cell, "text", "") or ""
    num = getattr(cell, "num", None)
    raw = getattr(cell, "raw", None)

    if isinstance(raw, (_dt.datetime, _dt.date)):
        return to_date_text(raw) or text
    if kind == KIND_NUMBER:
        if num is not None:
            return format_number(num)
        parsed = to_number(text)
        if parsed is not None:
            return format_number(parsed)
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return format_number(raw)
        return text
    if kind == KIND_DATE:
        got = to_date_text(num) if num is not None else None
        if got is None:
            got = to_date_text(text) or to_date_text(raw)
        if got:
            return got
        if text:
            return text
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return format_number(raw)
        return str(raw) if raw is not None else ""
    if kind == KIND_FORMULA:
        # 契约：能取到缓存值就用缓存值，取不到才退回公式文本 "=..."
        if num is not None:
            return format_number(num)
        if text:
            return text
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return format_number(raw)
        return ""
    if kind == KIND_BOOL:
        if isinstance(raw, bool):
            return _bool_display(raw)
        t = normalize_text(text)
        if t in ("true", "1", "是", "yes", "y"):
            return "TRUE"
        if t in ("false", "0", "否", "no", "n"):
            return "FALSE"
        if text:
            return text
        if isinstance(raw, (int, float)) and not isinstance(raw, bool):
            return format_number(raw)
        return str(raw) if raw is not None else ""

    # 未知 / blank：按内容尽量展示
    if text:
        return text
    if num is not None:
        return format_number(num)
    if raw is None:
        return ""
    if isinstance(raw, bool):
        return _bool_display(raw)
    if isinstance(raw, (int, float)):
        return format_number(raw)
    return str(raw)


def is_blank(value: Any) -> bool:
    """展示文本为空即视为空（None / 空串 / 纯空白 / 空 Cell）。"""
    return not cell_display(value).strip()


# --------------------------------------------------------------------------
# 解析
# --------------------------------------------------------------------------
def to_number(v: Any) -> float | None:
    """宽容转数字：``"1,234.50"`` → 1234.5，``"￥52800"`` → 52800.0，
    ``"6%"`` → 0.06，``"(123)"`` → -123.0；解析不了返回 ``None``。
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return f if math.isfinite(f) else None
    if isinstance(v, (_dt.datetime, _dt.date)):
        return None
    if isinstance(v, str):
        s = v
    else:
        s = cell_display(v)
    s = unicodedata.normalize("NFKC", s).strip()
    if not s:
        return None
    neg = False
    if len(s) >= 3 and s[0] in "(（" and s[-1] in ")）":
        neg = True
        s = s[1:-1].strip()
    pct = s.endswith("%")
    if pct:
        s = s[:-1].strip()
    s = s.replace(",", "").replace("，", "").replace(" ", "")
    for ch in _CURRENCY_CHARS:
        s = s.replace(ch, "")
    for unit in _CURRENCY_UNITS:
        s = s.replace(unit, "")
    s = s.strip()
    if not s or s in ("-", "--", "—", "/", "无", "n/a", "N/A"):
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    if not math.isfinite(f):
        return None
    if pct:
        f /= 100.0
    return -f if neg else f


def _valid_ymd(y: int, m: int, d: int) -> bool:
    try:
        _dt.date(y, m, d)
    except ValueError:
        return False
    return True


def _ymd_text(y: int, m: int, d: int) -> str:
    return f"{y}.{m}.{d}"


def _date_from_serial(serial: float) -> str | None:
    """Excel 1900 日期系统序列号 → 文本日期。"""
    if not math.isfinite(serial):
        return None
    if serial <= 0 or serial > 2958465:  # Excel 最大 9999-12-31
        return None
    base = _dt.date(1899, 12, 30)
    try:
        d = base + _dt.timedelta(days=int(math.floor(serial)))
    except (OverflowError, ValueError):
        return None
    if not 1900 <= d.year <= 2200:
        return None
    return _ymd_text(d.year, d.month, d.day)


def _date_from_compact_int(n: int) -> str | None:
    if not 19000101 <= n <= 29991231:
        return None
    y, m, d = n // 10000, (n // 100) % 100, n % 100
    return _ymd_text(y, m, d) if _valid_ymd(y, m, d) else None


def to_date_text(v: Any) -> str | None:
    """宽容转日期文本 ``"2024.3.9"``；解析不了返回 ``None``。

    支持：``datetime/date``、Excel 序列号（数字）、``2024.3.9`` / ``2024-03-09`` /
    ``2024/3/9 12:00`` / ``2024年3月9日`` / ``20240309``，以及八位 ``yyyymmdd`` 整数。
    纯数字**字符串**不按序列号解析（防手机号/编号误判）。
    """
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, _dt.datetime):
        return _ymd_text(v.year, v.month, v.day)
    if isinstance(v, _dt.date):
        return _ymd_text(v.year, v.month, v.day)
    if isinstance(v, (int, float)):
        if float(v).is_integer():
            got = _date_from_compact_int(int(v))
            if got:
                return got
        return _date_from_serial(float(v))
    if not isinstance(v, str):
        v = cell_display(v)
    s = unicodedata.normalize("NFKC", str(v)).strip()
    if not s:
        return None
    m = _DATE_TEXT_RE.match(s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return _ymd_text(y, mo, d) if _valid_ymd(y, mo, d) else None
    m = _DATE_COMPACT_RE.match(s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return _ymd_text(y, mo, d) if _valid_ymd(y, mo, d) else None
    return None


# --------------------------------------------------------------------------
# 类型猜测
# --------------------------------------------------------------------------
def _is_phone(text: str) -> bool:
    t = _PHONE_CLEAN_RE.sub("", text)
    if not t:
        return False
    if t.startswith("+"):
        t = t[1:]
    return bool(
        _PHONE_MOBILE_RE.match(t)
        or _PHONE_LAND_RE.match(t)
        or _PHONE_SERVICE_RE.match(t)
    )


def _money_evidence(text: str) -> bool:
    if any(ch in text for ch in _CURRENCY_CHARS):
        return True
    if "元" in text or "人民币" in text:
        return True
    if "," in text:
        return True
    if _TWO_DECIMAL_RE.search(text):
        n = to_number(text)
        if n is not None and abs(n) >= 100:
            return True
    return False


def _looks_like_code(text: str) -> bool:
    s = text.strip()
    if len(s) < 4 or _CJK_RE.search(s):
        return False
    has_alpha = bool(re.search(r"[A-Za-z]", s))
    has_digit = bool(re.search(r"\d", s))
    if has_alpha and has_digit:
        return True
    if not has_alpha and s.isdigit() and len(s) >= 15:
        return True  # 身份证 / 长数字编号
    return False


def _classify_one(v: Any) -> str:
    text = cell_display(v).strip()
    if not text:
        return "text"
    if _is_phone(text):
        return "phone"
    kind = getattr(v, "kind", "") or ""
    if isinstance(v, (_dt.datetime, _dt.date)) or kind == KIND_DATE:
        return "date"
    if not isinstance(v, str) and isinstance(getattr(v, "raw", v), bool):
        return "text"
    num = to_number(v)
    if num is not None:
        return "money" if _money_evidence(text) else "number"
    if to_date_text(text) is not None:
        return "date"
    if _looks_like_code(text):
        return "code"
    return "text"


def guess_kind(values: Any) -> str:
    """猜测一列值的类型：``text|number|date|phone|code|money``。

    ``values`` 里可以是 ``Cell``、字符串、数字；空值自动跳过；取前 300 个样本。
    """
    if values is None:
        return "text"
    if not isinstance(values, (list, tuple)):
        values = list(values)
    counts = {"text": 0, "number": 0, "date": 0, "phone": 0, "code": 0, "money": 0}
    total = 0
    for v in list(values)[:300]:
        if is_blank(v):
            continue
        counts[_classify_one(v)] += 1
        total += 1
    if total == 0:
        return "text"
    if counts["phone"] / total >= 0.8:
        return "phone"
    if counts["date"] / total >= 0.8:
        return "date"
    numlike = counts["number"] + counts["money"]
    if numlike / total >= 0.8:
        return "money" if counts["money"] >= counts["number"] else "number"
    if counts["code"] / total >= 0.8:
        return "code"
    return "text"
