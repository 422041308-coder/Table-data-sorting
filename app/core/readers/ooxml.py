# -*- coding: utf-8 -*-
"""OOXML（xlsx / xlsm / et-zip）读取器：纯 ``zipfile`` + ``xml.etree.iterparse`` 流式解析。

覆盖范围：
* 多工作表（按 workbook.xml 顺序，支持隐藏表；跳过图表工作表）；
* 合并单元格（输出 0 基闭区间）；
* 共享字符串 / 内联字符串（含富文本 run，忽略日文注音 rPh）；
* 数值、布尔、错误值、日期（按 cellXfs 的 numFmt 判定，序列号→展示文本）；
* 公式：优先取缓存值（给 raw/num/text 并保留 ``formula``）；无缓存时
  ``kind="formula"``、``text``/``formula`` 为 ``"=..."``；
* 共享公式（``<f t="shared" si="..">``）：主格保留公式文本；子格只保留缓存值，
  不还原公式文本（平移相对引用需要完整公式词法分析，宁可不给也不给错的），
  无缓存时 ``text="=<共享公式>"``、``formula=None``。

坐标：``SheetModel.cells`` 键为 0 基 ``(row, col)``；``merges`` 为 0 基闭区间。
解析失败一律抛 ``ValueError``（中文原因），不静默返回空表。
"""
from __future__ import annotations

import datetime as _dt
import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

from ..model import (
    KIND_BLANK,
    KIND_BOOL,
    KIND_DATE,
    KIND_FORMULA,
    KIND_NUMBER,
    KIND_TEXT,
    Cell,
    SheetModel,
    WorkbookModel,
    split_a1,
)

_WORKBOOK_XML = "xl/workbook.xml"
_WORKBOOK_BIN = "xl/workbook.bin"
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


# ==========================================================================
# 日期格式工具（BIFF 读取器也复用）
# ==========================================================================
# Excel 内置数字格式里属于日期/时间的编号（中文环境同样适用）
BUILTIN_DATE_FORMATS: dict[int, tuple[bool, bool]] = {}
for _i in (14, 15, 16, 17):
    BUILTIN_DATE_FORMATS[_i] = (True, False)  # 日期
for _i in (18, 19, 20, 21):
    BUILTIN_DATE_FORMATS[_i] = (False, True)  # 时间
BUILTIN_DATE_FORMATS[22] = (True, True)  # 日期+时间
for _i in range(27, 37):
    BUILTIN_DATE_FORMATS[_i] = (True, False)  # 东亚日期
for _i in (45, 46, 47):
    BUILTIN_DATE_FORMATS[_i] = (False, True)  # 时间（含 [h]:mm:ss）
for _i in range(50, 59):
    BUILTIN_DATE_FORMATS[_i] = (True, False)  # 东亚日期

_EPOCH_1900 = _dt.datetime(1899, 12, 30)
_EPOCH_1904 = _dt.datetime(1904, 1, 1)


def _strip_format_literals(code: str) -> str:
    """去掉格式代码里的字面量/颜色/条件，只留下决定类型的字符。"""
    out: list[str] = []
    i, n = 0, len(code)
    while i < n:
        ch = code[i]
        if ch == '"':  # 引号内是字面文本
            j = code.find('"', i + 1)
            i = n if j < 0 else j + 1
            continue
        if ch in "\\_*":  # 转义 / 占位宽度字符：跳过下一个字符
            i += 2
            continue
        if ch == "[":
            j = code.find("]", i + 1)
            seg = code[i + 1:j] if j >= 0 else ""
            if seg.strip().lower() in ("h", "m", "s"):  # 经过时长
                out.append(seg.strip().lower())
            i = n if j < 0 else j + 1
            continue
        if ch == ";":  # 只看正数段
            break
        out.append(ch)
        i += 1
    return "".join(out)


def classify_date_format(fmt_id: int, code: str | None) -> tuple[bool, bool] | None:
    """判断数字格式是否日期/时间，返回 ``(含日期, 含时间)``；非日期返回 None。"""
    builtin = BUILTIN_DATE_FORMATS.get(fmt_id)
    if builtin is not None:
        return builtin
    if not code:
        return None
    stripped = _strip_format_literals(code)
    if not stripped:
        return None
    low = stripped.lower()
    if not any(ch in low for ch in "ymdhs"):
        return None
    has_y = "y" in low
    has_d = "d" in low
    has_m = "m" in low
    has_h = "h" in low
    has_s = "s" in low
    has_colon = ":" in low
    elapsed = any(tok in low for tok in ("[h]", "[m]", "[s]"))
    has_time = bool(has_h or has_s or has_colon or elapsed)
    # m 只有在 h/s/冒号的上下文里才是"分钟"，否则按"月"算日期
    has_date = bool(has_y or has_d or (has_m and not has_time))
    if not has_date and not has_time:
        return None
    return (has_date, has_time)


def serial_to_datetime(serial: float, date1904: bool = False) -> _dt.datetime:
    """Excel 日期序列号 → datetime（处理 1900 闰年 bug 与 1904 体系）。"""
    value = float(serial)
    if value != value or value in (float("inf"), float("-inf")):
        raise ValueError(f"非法日期序列号：{serial!r}")
    if date1904:
        base, days = _EPOCH_1904, value
    else:
        if 0 < value < 60:  # Excel 虚构的 1900-02-29：1..59 需 +1 天
            value += 1
        base, days = _EPOCH_1900, value
    try:
        return base + _dt.timedelta(days=days)
    except (OverflowError, ValueError) as exc:
        raise ValueError(f"日期序列号超出范围：{serial!r}") from exc


def format_serial_date(serial: float, date1904: bool,
                       has_date: bool, has_time: bool) -> str | None:
    """把日期序列号格式化为展示文本；不像日期时返回 None（调用方回退成数字）。

    文本风格与 S2 约定一致：日期 ``2024.3.9``；含时间时追加 `` 12:30[:05]``；
    纯时间 ``12:30[:05]``。
    """
    if not has_date and not has_time:
        return None
    if has_date and float(serial) < 1:
        return None
    try:
        dt = serial_to_datetime(serial, date1904)
    except ValueError:
        return None
    time_text = ""
    if has_time:
        time_text = f"{dt.hour:02d}:{dt.minute:02d}"
        if dt.second or dt.microsecond:
            time_text += f":{dt.second:02d}"
    if has_date:
        date_text = f"{dt.year}.{dt.month}.{dt.day}"
        return f"{date_text} {time_text}" if has_time else date_text
    return time_text


# ==========================================================================
# 基础工具
# ==========================================================================
def _local(tag: str) -> str:
    """去掉 XML 命名空间，只留本地标签名。"""
    return tag.rsplit("}", 1)[-1]


def _find_member(zf: zipfile.ZipFile, member: str) -> str | None:
    """按大小写不敏感方式在压缩包里找成员，返回真实成员名。"""
    if not member:
        return None
    target = member.replace("\\", "/").lstrip("/").lower()
    for name in zf.namelist():
        if name.replace("\\", "/").lower() == target:
            return name
    return None


def _resolve_target(base_dir: str, target: str) -> str:
    if not target:
        return ""
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def _parse_cell_range(ref: str | None) -> tuple[int, int, int, int] | None:
    """'A1:BF1' / 'A1' → 0 基闭区间 (r1, r2, c1, c2)。"""
    if not ref:
        return None
    text = ref.strip()
    if not text:
        return None
    parts = text.split(":")
    try:
        r1, c1 = split_a1(parts[0])
        if len(parts) > 1:
            r2, c2 = split_a1(parts[1])
        else:
            r2, c2 = r1, c1
    except (ValueError, IndexError):
        return None
    return (min(r1, r2) - 1, max(r1, r2) - 1, min(c1, c2) - 1, max(c1, c2) - 1)


def _col_from_letters(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _format_number(value: float) -> str:
    """数值展示文本：整数不带 .0，浮点取最短可回读表示。"""
    if value == int(value) and abs(value) < 1e16:
        return str(int(value))
    return repr(float(value))


def _si_text(elem: ET.Element) -> str:
    """<si> / <is> 的文本：拼接直属 <t> 与 <r><t>，忽略注音 <rPh>。"""
    parts: list[str] = []
    for child in elem:
        tag = _local(child.tag)
        if tag == "t":
            parts.append(child.text or "")
        elif tag == "r":
            for sub in child:
                if _local(sub.tag) == "t":
                    parts.append(sub.text or "")
    return "".join(parts)


# ==========================================================================
# 工作簿结构
# ==========================================================================
def _load_workbook(zf: zipfile.ZipFile):
    """返回 (workbook 成员名, 工作表定义列表, date1904)。"""
    member = _find_member(zf, _WORKBOOK_XML)
    if member is None:
        if _find_member(zf, _WORKBOOK_BIN) is not None:
            raise ValueError("这是 xlsb 二进制工作簿，请另存为 xlsx 后再读取")
        raise ValueError("压缩包中找不到 xl/workbook.xml，不是有效的 Excel 文件")
    try:
        root = ET.fromstring(zf.read(member))
    except ET.ParseError as exc:
        raise ValueError(f"xl/workbook.xml 已损坏，无法解析：{exc}") from exc
    except KeyError as exc:
        raise ValueError(f"压缩包缺少工作簿文件：{exc}") from exc

    date1904 = False
    sheets: list[dict] = []
    for elem in root.iter():
        tag = _local(elem.tag)
        if tag == "workbookPr":
            value = (elem.get("date1904") or "").strip().lower()
            date1904 = value in ("1", "true")
        elif tag == "sheet":
            rid = ""
            for key, value in elem.attrib.items():
                if _local(key) == "id":
                    rid = value
                    break
            state = (elem.get("state") or "").strip().lower()
            sheets.append({
                "name": elem.get("name") or f"Sheet{len(sheets) + 1}",
                "rid": rid,
                "hidden": state in ("hidden", "veryhidden"),
            })
    return member, sheets, date1904


def _load_rels(zf: zipfile.ZipFile, wb_member: str) -> dict[str, dict]:
    rels_member = posixpath.join(
        posixpath.dirname(wb_member), "_rels", posixpath.basename(wb_member) + ".rels")
    actual = _find_member(zf, rels_member)
    rels: dict[str, dict] = {}
    if actual is None:
        return rels
    try:
        root = ET.fromstring(zf.read(actual))
    except ET.ParseError:
        return rels
    base_dir = posixpath.dirname(wb_member)
    for elem in root.iter():
        if _local(elem.tag) != "Relationship":
            continue
        rid = elem.get("Id") or ""
        if not rid:
            continue
        rels[rid] = {
            "type": elem.get("Type") or "",
            "target": _resolve_target(base_dir, elem.get("Target") or ""),
        }
    return rels


def _rel_member(zf: zipfile.ZipFile, rels: dict[str, dict],
                type_suffix: str, default: str) -> str | None:
    want = type_suffix.lower()
    for rel in rels.values():
        rel_type = rel.get("type", "").lower()
        if rel_type.rsplit("/", 1)[-1] == want:
            actual = _find_member(zf, rel.get("target", ""))
            if actual:
                return actual
    return _find_member(zf, default)


def _load_shared_strings(zf: zipfile.ZipFile, member: str | None) -> list[str]:
    if member is None:
        return []
    out: list[str] = []
    try:
        with zf.open(member) as fp:
            for _event, elem in ET.iterparse(fp, events=("end",)):
                if _local(elem.tag) == "si":
                    out.append(_si_text(elem))
                    elem.clear()
    except ET.ParseError as exc:
        raise ValueError(f"共享字符串表（{member}）已损坏：{exc}") from exc
    return out


def _load_style_dates(zf: zipfile.ZipFile,
                      member: str | None) -> list[tuple[bool, bool] | None]:
    """解析 cellXfs：返回"风格下标 → (含日期, 含时间) 或 None"。"""
    if member is None:
        return []
    try:
        root = ET.fromstring(zf.read(member))
    except ET.ParseError as exc:
        # 样式损坏不致命：日期退化成数字，记入 notes
        raise ValueError(f"样式文件（{member}）已损坏：{exc}") from exc
    except KeyError:
        return []

    codes: dict[int, str] = {}
    for elem in root.iter():
        if _local(elem.tag) != "numFmt":
            continue
        fmt_id = elem.get("numFmtId")
        code = elem.get("formatCode")
        if fmt_id and fmt_id.isdigit() and code is not None:
            codes[int(fmt_id)] = code

    style_dates: list[tuple[bool, bool] | None] = []
    for elem in root.iter():
        if _local(elem.tag) != "cellXfs":
            continue
        for xf in elem:
            if _local(xf.tag) != "xf":
                continue
            raw = xf.get("numFmtId")
            fmt_id = int(raw) if raw and raw.isdigit() else 0
            style_dates.append(classify_date_format(fmt_id, codes.get(fmt_id)))
        break
    return style_dates


# ==========================================================================
# 工作表
# ==========================================================================
def _read_sheet(zf: zipfile.ZipFile, member: str, name: str,
                sst: list[str], style_dates: list[tuple[bool, bool] | None],
                date1904: bool) -> SheetModel:
    cells: dict[tuple[int, int], Cell] = {}
    merges: list[tuple[int, int, int, int]] = []
    max_r = max_c = -1
    dim_rows = dim_cols = 0
    cur_row = -1
    last_col = -1

    try:
        with zf.open(member) as fp:
            for event, elem in ET.iterparse(fp, events=("start", "end")):
                tag = _local(elem.tag)
                if event == "start":
                    if tag == "row":
                        raw = elem.get("r")
                        if raw and raw.strip().isdigit():
                            cur_row = int(raw) - 1
                        else:
                            cur_row += 1
                        last_col = -1
                    continue

                if tag == "c":
                    row, col = _cell_position(elem, cur_row, last_col)
                    last_col = col
                    if row > max_r:
                        max_r = row
                    if col > max_c:
                        max_c = col
                    cell = _make_cell(elem, sst, style_dates, date1904)
                    if cell is not None:
                        cells[(row, col)] = cell
                    elem.clear()
                elif tag == "row":
                    elem.clear()
                elif tag == "dimension":
                    rng = _parse_cell_range(elem.get("ref"))
                    if rng is not None:
                        dim_rows = max(dim_rows, rng[1] + 1)
                        dim_cols = max(dim_cols, rng[3] + 1)
                elif tag == "mergeCell":
                    ref = elem.get("ref")
                    rng = _parse_cell_range(ref)
                    if rng is None:
                        raise ValueError(f"合并单元格引用非法：{ref!r}")
                    merges.append(rng)
                    if rng[1] > max_r:
                        max_r = rng[1]
                    if rng[3] > max_c:
                        max_c = rng[3]
    except ET.ParseError as exc:
        raise ValueError(f"工作表 {name!r}（{member}）的 XML 已损坏：{exc}") from exc

    if max_r < 0 and max_c < 0:
        n_rows, n_cols = max(1, dim_rows), max(1, dim_cols)
    else:
        n_rows, n_cols = max_r + 1, max_c + 1
    return SheetModel(name=name, n_rows=n_rows, n_cols=n_cols,
                      cells=cells, merges=merges)


def _cell_position(elem: ET.Element, cur_row: int,
                   last_col: int) -> tuple[int, int]:
    raw = elem.get("r")
    if raw:
        match = _CELL_RE.match(raw.strip())
        if match:
            return int(match.group(2)) - 1, _col_from_letters(match.group(1))
    return (cur_row if cur_row >= 0 else 0), last_col + 1


def _make_cell(elem: ET.Element, sst: list[str],
               style_dates: list[tuple[bool, bool] | None],
               date1904: bool) -> Cell | None:
    cell_type = elem.get("t") or ""
    style_raw = elem.get("s")
    style_index = int(style_raw) if style_raw and style_raw.isdigit() else -1

    v_elem = f_elem = is_elem = None
    for child in elem:
        tag = _local(child.tag)
        if tag == "v":
            v_elem = child
        elif tag == "f":
            f_elem = child
        elif tag == "is":
            is_elem = child
    value_text = v_elem.text if v_elem is not None else None
    formula = _formula_text(f_elem)

    if cell_type == "s":
        if not sst:
            raise ValueError("工作簿缺少共享字符串表（xl/sharedStrings.xml），无法解析文本单元格")
        if value_text is None or not value_text.strip():
            raise ValueError("共享字符串单元格缺少索引值")
        try:
            index = int(value_text.strip())
        except ValueError:
            raise ValueError(f"共享字符串索引非法：{value_text!r}") from None
        if not (0 <= index < len(sst)):
            raise ValueError(f"共享字符串索引越界：{index}（共 {len(sst)} 条）")
        text = sst[index]
        return Cell(raw=text, kind=KIND_TEXT, text=text, formula=formula)
    if cell_type == "inlineStr":
        text = _si_text(is_elem) if is_elem is not None else ""
        return Cell(raw=text, kind=KIND_TEXT, text=text, formula=formula)
    if cell_type == "str":
        text = value_text or ""
        return Cell(raw=text, kind=KIND_TEXT, text=text, formula=formula)
    if cell_type == "b":
        flag = str(value_text or "").strip().lower() in ("1", "true")
        return Cell(raw=flag, kind=KIND_BOOL,
                    text="TRUE" if flag else "FALSE", formula=formula)
    if cell_type == "e":
        text = (value_text or "").strip() or "#ERROR!"
        return Cell(raw=text, kind=KIND_TEXT, text=text, formula=formula)
    if cell_type == "d":
        text = (value_text or "").strip()
        return Cell(raw=text, kind=KIND_DATE, text=text, formula=formula)

    # 数值（默认类型）或公式无缓存
    if value_text is None or not value_text.strip():
        if f_elem is not None:
            return Cell(raw=None, kind=KIND_FORMULA,
                        text=formula or "=<共享公式>", formula=formula)
        return None
    try:
        number = float(value_text.strip())
    except ValueError:
        text = value_text
        return Cell(raw=text, kind=KIND_TEXT, text=text, formula=formula)

    if 0 <= style_index < len(style_dates):
        style = style_dates[style_index]
        if style is not None:
            text = format_serial_date(number, date1904, style[0], style[1])
            if text is not None:
                return Cell(raw=number, kind=KIND_DATE, text=text,
                            num=number, formula=formula)
    raw = int(number) if number == int(number) and abs(number) < 1e16 else number
    return Cell(raw=raw, kind=KIND_NUMBER, text=_format_number(number),
                num=number, formula=formula)


def _formula_text(f_elem: ET.Element | None) -> str | None:
    """取公式文本（不带 '=' 前缀则补上）。

    共享公式主格有公式文本；子格没有自己的文本，直接返回 None
    （相对引用平移不做，避免给出错误公式）。
    """
    if f_elem is None:
        return None
    text = (f_elem.text or "").strip()
    ftype = (f_elem.get("t") or "").lower()
    if ftype == "shared":
        return ("=" + text) if text else None
    if text:
        return "=" + text
    return "=<公式>"


# ==========================================================================
# 对外入口
# ==========================================================================
def read_ooxml(path: str) -> WorkbookModel:
    """读取 xlsx / xlsm / et-zip；失败抛 ValueError（中文原因）。"""
    path = os.fspath(path)
    name = os.path.basename(path)
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
        raise ValueError(f"不是有效的 xlsx/zip 文件：{name}（{exc}）") from exc

    with zf:
        wb_member, sheet_defs, date1904 = _load_workbook(zf)
        if not sheet_defs:
            raise ValueError(f"工作簿里没有任何工作表：{name}")
        rels = _load_rels(zf, wb_member)
        sst_member = _rel_member(zf, rels, "sharedstrings", "xl/sharedStrings.xml")
        style_member = _rel_member(zf, rels, "styles", "xl/styles.xml")
        sst = _load_shared_strings(zf, sst_member)
        wb = WorkbookModel(path=path)
        try:
            style_dates = _load_style_dates(zf, style_member)
        except ValueError as exc:
            # 样式只影响日期展示：损坏时降级为纯数字，但要留下说明
            style_dates = []
            wb.notes.append(str(exc))
        for index, sheet_def in enumerate(sheet_defs):
            rel = rels.get(sheet_def["rid"])
            rel_type = (rel or {}).get("type", "")
            rel_kind = rel_type.rsplit("/", 1)[-1].lower() if rel_type else ""
            if rel_kind == "chartsheet":
                wb.notes.append(f"跳过图表工作表：{sheet_def['name']}")
                continue
            member = (rel or {}).get("target") or f"xl/worksheets/sheet{index + 1}.xml"
            actual = _find_member(zf, member)
            if actual is None:
                wb.notes.append(f"工作表 {sheet_def['name']!r} 对应的部件缺失（{member}）")
                continue
            sheet = _read_sheet(zf, actual, sheet_def["name"], sst,
                                style_dates, date1904)
            sheet.hidden = sheet_def["hidden"]
            wb.sheets.append(sheet)
        if not wb.sheets:
            raise ValueError(f"工作簿里没有可读取的工作表：{name}")
    return wb
