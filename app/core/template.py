# -*- coding: utf-8 -*-
"""模板解析与"克隆注入"（CONTRACT v1.0 · S4）。

设计原则（D4）：**只克隆、不重绘**。

* 把模板 ``.xlsx`` 当 zip 打开，只重写目标工作表的 ``xl/worksheets/sheetN.xml``；
  其余 zip 条目（styles / theme / sharedStrings / drawings / comments …）按**原始字节**
  复制，因此列宽、合并、条件格式、批注、页眉页脚、打印设置全部原样保留。
* 数据区从 ``data_start`` 行开始：取模板该行的 ``<row>`` 作为"数据样板行"，
  复制 N 份（行号顺延）；逐格保留原 ``s``（样式索引），只替换值/公式。
  - 映射写入了值的格 → 写值：字符串用**内联串**（``t="inlineStr"``），
    数值/日期写 ``<v>``（样式索引不变 ⇒ numFmt 不变），布尔写 ``t="b"``；
  - 未写入值、但样板行有公式的格 → **克隆公式**并平移相对行引用（数组公式同时改 ``ref``）；
  - 其余格 → 留空但保留样式（保证边框/填充连续）。
* 数据区原有的样板行/小计行/占位行全部被生成的 N 行替换；数据区内（含跨入数据区的）
  合并单元格被移除/截断，避免引用不存在的数据行。
* 重算 ``<dimension>``、删除 ``calcChain``（连同 Content_Types/rels 引用）、
  给 ``calcPr`` 加 ``fullCalcOnLoad="1"``，WPS/Excel 打开即重算、不提示修复。
* 写盘走"临时文件 + 原子替换"；输出文件与模板路径相同会直接拒绝。

坐标约定：本模块对外坐标均为 **1 基**（与 Excel 界面一致），
``rows`` 的键是目标 1 基列号；``inject`` 的第 i 行写到 ``data_start_row + i``。

区域映射（block/transpose）的尺寸约定见 ``app/core/mapping.py`` 模块文档。
"""
from __future__ import annotations

import hashlib
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, time
from typing import Any

from .model import Field, TemplateDef, col_letter, split_a1

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS = "{%s}" % NS_MAIN

# OOXML 中 1900 日期系统的原点（含 Excel 的 1900-02-29 历史 bug，取 1899-12-30 可对齐常规日期）
_EXCEL_EPOCH = date(1899, 12, 30)

# XML 1.0 不允许的控制字符
_BAD_XML_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
# 公式中的 A1 引用：可带 $ 锚定；排除函数名（形如 LOG10( 的跳过）
_REF_RE = re.compile(r"(?<![A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])")

_ROW_RE = re.compile(r"<row\b[^>]*?(?:/>|>.*?</row>)", re.S)
_CELL_RE = re.compile(r"<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>", re.S)
_ATTR_RE = re.compile(r'([A-Za-z_][\w:.-]*)="([^"]*)"')
_F_RE = re.compile(r"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.S)
_DIMENSION_RE = re.compile(r'<dimension\b[^>]*?ref="([^"]*)"[^>]*?/>')
_MERGECELLS_RE = re.compile(r"<mergeCells\b[^>]*?(?:/>|>.*?</mergeCells>)", re.S)
_MERGECELL_RE = re.compile(r'<mergeCell\b[^>]*?ref="([^"]*)"[^>]*?/>')
_COL_RE = re.compile(r"<col\b[^>]*?/>")
_SHEETDATA_RE = re.compile(r"(<sheetData\b[^>]*>)(.*?)(</sheetData>)", re.S)


# ==========================================================================
# 公共小工具
# ==========================================================================
def _xml_escape_text(s: str) -> str:
    """转义 XML 文本节点，并剔除 XML 不允许的控制字符（否则 WPS 会报修复）。"""
    s = _BAD_XML_CHARS.sub(" ", str(s))
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _xml_escape_attr(s: str) -> str:
    return (
        _xml_escape_text(s)
        .replace('"', "&quot;")
        .replace("\n", "&#10;")
        .replace("\r", "&#13;")
        .replace("\t", "&#9;")
    )


def date_to_serial(value: Any) -> float | int:
    """date/datetime → Excel 序列值；非日期原样返回。"""
    if isinstance(value, datetime):
        d = value.date()
        frac = (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6) / 86400.0
        return (d - _EXCEL_EPOCH).days + frac
    if isinstance(value, date):
        return (value - _EXCEL_EPOCH).days
    if isinstance(value, time):
        return (value.hour * 3600 + value.minute * 60 + value.second) / 86400.0
    return value


def _join_part(base_dir: str, target: str) -> str:
    """把 rels 里的相对 Target 解析成 zip 内路径。"""
    target = (target or "").strip()
    if target.startswith("/"):
        return target.lstrip("/")
    parts = [p for p in (base_dir + "/" + target).replace("\\", "/").split("/") if p]
    norm: list[str] = []
    for p in parts:
        if p == ".":
            continue
        if p == "..":
            if norm:
                norm.pop()
            continue
        norm.append(p)
    return "/".join(norm)


def _num_text(v: Any) -> str:
    """数值 → <v> 文本。整数不带 .0；浮点保留足够精度（避免科学计数法）。"""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    f = float(v)
    if f != f or f in (float("inf"), float("-inf")):
        raise ValueError(f"无法写入非有限数值: {v!r}")
    if f.is_integer() and abs(f) < 1e15:
        return str(int(f))
    return format(f, ".15g")


# ==========================================================================
# 读取 xlsx（zip + XML，零依赖）
# ==========================================================================
class _RawCell:
    __slots__ = ("value", "formula", "style", "type")

    def __init__(self, value=None, formula=None, style=None, type_=None):
        self.value = value
        self.formula = formula
        self.style = style
        self.type = type_

    def __repr__(self):  # pragma: no cover - 调试用
        return f"_RawCell({self.value!r}, style={self.style!r})"


class _Styles:
    """styles.xml 的轻量视图：样式索引 → numFmtId / 格式代码。"""

    # 内置日期格式
    _BUILTIN_DATE = set(range(14, 23)) | {45, 46, 47}
    _BUILTIN_MONEY = {41, 42, 43, 44}

    def __init__(self, xf_numfmts: list[int], fmt_codes: dict[int, str]):
        self.xf_numfmts = xf_numfmts
        self.fmt_codes = fmt_codes

    @classmethod
    def parse(cls, xml_text: str | None) -> "_Styles":
        if not xml_text:
            return cls([], {})
        try:
            root = ET.fromstring(xml_text)
        except ET.ParseError:
            return cls([], {})
        codes: dict[int, str] = {}
        for nf in root.iter(_NS + "numFmt"):
            try:
                codes[int(nf.get("numFmtId"))] = nf.get("formatCode") or ""
            except (TypeError, ValueError):
                continue
        xf: list[int] = []
        for cell_xf in root.iter(_NS + "cellXf"):
            try:
                xf.append(int(cell_xf.get("numFmtId") or 0))
            except ValueError:
                xf.append(0)
        return cls(xf, codes)

    def numfmt_id(self, style: Any) -> int:
        try:
            idx = int(style)
        except (TypeError, ValueError):
            return 0
        if 0 <= idx < len(self.xf_numfmts):
            return self.xf_numfmts[idx]
        return 0

    def code(self, style: Any) -> str:
        nid = self.numfmt_id(style)
        return self.fmt_codes.get(nid, "")

    def is_date(self, style: Any) -> bool:
        nid = self.numfmt_id(style)
        if nid in self._BUILTIN_DATE:
            return True
        # 去掉引号里的字面量、方括号（颜色/条件）与转义字符后再看格式码
        code = re.sub(r'"[^"]*"|\[[^\]]*\]|_.|\\.', "", self.code(style)).lower()
        return bool(re.search(r"(yy|mm|dd|h:|ss|m/d|d/m)", code))

    def is_money(self, style: Any) -> bool:
        nid = self.numfmt_id(style)
        if nid in self._BUILTIN_MONEY:
            return True
        code = self.code(style)
        return any(ch in code for ch in ("￥", "$", "¥", "元")) or "accounting" in code.lower()


class _SheetGrid:
    """一个工作表的值网格（0 基），供 load_template / 服务回退使用。"""

    __slots__ = ("name", "part", "cells", "merges", "n_rows", "n_cols", "styles", "dimension")

    def __init__(self, name: str, part: str):
        self.name = name
        self.part = part
        self.cells: dict[tuple[int, int], _RawCell] = {}
        self.merges: list[tuple[int, int, int, int]] = []
        self.n_rows = 0
        self.n_cols = 0
        self.styles = _Styles([], {})
        self.dimension = ""

    def value(self, r0: int, c0: int) -> Any:
        cell = self.cells.get((r0, c0))
        return None if cell is None else cell.value

    def text(self, r0: int, c0: int) -> str:
        v = self.value(r0, c0)
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)


class _Xlsx:
    """模板 xlsx 的只读视图（zip + XML）。"""

    def __init__(self, path: str):
        self.path = os.path.abspath(str(path))
        if not os.path.isfile(self.path):
            raise ValueError(f"模板文件不存在：{path}")
        if not zipfile.is_zipfile(self.path):
            raise ValueError(f"不是 xlsx/zip 格式（可能是 .xls 或加密文件）：{path}")
        self.zf = zipfile.ZipFile(self.path)
        self.names = set(self.zf.namelist())
        self._sheets: list[dict] | None = None
        self._sst: list[str] | None = None
        self._styles: _Styles | None = None

    def close(self):
        try:
            self.zf.close()
        except Exception:  # noqa: BLE001
            pass

    def read(self, name: str) -> bytes:
        return self.zf.read(name)

    def read_text(self, name: str) -> str | None:
        if name not in self.names:
            return None
        return self.zf.read(name).decode("utf-8", "replace")

    # ---------------- 工作簿结构 ----------------
    def sheets(self) -> list[dict]:
        if self._sheets is not None:
            return self._sheets
        wb = self.read_text("xl/workbook.xml") or ""
        rels = self.read_text("xl/_rels/workbook.xml.rels") or ""
        targets: dict[str, str] = {}
        if rels:
            try:
                root = ET.fromstring(rels)
                for rel in root.iter("{%s}Relationship" % NS_PKG_REL):
                    targets[rel.get("Id") or ""] = rel.get("Target") or ""
            except ET.ParseError:
                pass
        out: list[dict] = []
        if wb:
            try:
                root = ET.fromstring(wb)
            except ET.ParseError as e:
                raise ValueError(f"工作簿 XML 解析失败：{e}") from e
            for idx, sh in enumerate(root.iter(_NS + "sheet")):
                name = sh.get("name") or f"Sheet{idx + 1}"
                rid = sh.get("{%s}id" % NS_REL) or ""
                part = _join_part("xl", targets.get(rid, "")) if targets.get(rid) else ""
                if part not in self.names:
                    guess = f"xl/worksheets/sheet{idx + 1}.xml"
                    part = guess if guess in self.names else ""
                out.append({"name": name, "part": part, "index": idx})
        if not out:  # 极端情况：workbook.xml 缺失
            for n in sorted(self.names):
                m = re.match(r"xl/worksheets/sheet(\d+)\.xml$", n)
                if m:
                    out.append({"name": f"Sheet{len(out) + 1}", "part": n, "index": int(m.group(1)) - 1})
        if not out:
            raise ValueError(f"模板中没有工作表：{self.path}")
        self._sheets = out
        return out

    def pick_sheet(self, name: str | None) -> dict:
        sheets = self.sheets()
        if name is None:
            return sheets[0]
        want = str(name).strip()
        for s in sheets:
            if s["name"] == want:
                return s
        for s in sheets:
            if s["name"].strip().lower() == want.lower():
                return s
        names = "、".join(s["name"] for s in sheets)
        raise ValueError(f"模板中找不到工作表 {name!r}（可选：{names}）")

    def shared_strings(self) -> list[str]:
        if self._sst is not None:
            return self._sst
        xml = self.read_text("xl/sharedStrings.xml")
        out: list[str] = []
        if xml:
            try:
                root = ET.fromstring(xml)
            except ET.ParseError:
                root = None
            if root is not None:
                for si in root.iter(_NS + "si"):
                    out.append(_si_text(si))
        self._sst = out
        return out

    def styles(self) -> _Styles:
        if self._styles is None:
            self._styles = _Styles.parse(self.read_text("xl/styles.xml"))
        return self._styles

    # ---------------- 工作表网格 ----------------
    def read_sheet(self, part: str, name: str) -> _SheetGrid:
        xml = self.read_text(part)
        if xml is None:
            raise ValueError(f"模板缺少工作表 XML：{part}")
        grid = _SheetGrid(name, part)
        grid.styles = self.styles()
        sst = self.shared_strings()
        try:
            root = ET.fromstring(xml)
        except ET.ParseError as e:
            raise ValueError(f"工作表 XML 解析失败（{part}）：{e}") from e
        for c in root.iter(_NS + "c"):
            ref = c.get("r")
            if not ref:
                continue
            try:
                row1, col1 = split_a1(ref)
            except ValueError:
                continue
            grid.cells[(row1 - 1, col1 - 1)] = _read_cell(c, sst)
        for mc in root.iter(_NS + "mergeCell"):
            ref = mc.get("ref") or ""
            if ":" not in ref:
                continue
            try:
                a, b = ref.split(":", 1)
                r1, c1 = split_a1(a)
                r2, c2 = split_a1(b)
            except ValueError:
                continue
            grid.merges.append((r1 - 1, r2 - 1, c1 - 1, c2 - 1))
        dim = None
        for d in root.iter(_NS + "dimension"):
            dim = d.get("ref")
            break
        grid.dimension = dim or ""
        if dim and ":" in dim:
            try:
                a, b = dim.split(":", 1)
                _, c1 = split_a1(a)
                r2, c2 = split_a1(b)
                grid.n_rows, grid.n_cols = r2, c2
            except ValueError:
                pass
        if not grid.n_rows:
            rows = [r for (r, _c) in grid.cells] or [0]
            cols = [c for (_r, c) in grid.cells] or [0]
            grid.n_rows, grid.n_cols = max(rows) + 1, max(cols) + 1
        return grid


def _si_text(si: ET.Element) -> str:
    # 跳过拼音注音（<rPh>）里的 <t>
    skip = {id(t) for rph in si.iter(_NS + "rPh") for t in rph.iter(_NS + "t")}
    return "".join(t.text or "" for t in si.iter(_NS + "t") if id(t) not in skip)


def _read_cell(c: ET.Element, sst: list[str]) -> _RawCell:
    t = c.get("t") or "n"
    style = c.get("s")
    f = c.find(_NS + "f")
    formula = "=" + (f.text or "") if f is not None else None
    v = c.find(_NS + "v")
    txt = v.text if v is not None else None
    if t == "s":
        try:
            value = sst[int(txt)] if txt not in (None, "") else None
        except (ValueError, IndexError):
            value = txt
    elif t == "inlineStr":
        is_el = c.find(_NS + "is")
        value = _si_text(is_el) if is_el is not None else (txt or "")
    elif t == "str":
        value = txt if txt is not None else None
    elif t == "b":
        value = (txt == "1") if txt is not None else None
    elif t in ("e", "d"):
        value = txt
    else:
        if txt in (None, ""):
            value = None
        else:
            try:
                fv = float(txt)
                if fv.is_integer() and abs(fv) < 1e15 and re.fullmatch(r"-?\d+", txt.strip() or ""):
                    value = int(txt)
                else:
                    value = fv
            except ValueError:
                value = txt
    return _RawCell(value=value, formula=formula, style=style, type_=t)


# ==========================================================================
# 表头 / 数据区识别
# ==========================================================================
_DATE_TEXT_RE = re.compile(
    r"^\s*(\d{4})\s*[.\-/年]\s*(\d{1,2})\s*[.\-/月]\s*(\d{1,2})\s*日?\s*$"
)


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _as_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else format(v, ".15g")
    return str(v)


def _merge_source_map(merges: list[tuple[int, int, int, int]]) -> dict[tuple[int, int], tuple[int, int]]:
    out: dict[tuple[int, int], tuple[int, int]] = {}
    for r1, r2, c1, c2 in merges:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out[(r, c)] = (r1, c1)
    return out


def _effective_value(
    grid: _SheetGrid, merge_src: dict[tuple[int, int], tuple[int, int]], r0: int, c0: int
) -> Any:
    """取 (r0,c0) 的值；空白格若属于合并区则回退到合并区左上角的值。"""
    cell = grid.cells.get((r0, c0))
    if cell is not None and cell.value not in (None, ""):
        return cell.value
    src = merge_src.get((r0, c0))
    if src is not None and src != (r0, c0):
        sc = grid.cells.get(src)
        if sc is not None and sc.value not in (None, ""):
            return sc.value
    return None


def _detect_title_rows(grid: _SheetGrid, used: list[int], first_col: int, last_col: int) -> list[int]:
    if not used:
        return []
    width = max(1, last_col - first_col + 1)
    wide_min = max(3, int(width * 0.6))
    title: list[int] = []
    for r in used[:3]:
        if r > 2:
            break
        # 同一行里有超宽合并 → 大标题
        for r1, r2, c1, c2 in grid.merges:
            if r1 == r and r2 == r and (c2 - c1 + 1) >= wide_min:
                title.append(r)
                break
    if not title:
        # 兜底：首行只有 1 个非空格，且下一行有 >=3 个非空格 → 第一行是标题
        def _row_init(r: int) -> int:
            return sum(1 for (rr, _c), cell in grid.cells.items() if rr == r and cell.value not in (None, ""))

        if used and _row_init(used[0]) == 1 and len(used) > 1 and _row_init(used[1]) >= 3:
            title = [used[0]]
    return title


def _detect_header_rows(grid: _SheetGrid, start_row: int, max_row: int, max_levels: int = 8) -> list[int]:
    """从 start_row 起，连续"无数字值"的非空行判定为表头行。"""
    row_cells: dict[int, list[Any]] = {}
    for (r, c), cell in grid.cells.items():
        if r > max_row:
            continue
        row_cells.setdefault(r, []).append(cell.value)
    out: list[int] = []
    r = start_row
    while r <= max_row and len(out) < max_levels:
        vals = [v for v in row_cells.get(r, []) if v not in (None, "")]
        if not vals:
            break
        if any(_is_number(v) for v in vals):
            break
        out.append(r)
        r += 1
    return out


_DATE_STYLE_HINT = re.compile(r"(yy|mm|dd|m/d|h:|ss)")


def _infer_kind(grid: _SheetGrid, col: int, data_start0: int, max_row0: int, header: str) -> str:
    values: list[tuple[Any, Any]] = []
    for r in range(data_start0, min(max_row0, data_start0 + 40) + 1):
        cell = grid.cells.get((r, col))
        if cell is None or cell.value in (None, ""):
            continue
        values.append((cell.value, cell.style))
        if len(values) >= 20:
            break
    h = header.replace(" ", "")
    if any(k in h for k in ("编码", "编号", "代码", "合同号")):
        return "code"
    if not values:
        return "money" if any(k in h for k in ("金额", "元", "价款", "费用")) else "text"
    nums = [v for v, _s in values if _is_number(v)]
    strs = [str(v) for v, _s in values if isinstance(v, str)]
    styles = [s for _v, s in values]
    if nums and len(nums) >= max(1, int(len(values) * 0.8)):
        if styles and all(grid.styles.is_date(s) for s in styles[:5]):
            return "date"
        if any(grid.styles.is_money(s) for s in styles[:5]) or any(
            k in h for k in ("金额", "元", "价款", "费用")
        ):
            return "money"
        return "number"
    if strs and len(strs) >= max(1, int(len(values) * 0.8)):
        if all(_DATE_TEXT_RE.match(s) for s in strs[:10]):
            return "date"
        if all(re.fullmatch(r"1\d{10}", s.strip()) for s in strs[:10]):
            return "phone"
        if any(k in h for k in ("金额", "元", "价款", "费用")):
            return "money"
    return "text"


def _parse_layout(
    grid: _SheetGrid,
    sheet_name: str,
    file_path: str,
    header_range: tuple[int, int, int, int] | None = None,
    data_start: tuple[int, int] | None = None,
    title_range: tuple[int, int, int, int] | None = None,
    tpl_id: str | None = None,
) -> TemplateDef:
    """由网格推导字段/表头区/数据起始格。``header_range`` 等为 1 基覆盖值。"""
    used_rows = sorted({r for (r, _c), cell in grid.cells.items() if cell.value not in (None, "")})
    used_cols = sorted({c for (_r, c), cell in grid.cells.items() if cell.value not in (None, "")})
    if not used_rows or not used_cols:
        # 空表：给一个可登记的默认骨架
        return TemplateDef(
            id=tpl_id or _template_id(file_path, sheet_name),
            name=os.path.splitext(os.path.basename(file_path))[0],
            file_path=file_path,
            sheet=sheet_name,
            title_range=None,
            header_range=(1, 1, 1, 1),
            data_start=(2, 1),
            fields=[],
        )
    first_col, last_col = used_cols[0], used_cols[-1]
    max_row0 = used_rows[-1]

    if header_range is not None:
        hr1, hr2, hc1, hc2 = [int(x) for x in header_range]
        header_rows = list(range(hr1 - 1, hr2))
        first_col, last_col = hc1 - 1, hc2 - 1
        title_rows = list(range(0, hr1 - 1)) if hr1 > 1 else []
    else:
        title_rows = _detect_title_rows(grid, used_rows, first_col, last_col)
        start = (title_rows[-1] + 1) if title_rows else 0
        header_rows = _detect_header_rows(grid, start, max_row0)
        if not header_rows:
            # 没有明显表头：把标题后的第一行当表头，避免 data_start 落到标题
            if start <= max_row0:
                header_rows = [start]
            else:
                header_rows = []
        title_rows = [r for r in title_rows if r < (header_rows[0] if header_rows else 0)]
    if title_range is not None:
        tr1, tr2, tc1, tc2 = [int(x) for x in title_range]
        title_rows = list(range(tr1 - 1, tr2)) if tr1 >= 1 else []
    if not header_rows:
        header_rows = [0] if used_rows else []

    data_start0 = data_start[0] - 1 if data_start else (header_rows[-1] + 1 if header_rows else 1)
    ds_col0 = data_start[1] - 1 if data_start else first_col
    ds_col0 = max(0, min(ds_col0, max(first_col, last_col)))
    merge_src = _merge_source_map(grid.merges)

    fields: list[Field] = []
    for col in range(first_col, last_col + 1):
        labels: list[str] = []
        for hr in header_rows:
            raw = _effective_value(grid, merge_src, hr, col)
            txt = re.sub(r"\s+", " ", _as_text(raw)).strip()
            if txt and (not labels or labels[-1] != txt):
                labels.append(txt)
        if not labels:
            continue
        name = " / ".join(labels)
        fields.append(
            Field(
                index=col,
                name=name,
                path=name,
                level=len(labels),
                kind_hint=_infer_kind(grid, col, data_start0, max_row0, name),
                col_span=(col, col),
            )
        )

    title_range_out: tuple[int, int, int, int] | None = None
    if title_rows:
        title_text = _as_text(_effective_value(grid, merge_src, title_rows[0], first_col))
        title_range_out = (title_rows[0] + 1, title_rows[-1] + 1, first_col + 1, last_col + 1)
    else:
        title_text = ""

    header_range_out = (
        (header_rows[0] + 1, header_rows[-1] + 1, first_col + 1, last_col + 1)
        if header_rows
        else (1, 1, first_col + 1, last_col + 1)
    )
    name = (title_text.strip() if title_text.strip() else "") or os.path.splitext(
        os.path.basename(file_path)
    )[0]
    return TemplateDef(
        id=tpl_id or _template_id(file_path, sheet_name),
        name=name,
        file_path=os.path.abspath(file_path),
        sheet=sheet_name,
        title_range=title_range_out,
        header_range=header_range_out,
        data_start=(data_start0 + 1, ds_col0 + 1),
        fields=fields,
    )


def _template_id(file_path: str, sheet_name: str) -> str:
    raw = (os.path.abspath(file_path) + "|" + sheet_name).encode("utf-8")
    return "tpl-" + hashlib.sha1(raw).hexdigest()[:10]


# ==========================================================================
# 公共 API：模板登记解析
# ==========================================================================
def parse_template(
    path: str,
    sheet: str | None = None,
    header_range: tuple[int, int, int, int] | None = None,
    data_start: tuple[int, int] | None = None,
    title_range: tuple[int, int, int, int] | None = None,
    tpl_id: str | None = None,
) -> TemplateDef:
    """解析模板（可覆盖表头区/数据起始格，供"重新框选锚点"复用）。"""
    xlsx = _Xlsx(path)
    try:
        picked = xlsx.pick_sheet(sheet)
        if not picked.get("part"):
            raise ValueError(f"模板缺少工作表 XML：{picked.get('name')}")
        grid = xlsx.read_sheet(picked["part"], picked["name"])
    finally:
        xlsx.close()
    return _parse_layout(
        grid,
        picked["name"],
        path,
        header_range=header_range,
        data_start=data_start,
        title_range=title_range,
        tpl_id=tpl_id,
    )


def load_template(path: str, sheet: str | None = None) -> TemplateDef:
    """读取任意 xlsx 模板 → TemplateDef（CONTRACT §4 S4）。"""
    return parse_template(path, sheet)


def read_grid(path: str, sheet: str | None = None) -> dict:
    """读取任意 xlsx 单表为纯值网格（服务层在 S1/S2 就绪前的回退实现）。

    返回::

        {"path","sheet","cells": {(r0,c0): value}, "merges": [...],
         "nRows","nCols","dimension"}
    """
    xlsx = _Xlsx(path)
    try:
        picked = xlsx.pick_sheet(sheet)
        grid = xlsx.read_sheet(picked["part"], picked["name"])
    finally:
        xlsx.close()
    return {
        "path": os.path.abspath(path),
        "sheet": grid.name,
        "cells": {k: v.value for k, v in grid.cells.items()},
        "merges": list(grid.merges),
        "nRows": grid.n_rows,
        "nCols": grid.n_cols,
        "dimension": grid.dimension,
    }


def list_sheets(path: str) -> list[str]:
    """列出 xlsx/zip 格式的工作表名（按工作簿顺序）。"""
    xlsx = _Xlsx(path)
    try:
        return [s["name"] for s in xlsx.sheets()]
    finally:
        xlsx.close()


# ==========================================================================
# 公式平移
# ==========================================================================
def _shift_formula_text(text: str, row_delta: int) -> str:
    """把公式中"相对行引用"整体平移 ``row_delta`` 行（$ 锚定的行不动）。"""
    if not text or row_delta == 0:
        return text

    def _fn(chunk: str) -> str:
        def _rep(m: re.Match) -> str:
            dollar_before, letters, dollar_after, rows = m.groups()
            if dollar_after == "$":
                return m.group(0)
            if m.end() < len(chunk) and chunk[m.end()] == "(":
                return m.group(0)  # 形如 LOG10( 的函数名，不处理
            try:
                new_row = int(rows) + row_delta
            except ValueError:
                return m.group(0)
            if new_row < 1:
                return m.group(0)
            return f"{dollar_before}{letters}{new_row}"

        return _REF_RE.sub(_rep, chunk)

    out: list[str] = []
    pos = 0
    for m in re.finditer(r'"(?:[^"]|"")*"', text):
        out.append(_fn(text[pos : m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(_fn(text[pos:]))
    return "".join(out)


def _shift_ref_range(ref: str, row_delta: int) -> str:
    if not ref or row_delta == 0:
        return ref
    parts = ref.split(":")
    out: list[str] = []
    for p in parts:
        try:
            r1, c1 = split_a1(p)
        except ValueError:
            out.append(p)
            continue
        out.append(f"{col_letter(c1)}{r1 + row_delta}")
    return ":".join(out)


# ==========================================================================
# 注入实现
# ==========================================================================
def _parse_attrs(open_tag: str) -> list[tuple[str, str]]:
    return _ATTR_RE.findall(open_tag)


def _attrs_dict(attrs: list[tuple[str, str]]) -> dict[str, str]:
    return {k: v for k, v in attrs}


def _attrs_str(items: list[tuple[str, str]]) -> str:
    return "".join(f' {k}="{_xml_escape_attr(v)}"' for k, v in items)


def _split_row_tag(tag: str) -> tuple[list[tuple[str, str]], str, bool]:
    """拆 <row ...> 的属性与内部 XML；返回 (attrs, inner, self_closing)。"""
    m = re.match(r"<row\b([^>]*?)(/?)>", tag, re.S)
    if not m:
        return [], "", True
    attrs = _parse_attrs(m.group(1))
    self_closing = bool(m.group(2)) or tag.rstrip().endswith("/>")
    inner = "" if self_closing else tag[m.end() : tag.rfind("</row>")]
    return attrs, inner, self_closing


def _xml_unescape_text(s: str) -> str:
    """反转义 XML 文本/属性里的实体（用于把原始 f 文本交给公式平移器处理）。"""
    return (
        str(s)
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&#10;", "\n")
        .replace("&#13;", "\r")
        .replace("&#9;", "\t")
        .replace("&amp;", "&")
    )


def _col_default_styles(sheet_xml: str) -> dict[int, str]:
    out: dict[int, str] = {}
    for m in _COL_RE.finditer(sheet_xml):
        attrs = _attrs_dict(_parse_attrs(m.group(0)))
        style = attrs.get("style")
        if not style:
            continue
        try:
            lo = int(attrs.get("min", "0"))
            hi = int(attrs.get("max", str(lo)))
        except ValueError:
            continue
        for c in range(lo, hi + 1):
            out[c] = style
    return out


def _cell_open_tag(token: str) -> str:
    idx = token.find(">")
    if token.endswith("/>"):
        return token[: idx + 1]  # '<c ...>'（去掉结尾 '/' 由调用方处理）
    return token[: idx + 1]


def _build_data_row(
    template_tag: str,
    orig_row: int,
    new_row: int,
    values: dict[int, Any],
    col_default_styles: dict[int, str],
) -> str:
    """克隆一行：保留样式/行高，写入 values（1 基列号），公式平移。"""
    row_attrs, inner, _self = _split_row_tag(template_tag)
    row_attrs = [(k, v) for k, v in row_attrs if k not in ("r", "spans")]
    cell_infos: dict[int, dict] = {}
    max_col = 0
    for m in _CELL_RE.finditer(inner):
        token = m.group(0)
        head = _cell_open_tag(token)
        attrs = _parse_attrs(head)
        ad = _attrs_dict(attrs)
        ref = ad.get("r")
        if not ref:
            continue
        try:
            _r, col = split_a1(ref)
        except ValueError:
            continue
        max_col = max(max_col, col)
        body = "" if token.endswith("/>") else token[token.find(">") + 1 : token.rfind("</c>")]
        f_m = _F_RE.search(body)
        cell_infos[col] = {
            "attrs": [(k, v) for k, v in attrs if k not in ("r", "t", "cm")],
            "style": ad.get("s"),
            "cm": ad.get("cm"),
            "f_attrs": _attrs_dict(_parse_attrs(f_m.group(1))) if f_m else {},
            "f_text": _xml_unescape_text(f_m.group(2)) if f_m and f_m.group(2) is not None else "",
        }

    target_cols = sorted(set(cell_infos.keys()) | {int(c) for c in values.keys()})
    if not target_cols:
        return f"<row r=\"{new_row}\"{_attrs_str(row_attrs)}/>"
    delta = new_row - orig_row
    cells_xml: list[str] = []
    for col in target_cols:
        info = cell_infos.get(col)
        base_attrs = list(info["attrs"]) if info else []
        style = info["style"] if info else col_default_styles.get(col)
        if style is not None and not any(k == "s" for k, _v in base_attrs):
            base_attrs.append(("s", style))
        val = values.get(col, None)
        has_value = col in values and val is not None
        if has_value:
            cell = _value_cell(new_row, col, base_attrs, val)
        elif info is not None and info.get("f_text"):
            cell = _formula_cell(new_row, col, base_attrs, info, delta)
        elif info is not None:
            cell = f'<c r="{col_letter(col)}{new_row}"{_attrs_str(base_attrs)}/>'
        else:
            # 目标列在样板行没有对应格：有值时写无样式值格，否则跳过
            if col in values and values[col] is not None:
                cell = _value_cell(new_row, col, base_attrs, values[col])
            else:
                continue
        cells_xml.append(cell)
    spans = f' spans="1:{max(max_col, max(target_cols))}"'
    return f'<row r="{new_row}"{spans}{_attrs_str(row_attrs)}>' + "".join(cells_xml) + "</row>"


def _value_cell(new_row: int, col: int, base_attrs: list[tuple[str, str]], value: Any) -> str:
    ref = f"{col_letter(col)}{new_row}"
    attrs = [(k, v) for k, v in base_attrs if k not in ("t",)]
    if value is None or (isinstance(value, str) and value == ""):
        return f'<c r="{ref}"{_attrs_str(attrs)}/>'
    if isinstance(value, bool):
        return f'<c r="{ref}"{_attrs_str(attrs)} t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{_attrs_str(attrs)}><v>{_num_text(value)}</v></c>'
    if isinstance(value, (datetime, date, time)):
        return f'<c r="{ref}"{_attrs_str(attrs)}><v>{_num_text(date_to_serial(value))}</v></c>'
    text = _xml_escape_text(str(value))
    space = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return (
        f'<c r="{ref}"{_attrs_str(attrs)} t="inlineStr"><is><t{space}>{text}</t></is></c>'
    )


def _formula_cell(
    new_row: int, col: int, base_attrs: list[tuple[str, str]], info: dict, delta: int
) -> str:
    ref = f"{col_letter(col)}{new_row}"
    f_attrs = {
        k: v for k, v in info.get("f_attrs", {}).items() if k not in ("t", "si", "ref", "ca")
    }
    f_type = info.get("f_attrs", {}).get("t")
    formula = _shift_formula_text(info.get("f_text") or "", delta)
    extra = ""
    if f_type == "array":
        f_attrs["t"] = "array"
        f_attrs["ref"] = _shift_ref_range(info["f_attrs"].get("ref") or ref, delta)
        # 数组公式的 cm 元数据索引需要保留（metadata.xml 未改动）
        cm = info.get("cm")
        if cm:
            extra = f' cm="{_xml_escape_attr(cm)}"'
    # shared 公式：克隆时降级为普通公式（去掉 t/si），否则 Excel 认为共享定义重复
    f_xml = f"<f{_attrs_str(list(f_attrs.items()))}>{_xml_escape_text(formula)}</f>"
    return f'<c r="{ref}"{_attrs_str(base_attrs)}{extra}>{f_xml}</c>'


def _rebuild_merges(sheet_xml: str, data_start_row: int) -> str:
    """数据区内（以及跨入数据区）的合并被移除/截断。"""
    m = _MERGECELLS_RE.search(sheet_xml)
    if not m:
        return sheet_xml
    block = m.group(0)
    kept: list[str] = []
    for mm in _MERGECELL_RE.finditer(block):
        ref = mm.group(1)
        if ":" not in ref:
            continue
        try:
            a, b = ref.split(":", 1)
            r1, c1 = split_a1(a)
            r2, c2 = split_a1(b)
        except ValueError:
            continue
        if r1 >= data_start_row:
            continue  # 数据区内的样板合并：丢弃
        if r2 >= data_start_row:
            r2 = data_start_row - 1  # 跨入数据区：截断
        if r2 < r1:
            continue
        kept.append(f'<mergeCell ref="{col_letter(c1)}{r1}:{col_letter(c2)}{r2}"/>')
    if not kept:
        return sheet_xml[: m.start()] + sheet_xml[m.end() :]
    new_block = f'<mergeCells count="{len(kept)}">' + "".join(kept) + "</mergeCells>"
    return sheet_xml[: m.start()] + new_block + sheet_xml[m.end() :]


def _filter_hyperlinks(sheet_xml: str, data_start_row: int) -> str:
    """丢弃数据区内的超链接节点（它们指向被替换掉的样板行）。"""
    m = re.search(r"<hyperlinks\b[^>]*?(?:/>|>.*?</hyperlinks>)", sheet_xml, re.S)
    if not m:
        return sheet_xml
    block = m.group(0)
    news: list[str] = []
    for hm in re.finditer(r"<hyperlink\b[^>]*?/>", block):
        tag = hm.group(0)
        ad = _attrs_dict(_parse_attrs(tag))
        ref = ad.get("ref") or ""
        try:
            start = ref.split(":")[0]
            r1, _c1 = split_a1(start)
        except (ValueError, IndexError):
            news.append(tag)
            continue
        if r1 >= data_start_row:
            continue
        news.append(tag)
    if not news:
        return sheet_xml[: m.start()] + sheet_xml[m.end() :]
    new_block = "<hyperlinks>" + "".join(news) + "</hyperlinks>"
    return sheet_xml[: m.start()] + new_block + sheet_xml[m.end() :]


def _update_dimension(sheet_xml: str, ref: str) -> str:
    m = _DIMENSION_RE.search(sheet_xml)
    if m:
        return sheet_xml[: m.start()] + f'<dimension ref="{ref}"/>' + sheet_xml[m.end() :]
    # 没有 dimension：按 schema 顺序插在 sheetPr 之后，否则紧跟 <worksheet ...>
    m2 = re.search(r"<sheetPr\b[^>]*?(?:/>|>.*?</sheetPr>)", sheet_xml, re.S)
    if m2:
        insert_at = m2.end()
    else:
        wm = re.search(r"<worksheet\b[^>]*>", sheet_xml)
        insert_at = wm.end() if wm else sheet_xml.find(">") + 1
    return sheet_xml[:insert_at] + f'<dimension ref="{ref}"/>' + sheet_xml[insert_at:]


def _force_recalc(workbook_xml: str) -> str:
    m = re.search(r"<calcPr\b[^>]*?/>", workbook_xml)
    if m:
        tag = m.group(0)
        if "fullCalcOnLoad" in tag:
            new_tag = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', tag)
        else:
            new_tag = tag[:-2].rstrip() + ' fullCalcOnLoad="1"/>'
        return workbook_xml[: m.start()] + new_tag + workbook_xml[m.end() :]
    new_tag = '<calcPr calcId="191029" fullCalcOnLoad="1"/>'
    m2 = re.search(r"<extLst\b", workbook_xml)
    pos = m2.start() if m2 else workbook_xml.rfind("</workbook>")
    if pos < 0:
        return workbook_xml
    return workbook_xml[:pos] + new_tag + workbook_xml[pos:]


def _strip_calcchain(content_types: str, wb_rels: str) -> tuple[str, str]:
    ct = re.sub(
        r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", content_types
    )
    rels = re.sub(
        r'<Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*/>', "", wb_rels
    )
    return ct, rels


def _write_zip(
    zf: zipfile.ZipFile,
    out_path: str,
    replace: dict[str, bytes],
    drop: set[str],
) -> None:
    out_dir = os.path.dirname(os.path.abspath(out_path)) or "."
    os.makedirs(out_dir, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".inject-", suffix=".xlsx.tmp", dir=out_dir)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for info in zf.infolist():
                name = info.filename
                if name in drop:
                    continue
                data = replace.get(name)
                if data is None:
                    with zf.open(name) as src:
                        data = src.read()
                zo.writestr(info, data)
        os.replace(tmp, out_path)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _validate_output(out_path: str, sheet_part: str, expect_rows: int, last_row: int) -> None:
    """轻量自检：zip 可读、关键 XML 可解析、期望的数据行存在。"""
    if not zipfile.is_zipfile(out_path):
        raise ValueError(f"生成文件不是有效 xlsx：{out_path}")
    with zipfile.ZipFile(out_path) as z:
        names = z.namelist()
        if sheet_part not in names:
            raise ValueError(f"生成文件缺少工作表：{sheet_part}")
        for n in names:
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    ET.fromstring(z.read(n))
                except ET.ParseError as e:
                    raise ValueError(f"生成文件 XML 损坏（{n}）：{e}") from e
        if expect_rows > 0:
            sheet_xml = z.read(sheet_part).decode("utf-8", "replace")
            if f'<row r="{last_row}"' not in sheet_xml:
                raise ValueError(f"生成文件缺少数据行（期望第 {last_row} 行）")


def inject(
    template_path: str,
    out_path: str,
    data_start: tuple[int, int],
    rows: list[dict[int, Any]],
    sheet: str | None = None,
) -> str:
    """把数据"克隆注入"模板，保样式生成新 xlsx（CONTRACT §4 S4）。

    参数
    ----
    template_path : 模板 xlsx（只读）
    out_path      : 输出 xlsx（不存在则创建；与模板相同会被拒绝）
    data_start    : 1 基 (起始行, 起始列)；第 i 条数据写到 ``起始行 + i``
    rows          : list[dict]，键为**目标 1 基列号**，值为要写入的值
                    （str/int/float/bool/date/datetime；None/"" 视为留空）
    sheet         : 目标工作表名；None=第一个工作表

    返回输出文件的绝对路径。
    """
    if not isinstance(data_start, (tuple, list)) or len(data_start) != 2:
        raise ValueError(f"data_start 必须是 (行, 列) 1 基元组，收到 {data_start!r}")
    try:
        ds_row, ds_col = int(data_start[0]), int(data_start[1])
    except (TypeError, ValueError) as e:
        raise ValueError(f"data_start 必须是整数元组：{data_start!r}") from e
    if ds_row < 1 or ds_col < 1:
        raise ValueError(f"data_start 行列必须 >= 1：{data_start!r}")
    if not isinstance(rows, (list, tuple)):
        raise ValueError("rows 必须是 list[dict]")
    norm_rows: list[dict[int, Any]] = []
    for i, rec in enumerate(rows):
        if not isinstance(rec, dict):
            raise ValueError(f"rows[{i}] 必须是 dict，收到 {type(rec).__name__}")
        norm: dict[int, Any] = {}
        for k, v in rec.items():
            try:
                col = int(k)
            except (TypeError, ValueError) as e:
                raise ValueError(f"rows[{i}] 的列键必须是整数：{k!r}") from e
            if col >= 1:
                norm[col] = v
        norm_rows.append(norm)

    src = os.path.abspath(str(template_path))
    dst = os.path.abspath(str(out_path))
    if src == dst:
        raise ValueError("输出路径不能与模板路径相同，请换一个输出文件名")
    if not os.path.isfile(src):
        raise ValueError(f"模板文件不存在：{template_path}")

    xlsx = _Xlsx(src)
    try:
        picked = xlsx.pick_sheet(sheet)
        part = picked.get("part")
        if not part:
            raise ValueError(f"模板工作表缺少 XML：{picked.get('name')}")
        sheet_xml = xlsx.read_text(part)
        if sheet_xml is None:
            raise ValueError(f"模板缺少工作表 XML：{part}")
        wb_xml = xlsx.read_text("xl/workbook.xml") or ""
        ct_xml = xlsx.read_text("[Content_Types].xml") or ""
        wb_rels = xlsx.read_text("xl/_rels/workbook.xml.rels") or ""

        m = _SHEETDATA_RE.search(sheet_xml)
        if not m:
            raise ValueError("模板工作表缺少 <sheetData>，无法注入")
        head, sheet_data_inner, tail = sheet_xml[: m.start(2)], m.group(2), sheet_xml[m.end(2) :]

        # 1) 拆分现有行，保留数据区之前的行（标题/表头）
        template_row_tag: str | None = None
        kept_tags: list[str] = []
        max_kept_col = 0
        for rm in _ROW_RE.finditer(sheet_data_inner):
            tag = rm.group(0)
            r_attr = re.search(r'\br="(\d+)"', tag)
            if not r_attr:
                continue
            row_num = int(r_attr.group(1))
            if row_num < ds_row:
                kept_tags.append(tag)
                for cm in re.finditer(r'\br="([A-Z]+)(\d+)"', tag):
                    try:
                        max_kept_col = max(max_kept_col, split_a1(f"{cm.group(1)}1")[1])
                    except ValueError:
                        pass
            elif row_num == ds_row and template_row_tag is None:
                template_row_tag = tag
        if template_row_tag is None:
            # 模板在 data_start 行没有样板行：合成一个空样板行（无样式可继承）
            template_row_tag = f'<row r="{ds_row}" spans="1:{max(1, max_kept_col)}"/>'
            orig_row = ds_row
        else:
            _m2 = re.search(r'\br="(\d+)"', template_row_tag)
            orig_row = int(_m2.group(1)) if _m2 else ds_row

        # 2) 生成数据行
        col_styles = _col_default_styles(sheet_xml)
        new_tags: list[str] = []
        for i, rec in enumerate(norm_rows):
            new_tags.append(
                _build_data_row(template_row_tag, orig_row, ds_row + i, rec, col_styles)
            )

        # 3) 拼回 sheetData；数据区之后原有行（样板/小计/占位）全部被替换
        new_sheet_xml = head + "".join(kept_tags) + "".join(new_tags) + tail

        # 4) 合并区、超链接、dimension
        new_sheet_xml = _rebuild_merges(new_sheet_xml, ds_row)
        new_sheet_xml = _filter_hyperlinks(new_sheet_xml, ds_row)
        last_row = ds_row + len(norm_rows) - 1 if norm_rows else max(1, ds_row - 1)
        # 扫描最终 XML 中实际出现的列/行（标题/表头 + 新数据行）
        max_col = 1
        scan_max_row = 0
        for cm in re.finditer(r'\br="([A-Z]+)(\d+)"', new_sheet_xml):
            try:
                _r, c = split_a1(f"{cm.group(1)}{cm.group(2)}")
            except ValueError:
                continue
            max_col = max(max_col, c)
            scan_max_row = max(scan_max_row, _r)
        new_sheet_xml = _update_dimension(
            new_sheet_xml, f"A1:{col_letter(max_col)}{max(last_row, scan_max_row, 1)}"
        )

        # 5) 工作簿：强制重算、清 calcChain 引用
        new_wb = _force_recalc(wb_xml)
        new_ct, new_rels = _strip_calcchain(ct_xml, wb_rels)
        drop = {n for n in xlsx.names if re.search(r"(^|/)calcChain\.xml$", n)}

        replace: dict[str, bytes] = {
            part: new_sheet_xml.encode("utf-8"),
            "xl/workbook.xml": new_wb.encode("utf-8"),
        }
        if wb_rels:
            replace["xl/_rels/workbook.xml.rels"] = new_rels.encode("utf-8")
        if ct_xml:
            replace["[Content_Types].xml"] = new_ct.encode("utf-8")
        _write_zip(xlsx.zf, dst, replace, drop)
    finally:
        xlsx.close()

    _validate_output(dst, part, len(norm_rows), last_row)
    return dst


# ==========================================================================
# 输出校验（服务/测试用）
# ==========================================================================
def verify_injected(
    path: str, data_start: tuple[int, int], rows: list[dict[int, Any]], sheet: str | None = None
) -> list[str]:
    """重新解析输出文件，校验数据格是否与 rows 一致；返回问题列表（空=通过）。"""
    problems: list[str] = []
    ds_row, _ds_col = int(data_start[0]), int(data_start[1])
    try:
        grid = read_grid(path, sheet)
    except Exception as e:  # noqa: BLE001
        return [f"重新解析失败：{e}"]
    for i, rec in enumerate(rows):
        r0 = ds_row - 1 + i
        for col, want in rec.items():
            c0 = int(col) - 1
            got = grid["cells"].get((r0, c0))
            if isinstance(want, float) and isinstance(got, (int, float)):
                if abs(float(got) - want) > 1e-9:
                    problems.append(f"第{ds_row + i}行 {col_letter(int(col))}: 期望 {want!r} 实际 {got!r}")
            elif isinstance(want, (datetime, date, time)):
                continue
            elif want not in (None, "") and str(got) != _as_text(want):
                problems.append(f"第{ds_row + i}行 {col_letter(int(col))}: 期望 {want!r} 实际 {got!r}")
    return problems
