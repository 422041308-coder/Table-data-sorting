# -*- coding: utf-8 -*-
"""OOXML（xlsx/xlsm/et-zip）最小改动写回器（S8）。

设计原则（硬性要求，见 docs/CONTRACT.md 第 4 节 S8）：
1. **zip 条目级最小改动**：读出 xlsx 全部条目 → 只改目标工作表 XML 里目标 ``<c r="A1">``
   节点 → 其余条目按原始字节原样写入新 zip → 写临时文件成功后 ``os.replace`` 原子替换。
   绝不用“整表重建 / ElementTree 序列化整张工作表”的方式，避免样式、扩展命名空间、
   批注、图片、VBA 等被破坏。
2. 目标工作表内部也只做**字符串级定点替换**：只替换目标 ``<c>`` 节点，其余字节保持原样；
   仅在必要时补一个新的 ``<c>``/``<row>`` 或修正 ``<dimension>``。
3. 安全链路（顺序固定）：
   冲突检测（mtime/size vs expect_*）→ 备份 → 写临时文件 → 用 zipfile + XML **重新解析自检**
   （目标格=新值、工作表数不变、非目标条目字节完全一致）→ 原子替换；
   任一步失败删除临时文件并抛 ``WriteError``（原文件不动）。
4. 合并区只允许写左上角，其余坐标抛 ``ValueError("合并区，请编辑左上角单元格")``。
5. 公式格：``kind="formula"`` 改 ``<f>``；``kind="value"`` 改缓存 ``<v>``（保留 ``<f>``）；
   ``kind="clear"`` 删除节点并处理 ``<dimension>``。
6. 支持 xlsx / xlsm / et-zip；xls / et-ole（OLE 复合文档）抛 ``ApiError("UNSUPPORTED_FORMAT")``。
7. 只用标准库。

对外接口（契约冻结）：
    write_cell(path, ref, value, kind="value") -> dict
        # {"ref", "oldValue", "newValue", "result", "backupPath", ...}
    write_cells(path, edits) -> int
        # 同文件多格：一次备份、一次落盘
额外提供给 edit_service / 检索页的只读辅助（不属于冻结面，但对齐返回结构）：
    read_cell / read_cells / scan_cells
"""
from __future__ import annotations

import math
import os
import posixpath
import re
import stat as _stat
import threading
import time
import uuid
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Iterable
from xml.sax.saxutils import escape as _xml_escape

from .. import audit as _audit
from ..model import (
    KIND_BLANK,
    KIND_BOOL,
    KIND_FORMULA,
    KIND_NUMBER,
    KIND_TEXT,
    CellRef,
    EditOp,
    col_letter,
    split_a1,
)

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------
_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OD = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
M = "{" + _MAIN + "}"

ZIP_MAGIC = b"PK\x03\x04"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

SUPPORTED_FORMATS = ("xlsx", "xlsm", "et-zip")
_MAX_CELL_TEXT = 32767

_WRITE_LOCK = threading.RLock()

# 单元格 / 行 / sheetData / dimension / mergeCell 的“字符串级”正则。
# 说明：worksheet XML 里 <c> 不会嵌套，所以非贪婪匹配是安全的。
_CELL_RE = re.compile(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.DOTALL)
_ROW_RE = re.compile(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.DOTALL)
_SD_RE = re.compile(r"<sheetData\b([^>]*?)(?:/>|>(.*?)</sheetData>)", re.DOTALL)
_DIM_RE = re.compile(r"<dimension\b([^>]*?)/>")
_MERGE_RE = re.compile(r"<mergeCell\b[^>]*?\bref=\"([^\"]+)\"[^>]*/>")
_F_RE = re.compile(r"<f\b[^>]*?/>|<f\b[^>]*?>.*?</f>", re.DOTALL)
_WS_RE = re.compile(r"<worksheet\b[^>]*>", re.DOTALL)
_SHEETPR_RE = re.compile(r"<sheetPr\b[^>]*?/>|<sheetPr\b[^>]*?>.*?</sheetPr>", re.DOTALL)
_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_NUMERIC_RE = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")
_INT_RE = re.compile(r"[+-]?\d+")


# --------------------------------------------------------------------------
# 异常
# --------------------------------------------------------------------------
class UnsupportedFormatError(ValueError):
    """后备异常：app.server 不可导入时替代 ApiError。

    正常运行时 ``ensure_supported`` 会延迟导入 ``app.server.ApiError`` 并抛它。
    """

    code = "UNSUPPORTED_FORMAT"
    status = 400


class WriteError(Exception):
    """写回失败（自检失败 / 落盘失败）；``result`` 通常是 ``rolled_back``。"""

    def __init__(self, message: str, *, backup_path: str = "",
                 result: str = "rolled_back", cause: BaseException | None = None):
        super().__init__(message)
        self.message = message
        self.backup_path = backup_path
        self.result = result
        self.cause = cause


class ConflictError(ValueError):
    """写回前的冲突：磁盘 mtime/size 与调用方 expect_* 不一致。

    上层（edit_service）翻译成 ``ApiError("CONFLICT", ..., 409)``。
    """

    def __init__(self, message: str, *, path: str = "",
                 expected_mtime: float | None = None, actual_mtime: float | None = None,
                 expected_size: int | None = None, actual_size: int | None = None):
        super().__init__(message)
        self.path = path
        self.expected_mtime = expected_mtime
        self.actual_mtime = actual_mtime
        self.expected_size = expected_size
        self.actual_size = actual_size


def _api_error(code: str, message: str, status: int = 400):
    """延迟导入 app.server.ApiError，避免 core 层在导入期依赖 server。"""
    try:
        from app.server import ApiError  # noqa: WPS433 - 延迟导入

        return ApiError(code, message, status)
    except Exception:  # noqa: BLE001 - 独立使用 core 时退化为 ValueError
        err = UnsupportedFormatError(message)
        err.code = code  # type: ignore[attr-defined]
        err.status = status  # type: ignore[attr-defined]
        return err


# --------------------------------------------------------------------------
# 格式识别
# --------------------------------------------------------------------------
def detect_format(path: str) -> str:
    """识别格式；优先复用 S1 的 ``app.core.sniff.detect_format``，未就绪时用魔数兜底。"""
    try:
        from ..sniff import detect_format as _sniff  # type: ignore

        fmt = _sniff(path)
        if fmt:
            return fmt
    except Exception:  # noqa: BLE001 - S1 未就绪/实现异常都回退
        pass

    p = os.path.abspath(path)
    try:
        with open(p, "rb") as f:
            head = f.read(8)
    except OSError as exc:
        raise FileNotFoundError(f"无法读取文件：{p}（{exc}）") from exc
    if head.startswith(ZIP_MAGIC):
        try:
            with zipfile.ZipFile(p) as z:
                names = set(z.namelist())
                if "xl/workbook.xml" in names or "[Content_Types].xml" in names:
                    if "xl/vbaProject.bin" in names:
                        return "xlsm"
                    return "xlsx"
        except (zipfile.BadZipFile, OSError):
            return "unknown"
        return "unknown"
    if head.startswith(OLE_MAGIC):
        return "xls"  # et-ole 同为 OLE 复合文档，按不支持处理
    return "unknown"


def ensure_supported(path: str) -> str:
    """校验文件存在且是可写格式；不支持时抛 ``ApiError("UNSUPPORTED_FORMAT")``。"""
    p = os.path.abspath(path)
    if not os.path.isfile(p):
        raise FileNotFoundError(f"文件不存在：{p}")
    fmt = detect_format(p)
    if fmt in SUPPORTED_FORMATS:
        return fmt
    if fmt in ("xls", "et-ole"):
        raise _api_error(
            "UNSUPPORTED_FORMAT",
            "不支持旧版 .xls / .et（OLE）格式写回，请用 Excel/WPS 另存为 xlsx 后编辑",
            400,
        )
    raise _api_error(
        "UNSUPPORTED_FORMAT",
        f"不支持的文件格式（{fmt or '未知'}），请另存为 xlsx 后编辑",
        400,
    )


# --------------------------------------------------------------------------
# 基础 XML 工具
# --------------------------------------------------------------------------
def _decode_xml(data: bytes) -> str:
    """字节 → XML 文本（处理 BOM / UTF-16）。"""
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        return data.decode("utf-16")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise WriteError(f"工作表 XML 不是 UTF-8/UTF-16 编码，拒绝写入：{exc}") from exc


def _encode_xml(text: str) -> bytes:
    """文本 → 字节；若原声明是 UTF-16 则改成 UTF-8（我们统一按 UTF-8 写回）。"""
    head = text[:300]
    m = re.search(r'encoding\s*=\s*"([^"]+)"', head)
    if m and m.group(1).upper().replace("-", "") in ("UTF16", "UTF16LE", "UTF16BE"):
        text = text[: m.start(1)] + "UTF-8" + text[m.end(1):]
    return text.encode("utf-8")


def _escape_text(s: str) -> str:
    """单元格文本转义（保留换行，去掉 XML 非法控制字符）。"""
    s = _BAD_XML.sub("", str(s))
    if len(s) > _MAX_CELL_TEXT:
        s = s[:32000]
    return _xml_escape(s)


def _escape_attr(s: str) -> str:
    return _xml_escape(str(s), {'"': "&quot;"})


def _get_attr(attrs: str, name: str) -> str | None:
    if not attrs:
        return None
    m = re.search(r"\b" + re.escape(name) + r"\s*=\s*\"([^\"]*)\"", attrs)
    return m.group(1) if m else None


def _del_attr(attrs: str, name: str) -> str:
    if not attrs:
        return attrs or ""
    return re.sub(r"\s+" + re.escape(name) + r"\s*=\s*\"[^\"]*\"", "", attrs)


def _set_attr(attrs: str, name: str, value: str) -> str:
    attrs = attrs or ""
    pat = re.compile(r"(\s+" + re.escape(name) + r"\s*=\s*)\"[^\"]*\"")
    if pat.search(attrs):
        return pat.sub(lambda m: m.group(1) + '"' + _escape_attr(value) + '"', attrs, count=1)
    return attrs + " " + name + '="' + _escape_attr(value) + '"'


def _num_text(value: int | float) -> str:
    if isinstance(value, bool):  # bool 走 bool 分支，这里只是防御
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    f = float(value)
    if not math.isfinite(f):
        raise ValueError(f"非法数值：{f}")
    if f.is_integer() and abs(f) < 1e16:
        return str(int(f))
    return repr(f)


def _num_display(value: int | float) -> str:
    if isinstance(value, int):
        return str(value)
    f = float(value)
    if not math.isfinite(f):
        return str(f)
    if f.is_integer() and abs(f) < 1e16:
        return str(int(f))
    return repr(f)


def _format_value(value: Any) -> str:
    """把 Python 值转成审计/返回值里的展示字符串。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return _num_display(value)
    return str(value)


def _try_number(s: str):
    """字符串 → int/float；保留电话/编号类前导 0 为文本。"""
    if not isinstance(s, str):
        return None
    t = s.strip()
    if not t or not _NUMERIC_RE.fullmatch(t):
        return None
    core = t.lstrip("+-")
    if len(core) > 1 and core[0] == "0" and core[1] not in ".eE":
        return None
    try:
        if _INT_RE.fullmatch(t):
            return int(t)
        return float(t)
    except ValueError:
        return None


def _norm_path(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


# --------------------------------------------------------------------------
# 工作簿结构解析（只读）
# --------------------------------------------------------------------------
def _rels_name(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _resolve_part(base_part: str, target: str) -> str:
    target = (target or "").replace("\\", "/")
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


class _Book:
    """打开的 xlsx 包（只读结构 + 条目字节读取）。"""

    def __init__(self, path: str):
        self.path = os.path.abspath(path)
        try:
            self.zf = zipfile.ZipFile(self.path, "r")
        except (zipfile.BadZipFile, OSError) as exc:
            raise WriteError(f"无法以 ZIP 方式打开文件：{self.path}（{exc}）") from exc
        self.names = self.zf.namelist()
        self._name_set = set(self.names)
        self.comment = self.zf.comment
        self.workbook_part = self._find_workbook_part()
        self.rels = self._read_rels(self.workbook_part)
        self.sheets = self._load_sheets()
        self.sheet_count = len(self.sheets)
        self.sst_part = self._find_sst_part()
        self._sst: list[str] | None = None

    # ---------------- 基础 ----------------
    def read(self, name: str) -> bytes:
        return self.zf.read(name)

    def has(self, name: str) -> bool:
        return name in self._name_set

    def close(self) -> None:
        try:
            self.zf.close()
        except Exception:  # noqa: BLE001
            pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ---------------- 结构 ----------------
    def _find_workbook_part(self) -> str:
        try:
            rels = self._read_rels("")  # package rels：_rels/.rels
        except Exception:  # noqa: BLE001
            rels = {}
        for rel in rels.values():
            if rel["type"].endswith("/officeDocument") and self.has(rel["target"]):
                return rel["target"]
        if self.has("xl/workbook.xml"):
            return "xl/workbook.xml"
        raise WriteError("不是有效的 Excel 工作簿：缺少 workbook.xml")

    def _read_rels(self, part: str) -> dict[str, dict]:
        """读某 part 的 .rels；part 为空串表示包级 ``_rels/.rels``。"""
        name = "_rels/.rels" if part == "" else _rels_name(part)
        if not self.has(name):
            return {}
        out: dict[str, dict] = {}
        try:
            root = ET.fromstring(self.read(name))
        except ET.ParseError:
            return {}
        for el in root.findall(f"{{{_PKG_REL}}}Relationship"):
            rid = el.get("Id") or ""
            if not rid:
                continue
            target = el.get("Target") or ""
            if not target:
                continue
            out[rid] = {
                "type": el.get("Type") or "",
                "target": _resolve_part(part, target),
            }
        return out

    def _load_sheets(self) -> list[dict]:
        try:
            root = ET.fromstring(self.read(self.workbook_part))
        except (ET.ParseError, KeyError) as exc:
            raise WriteError(f"workbook.xml 解析失败：{exc}") from exc
        sheets: list[dict] = []
        for el in root.findall(f".//{M}sheet"):
            name = el.get("name") or ""
            rid = el.get(f"{{{_OD}}}id") or ""
            rel = self.rels.get(rid) or {}
            part = rel.get("target") or ""
            typ = rel.get("type") or ""
            if not part or not self.has(part):
                idx = len(sheets) + 1
                cand = f"xl/worksheets/sheet{idx}.xml"
                if self.has(cand):
                    part = cand
                else:
                    ws = sorted(
                        n for n in self.names
                        if n.endswith(".xml") and "/worksheets/" in n
                    )
                    part = ws[idx - 1] if idx - 1 < len(ws) else ""
            if not part:
                raise WriteError(f"无法定位工作表「{name}」的 XML 部件")
            sheets.append({"name": name, "part": part, "type": typ, "rid": rid})
        if not sheets:
            raise WriteError("工作簿中没有工作表")
        return sheets

    def _find_sst_part(self) -> str:
        for rel in self.rels.values():
            if rel["type"].endswith("/sharedStrings") and self.has(rel["target"]):
                return rel["target"]
        if self.has("xl/sharedStrings.xml"):
            return "xl/sharedStrings.xml"
        return ""

    @property
    def sst(self) -> list[str]:
        if self._sst is None:
            self._sst = []
            if self.sst_part and self.has(self.sst_part):
                try:
                    root = ET.fromstring(self.read(self.sst_part))
                except ET.ParseError:
                    root = None
                if root is not None:
                    for si in root.findall(f"{M}si"):
                        parts: list[str] = []
                        for child in si:
                            if child.tag == f"{M}t":
                                parts.append(child.text or "")
                            elif child.tag == f"{M}r":
                                for t in child.findall(f"{M}t"):
                                    parts.append(t.text or "")
                        self._sst.append("".join(parts))
        return self._sst

    def sheet_part(self, sheet_name: str) -> dict:
        for s in self.sheets:
            if s["name"] == sheet_name:
                if s["type"] and s["type"].endswith("/chartsheet"):
                    raise ValueError(f"「{sheet_name}」是图表工作表，不支持编辑")
                return s
        raise ValueError(f"工作表不存在：{sheet_name}")


# --------------------------------------------------------------------------
# 单元格解析（只读）
# --------------------------------------------------------------------------
def _merges_from_root(root) -> list[tuple[int, int, int, int]]:
    out: list[tuple[int, int, int, int]] = []
    mc = root.find(f"{M}mergeCells")
    if mc is None:
        return out
    for el in mc.findall(f"{M}mergeCell"):
        ref = el.get("ref") or ""
        if ":" not in ref:
            continue
        try:
            a, b = ref.split(":", 1)
            r1, c1 = split_a1(a)
            r2, c2 = split_a1(b)
        except ValueError:
            continue
        out.append((min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)))
    return out


def _parse_merges_text(xml: str) -> list[tuple[int, int, int, int]]:
    out: list[tuple[int, int, int, int]] = []
    for m in _MERGE_RE.finditer(xml):
        ref = m.group(1)
        if ":" not in ref:
            continue
        try:
            a, b = ref.split(":", 1)
            r1, c1 = split_a1(a)
            r2, c2 = split_a1(b)
        except ValueError:
            continue
        out.append((min(r1, r2), max(r1, r2), min(c1, c2), max(c1, c2)))
    return out


def _merge_for(merges: list[tuple[int, int, int, int]], row: int, col: int):
    for r1, r2, c1, c2 in merges:
        if r1 <= row <= r2 and c1 <= col <= c2:
            return (r1, r2, c1, c2)
    return None


def _merge_owner_info(merges, row: int, col: int) -> dict | None:
    rect = _merge_for(merges, row, col)
    if not rect:
        return None
    r1, r2, c1, c2 = rect
    return {
        "merge": f"{col_letter(c1)}{r1}:{col_letter(c2)}{r2}",
        "owner": f"{col_letter(c1)}{r1}",
        "editable": (row == r1 and col == c1),
    }


def _cell_element_info(elem, sst: list[str], a1: str = "") -> dict:
    """把一个 ``<c>`` 元素解析成统一结构。"""
    t = elem.get("t")
    f_el = elem.find(f"{M}f")
    v_el = elem.find(f"{M}v")
    is_el = elem.find(f"{M}is")

    formula: str | None = None
    if f_el is not None:
        formula = "=" + (f_el.text or "")

    v_text = v_el.text if v_el is not None else None
    kind: str = KIND_BLANK
    raw: Any = None
    text = ""
    num: float | int | None = None
    cached_text = ""

    if t == "inlineStr" and is_el is not None:
        text = "".join((x.text or "") for x in is_el.iter(f"{M}t"))
        raw = text
        kind = KIND_TEXT
    elif t == "s":
        idx = -1
        try:
            idx = int((v_text or "").strip())
        except (TypeError, ValueError):
            idx = -1
        text = sst[idx] if 0 <= idx < len(sst) else ""
        raw = text
        kind = KIND_TEXT
        if formula is not None:
            cached_text = text
    elif t == "str":
        text = v_text or ""
        raw = text
        kind = KIND_TEXT
        if formula is not None:
            cached_text = text
    elif t == "b":
        b = (v_text or "").strip() in ("1", "true", "TRUE")
        raw = b
        text = "TRUE" if b else "FALSE"
        kind = KIND_BOOL
    elif t == "e":
        text = v_text or ""
        raw = text
        kind = KIND_TEXT
    elif v_text is not None and v_text != "":
        try:
            num = float(v_text)
        except ValueError:
            text = v_text or ""
            raw = text
            kind = KIND_TEXT
        else:
            raw = num
            kind = KIND_NUMBER
            text = _num_display(num)
    else:
        kind = KIND_BLANK

    if formula is not None:
        kind = KIND_FORMULA
        # 公式格展示文本：有缓存值就展示缓存值，没有就展示公式本身
        if cached_text:
            text = cached_text
        elif kind == KIND_FORMULA and text in ("", KIND_BLANK):
            text = formula
        if text == "":
            text = formula
        value: Any = raw if raw is not None else formula
    else:
        value = raw

    return {
        "a1": a1,
        "kind": kind,
        "value": value,
        "text": text,
        "formula": formula,
        "num": num,
        "has_formula": formula is not None,
    }


def _read_cell_from_root(root, sst: list[str], a1: str):
    for c in root.iter(f"{M}c"):
        if c.get("r") == a1:
            return _cell_element_info(c, sst, a1), c
    return None, None


def _iter_sheet_cells(root, sst: list[str]):
    """按工作表顺序产出 (a1, <c> 元素)；容错缺失 r 属性。"""
    sd = root.find(f"{M}sheetData")
    if sd is None:
        return
    row_no = 0
    for row in sd.findall(f"{M}row"):
        r_attr = row.get("r")
        try:
            row_no = int(r_attr) if r_attr else row_no + 1
        except (TypeError, ValueError):
            row_no += 1
        col_no = 0
        for c in row.findall(f"{M}c"):
            ref = c.get("r") or ""
            if ref:
                try:
                    row_no, col_no = split_a1(ref)
                except ValueError:
                    col_no += 1
                    ref = f"{col_letter(col_no)}{row_no}"
            else:
                col_no += 1
                ref = f"{col_letter(col_no)}{row_no}"
            yield ref, c


# --------------------------------------------------------------------------
# 只读对外辅助
# --------------------------------------------------------------------------
def read_cell(path: str, sheet: str, row: int, col: int) -> dict:
    """读取一个单元格的完整信息（get_cell 的底层）。

    返回 ``{"sheet","a1","row","col","kind","value","text","formula","num",
    "merge","mergeOwner","editable"}``。
    """
    ensure_supported(path)
    p = os.path.abspath(path)
    row = int(row)
    col = int(col)
    if row < 1 or col < 1:
        raise ValueError(f"坐标必须 >= 1（1 基）：row={row}, col={col}")
    a1 = f"{col_letter(col)}{row}"
    book = _Book(p)
    try:
        s = book.sheet_part(sheet)
        try:
            root = ET.fromstring(book.read(s["part"]))
        except ET.ParseError as exc:
            raise WriteError(f"工作表 XML 解析失败：「{sheet}」（{exc}）") from exc
        merges = _merges_from_root(root)
        info, _elem = _read_cell_from_root(root, book.sst, a1)
        owner = _merge_owner_info(merges, row, col)
        if info is None:
            info = {"a1": a1, "kind": KIND_BLANK, "value": None, "text": "",
                    "formula": None, "num": None, "has_formula": False}
        return {
            "sheet": sheet,
            "a1": a1,
            "row": row,
            "col": col,
            "kind": info["kind"],
            "value": info["value"],
            "text": info["text"],
            "formula": info["formula"],
            "num": info["num"],
            "merge": owner["merge"] if owner else None,
            "mergeOwner": owner["owner"] if owner else None,
            "editable": owner["editable"] if owner else True,
        }
    finally:
        book.close()


def read_cells(path: str, refs: Iterable[CellRef]) -> dict[tuple[str, str], dict]:
    """批量读单元格，返回 ``{(sheet, a1): info}``；同一文件只解一次包。"""
    ensure_supported(path)
    p = os.path.abspath(path)
    ref_list = list(refs)
    book = _Book(p)
    try:
        groups: dict[str, list[CellRef]] = {}
        for ref in ref_list:
            groups.setdefault(ref.sheet, []).append(ref)
        out: dict[tuple[str, str], dict] = {}
        for sheet, sheet_refs in groups.items():
            s = book.sheet_part(sheet)
            try:
                root = ET.fromstring(book.read(s["part"]))
            except ET.ParseError as exc:
                raise WriteError(f"工作表 XML 解析失败：「{sheet}」（{exc}）") from exc
            merges = _merges_from_root(root)
            for ref in sheet_refs:
                a1 = ref.a1
                info, _elem = _read_cell_from_root(root, book.sst, a1)
                owner = _merge_owner_info(merges, ref.row, ref.col)
                if info is None:
                    info = {"a1": a1, "kind": KIND_BLANK, "value": None, "text": "",
                            "formula": None, "num": None, "has_formula": False}
                out[(sheet, a1)] = {
                    "sheet": sheet,
                    "a1": a1,
                    "row": ref.row,
                    "col": ref.col,
                    "kind": info["kind"],
                    "value": info["value"],
                    "text": info["text"],
                    "formula": info["formula"],
                    "num": info["num"],
                    "merge": owner["merge"] if owner else None,
                    "mergeOwner": owner["owner"] if owner else None,
                    "editable": owner["editable"] if owner else True,
                }
        return out
    finally:
        book.close()


def scan_cells(path: str, sheet: str | None = None) -> list[dict]:
    """扫描工作表全部有值单元格（查找替换预览用）。"""
    ensure_supported(path)
    p = os.path.abspath(path)
    book = _Book(p)
    try:
        targets = [book.sheet_part(sheet)] if sheet else [
            s for s in book.sheets
            if not (s["type"] and s["type"].endswith("/chartsheet"))
        ]
        sst = book.sst
        out: list[dict] = []
        for s in targets:
            try:
                root = ET.fromstring(book.read(s["part"]))
            except ET.ParseError as exc:
                raise WriteError(f"工作表 XML 解析失败：「{s['name']}」（{exc}）") from exc
            merges = _merges_from_root(root)
            for a1, elem in _iter_sheet_cells(root, sst):
                info = _cell_element_info(elem, sst, a1)
                if info["kind"] == KIND_BLANK:
                    continue
                try:
                    r, c = split_a1(a1)
                except ValueError:
                    continue
                owner = _merge_owner_info(merges, r, c)
                info.update({
                    "sheet": s["name"],
                    "row": r,
                    "col": c,
                    "merge": owner["merge"] if owner else None,
                    "mergeOwner": owner["owner"] if owner else None,
                    "editable": owner["editable"] if owner else True,
                })
                out.append(info)
        return out
    finally:
        book.close()


# --------------------------------------------------------------------------
# 写回计划（spec）
# --------------------------------------------------------------------------
class _Spec:
    __slots__ = ("ref", "value", "kind", "expect_mtime", "expect_size")

    def __init__(self, ref: CellRef, value: Any = None, kind: str = "value",
                 expect_mtime: float = 0.0, expect_size: int = 0):
        self.ref = ref
        self.value = value
        self.kind = kind
        self.expect_mtime = expect_mtime
        self.expect_size = expect_size


def _ref_from_any(ref, default_path: str | None = None) -> CellRef:
    """把 dict / CellRef 统一成 CellRef。"""
    if isinstance(ref, CellRef):
        if not ref.file_path and default_path:
            ref = CellRef(default_path, ref.sheet, ref.row, ref.col)
        return ref
    if isinstance(ref, dict):
        file_path = (
            ref.get("filePath") or ref.get("file_path") or ref.get("path")
            or default_path or ""
        )
        sheet = ref.get("sheet") or ref.get("sheetName") or ""
        if not sheet:
            raise ValueError("写回坐标缺少 sheet")
        row = ref.get("row")
        col = ref.get("col")
        a1 = ref.get("a1")
        if (row in (None, "") or col in (None, "")) and a1:
            row, col = split_a1(a1)
        try:
            row_i, col_i = int(row), int(col)
        except (TypeError, ValueError):
            raise ValueError(f"写回坐标缺少合法 row/col：{ref!r}")
        return CellRef(os.path.abspath(file_path) if file_path else "", str(sheet), row_i, col_i)
    if isinstance(ref, (tuple, list)) and len(ref) >= 4:
        return CellRef(str(ref[0]), str(ref[1]), int(ref[2]), int(ref[3]))
    raise ValueError(f"无法识别的写回坐标：{ref!r}")


def _coerce_kind(kind: str) -> str:
    k = str(kind or "value").strip().lower()
    if k in ("value", "formula", "clear"):
        return k
    if k in ("text", "string", "number", "num", "bool"):
        # 兼容前端/其他切片把“写入类型”当“值类型”传；统一按 value 处理，
        # 具体按 Python 值类型决定 XML（数值→<v>、文本→inlineStr、布尔→t="b"）。
        return "value"
    raise ValueError(f"未知写回类型：{kind!r}（只支持 value/formula/clear）")


def _norm_edit(edit, path: str) -> _Spec:
    p = os.path.abspath(path)
    if isinstance(edit, EditOp):
        ref = _ref_from_any(edit.ref, p)
        return _Spec(ref, edit.new_value, _coerce_kind(edit.kind),
                     float(edit.expect_mtime or 0), int(edit.expect_size or 0))
    if isinstance(edit, dict):
        ref = _ref_from_any(edit.get("ref"), p)
        if "value" in edit:
            value = edit.get("value")
        else:
            value = edit.get("newValue")
        kind = _coerce_kind(edit.get("kind", "value"))
        return _Spec(ref, value, kind,
                     float(edit.get("expectMtime") or edit.get("expect_mtime") or 0),
                     int(edit.get("expectSize") or edit.get("expect_size") or 0))
    if isinstance(edit, (tuple, list)) and len(edit) >= 2:
        ref = _ref_from_any(edit[0], p)
        kind = _coerce_kind(edit[2] if len(edit) > 2 else "value")
        return _Spec(ref, edit[1], kind)
    raise ValueError(f"无法识别的写回项：{edit!r}")


def _validate_spec_path(spec: _Spec, path: str) -> None:
    if not spec.ref.file_path:
        spec.ref.file_path = path
        return
    if _norm_path(spec.ref.file_path) != _norm_path(path):
        raise ValueError(
            f"一次写回只能操作同一文件：{spec.ref.file_path} != {path}"
        )


def _check_conflicts(path: str, specs: list[_Spec]) -> None:
    """冲突检测：磁盘 mtime/size 与调用方 expect_* 不一致则拒绝（expect 为 0 视为未传）。"""
    try:
        st = os.stat(path)
    except OSError as exc:
        raise WriteError(f"无法读取文件状态：{path}（{exc}）") from exc
    for spec in specs:
        if spec.expect_mtime:
            if abs(st.st_mtime - float(spec.expect_mtime)) > 1e-6:
                raise ConflictError(
                    "文件已被修改（mtime 与打开时不一致），请刷新后重试",
                    path=path,
                    expected_mtime=float(spec.expect_mtime),
                    actual_mtime=st.st_mtime,
                )
        if spec.expect_size:
            if int(st.st_size) != int(spec.expect_size):
                raise ConflictError(
                    "文件已被修改（大小与打开时不一致），请刷新后重试",
                    path=path,
                    expected_size=int(spec.expect_size),
                    actual_size=int(st.st_size),
                )


# --------------------------------------------------------------------------
# 目标工作表 XML 的字符串级改写
# --------------------------------------------------------------------------
def _find_cell_match(xml: str, a1: str):
    for m in _CELL_RE.finditer(xml):
        if _get_attr(m.group(1), "r") == a1:
            return m
    return None


def _replace_span(xml: str, match, new: str) -> str:
    return xml[: match.start()] + new + xml[match.end():]


def _insert_into_row(inner: str, cell_xml: str, col_no: int) -> str:
    for m in _CELL_RE.finditer(inner):
        ref = _get_attr(m.group(1), "r")
        if not ref:
            continue
        try:
            c = split_a1(ref)[1]
        except ValueError:
            continue
        if c > col_no:
            return inner[: m.start()] + cell_xml + inner[m.start():]
    return inner + cell_xml


def _insert_cell(xml: str, a1: str, cell_xml: str) -> str:
    """把新 ``<c>`` 插到正确行、正确列序位置。"""
    row_no, col_no = split_a1(a1)
    for m in _ROW_RE.finditer(xml):
        r_attr = _get_attr(m.group(1), "r")
        try:
            rnum = int(r_attr) if r_attr is not None else -1
        except ValueError:
            continue
        if rnum == row_no:
            inner = m.group(2) if m.group(2) is not None else ""
            new_inner = _insert_into_row(inner, cell_xml, col_no)
            new_row = "<row" + m.group(1) + ">" + new_inner + "</row>"
            return _replace_span(xml, m, new_row)
        if rnum > row_no:
            new_row = f'<row r="{row_no}">{cell_xml}</row>'
            return xml[: m.start()] + new_row + xml[m.start():]
    sd = _SD_RE.search(xml)
    if sd is None:
        raise WriteError("工作表结构异常：找不到 <sheetData>")
    if sd.group(2) is None:
        new_sd = "<sheetData>" + f'<row r="{row_no}">{cell_xml}</row>' + "</sheetData>"
        return _replace_span(xml, sd, new_sd)
    close = "</sheetData>"
    pos = xml.rfind(close, sd.start(), sd.end())
    new_row = f'<row r="{row_no}">{cell_xml}</row>'
    return xml[:pos] + new_row + xml[pos:]


def _parse_dimension(xml: str):
    m = _DIM_RE.search(xml)
    if not m:
        return None
    ref = _get_attr(m.group(1), "ref") or ""
    if ":" not in ref:
        return None
    try:
        a, b = ref.split(":", 1)
        r1, c1 = split_a1(a)
        r2, c2 = split_a1(b)
    except ValueError:
        return None
    return (min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))


def _compute_dimension(xml: str):
    min_r = min_c = 10 ** 9
    max_r = max_c = 0
    for m in _CELL_RE.finditer(xml):
        ref = _get_attr(m.group(1), "r")
        if not ref:
            continue
        try:
            r, c = split_a1(ref)
        except ValueError:
            continue
        min_r = min(min_r, r)
        min_c = min(min_c, c)
        max_r = max(max_r, r)
        max_c = max(max_c, c)
    if max_r == 0:
        return None
    return (min_r, min_c, max_r, max_c)


def _dim_ref(rect: tuple[int, int, int, int]) -> str:
    r1, c1, r2, c2 = rect
    start = f"{col_letter(c1)}{r1}"
    end = f"{col_letter(c2)}{r2}"
    return start if start == end else f"{start}:{end}"


def _fix_dimension(xml: str, specs: list[_Spec]) -> str:
    """clear 后重算 dimension；写入超出原区域时扩展 dimension。"""
    has_clear = any(
        s.kind == "clear" or (s.kind == "value" and s.value is None) for s in specs
    )
    rect = _parse_dimension(xml)
    if rect is None or has_clear:
        rect = _compute_dimension(xml)
    if rect is None:
        rect = (1, 1, 1, 1)
    for spec in specs:
        if spec.kind == "clear" or (spec.kind == "value" and spec.value is None):
            continue
        try:
            row, col = split_a1(spec.ref.a1)
        except ValueError:
            continue
        r1, c1, r2, c2 = rect
        rect = (min(r1, row), min(c1, col), max(r2, row), max(c2, col))
    new_ref = _dim_ref(rect)
    m = _DIM_RE.search(xml)
    if m is not None:
        old_ref = _get_attr(m.group(1), "ref") or ""
        if old_ref == new_ref:
            return xml
        return _replace_span(xml, m, f'<dimension ref="{new_ref}"/>')
    # 原文件没有 dimension：写一个（按 OOXML 顺序放在 sheetPr 之后、sheetViews 之前）
    ws = _WS_RE.search(xml)
    if ws is None:
        return xml
    pos = ws.end()
    sp = _SHEETPR_RE.search(xml)
    if sp is not None and sp.end() <= pos:
        pos = sp.end()
    return xml[:pos] + f'<dimension ref="{new_ref}"/>' + xml[pos:]


def _formula_body(value: Any) -> str:
    body = "" if value is None else str(value)
    body = _BAD_XML.sub("", body).strip()
    if body.startswith("="):
        body = body[1:]
    if not body:
        raise ValueError("公式不能为空")
    if len(body) > _MAX_CELL_TEXT:
        body = body[:_MAX_CELL_TEXT]
    return body


def _value_mode(value: Any, existing_t: str | None, for_cache: bool) -> str:
    """决定写 XML 的方式：text / number / bool / cached_text。"""
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            return "text"
        return "number"
    if isinstance(value, str):
        if for_cache:
            return "cached_text"
        if existing_t == "b" and value.strip().lower() in ("true", "false", "1", "0"):
            return "bool"
        if existing_t in (None, "n"):
            # 原本是数值/空单元格：数字样字符串写成数值（保留 00123 这类为文本）
            if _try_number(value) is not None:
                return "number"
        return "text"
    return "cached_text" if for_cache else "text"


def _apply_edit(xml: str, spec: _Spec):
    """应用一条编辑，返回 (新 XML, 自检期望)。"""
    a1 = spec.ref.a1
    match = _find_cell_match(xml, a1)
    ex_attrs = match.group(1) if match is not None else ""
    ex_inner = match.group(2) if match is not None else None
    ex_has_formula = bool(ex_inner and _F_RE.search(ex_inner))

    # 1) clear：删节点
    if spec.kind == "clear" or (spec.kind == "value" and spec.value is None):
        if match is not None:
            xml = _replace_span(xml, match, "")
        return xml, {"kind": "clear", "a1": a1}

    # 2) formula：改 <f>
    if spec.kind == "formula":
        body = _formula_body(spec.value)
        base_attrs = _del_attr(ex_attrs, "t") if match is not None else ""
        if match is not None:
            new_cell = "<c" + base_attrs + "><f>" + _escape_text(body) + "</f></c>"
            xml = _replace_span(xml, match, new_cell)
        else:
            cell = f'<c r="{a1}"><f>' + _escape_text(body) + "</f></c>"
            xml = _insert_cell(xml, a1, cell)
        return xml, {"kind": "formula", "a1": a1, "formula": "=" + body}

    # 3) value
    value = spec.value
    if ex_has_formula:
        # 公式格：只改缓存值 <v>，保留 <f>
        f_xml = _F_RE.search(ex_inner).group(0)
        mode = _value_mode(value, _get_attr(ex_attrs, "t"), for_cache=True)
        base_attrs = _del_attr(ex_attrs, "t")
        if mode == "bool":
            t_attr = ' t="b"'
            v_xml = f"<v>{1 if value else 0}</v>"
            expected: Any = bool(value)
        elif mode == "number":
            t_attr = ""
            v_xml = f"<v>{_num_text(value)}</v>"
            expected = value
        else:
            t_attr = ' t="str"'
            text = _escape_text(value if isinstance(value, str) else str(value))
            v_xml = f"<v>{text}</v>"
            expected = str(value)
        new_cell = "<c" + base_attrs + t_attr + ">" + f_xml + v_xml + "</c>"
        xml = _replace_span(xml, match, new_cell)
        return xml, {"kind": "cached", "a1": a1, "value": expected}

    mode = _value_mode(value, _get_attr(ex_attrs, "t") if match is not None else None,
                       for_cache=False)
    base_attrs = _del_attr(ex_attrs, "t") if match is not None else ""
    if mode == "text":
        text = _escape_text(value if isinstance(value, str) else str(value))
        attrs = base_attrs + ' t="inlineStr"'
        cell = "<c" + attrs + '><is><t xml:space="preserve">' + text + "</t></is></c>"
        expected = value if isinstance(value, str) else str(value)
    elif mode == "number":
        num_text = _num_text(value)
        cell = "<c" + base_attrs + "><v>" + num_text + "</v></c>"
        expected = value
    else:  # bool
        attrs = base_attrs + ' t="b"'
        cell = "<c" + attrs + "><v>" + ("1" if value else "0") + "</v></c>"
        expected = bool(value)

    if match is not None:
        xml = _replace_span(xml, match, cell)
    else:
        # 新单元格：不带样式（若未来需要继承列样式，可在这里补 s 属性）
        if mode == "text":
            text = _escape_text(value if isinstance(value, str) else str(value))
            new_cell = f'<c r="{a1}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
        elif mode == "number":
            new_cell = f'<c r="{a1}"><v>{_num_text(value)}</v></c>'
        else:
            new_cell = f'<c r="{a1}" t="b"><v>{1 if value else 0}</v></c>'
        xml = _insert_cell(xml, a1, new_cell)
    return xml, {"kind": "value", "a1": a1, "value": expected}


# --------------------------------------------------------------------------
# 临时 zip / 自检 / 原子替换
# --------------------------------------------------------------------------
def _write_temp_zip(book: _Book, new_parts: dict[str, bytes]) -> str:
    tmp = book.path + f".tmp-{uuid.uuid4().hex[:8]}"
    try:
        with zipfile.ZipFile(tmp, "w") as out:
            out.comment = book.comment
            for name in book.names:
                zi = book.zf.getinfo(name)
                data = new_parts[name] if name in new_parts else book.zf.read(name)
                nzi = zipfile.ZipInfo(name, date_time=zi.date_time)
                if zi.is_dir():
                    nzi.compress_type = zipfile.ZIP_STORED
                    data = b""
                else:
                    ctype = zi.compress_type
                    if ctype not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED,
                                     zipfile.ZIP_BZIP2, zipfile.ZIP_LZMA):
                        ctype = zipfile.ZIP_DEFLATED
                    nzi.compress_type = ctype
                nzi.external_attr = zi.external_attr
                nzi.internal_attr = zi.internal_attr
                nzi.create_system = zi.create_system
                nzi.flag_bits = 0
                out.writestr(nzi, data)
    except BaseException:
        _safe_remove(tmp)
        raise
    return tmp


def _safe_remove(path: str) -> None:
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def _values_equal(actual: Any, expected: Any) -> bool:
    if isinstance(expected, bool) or isinstance(actual, bool):
        return bool(actual) == bool(expected)
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if isinstance(expected, float) and not math.isfinite(expected):
            return False
        return float(actual) == float(expected)
    if expected is None:
        return actual is None or actual == ""
    return str(actual) == str(expected)


def _verify_expected(root, a1: str, exp: dict) -> None:
    elem = None
    for c in root.iter(f"{M}c"):
        if c.get("r") == a1:
            elem = c
            break
    if exp["kind"] == "clear":
        if elem is not None:
            raise WriteError(f"自检失败：{a1} 应被清空但仍存在")
        return
    if elem is None:
        raise WriteError(f"自检失败：{a1} 不存在")
    if exp["kind"] == "formula":
        f_el = elem.find(f"{M}f")
        if f_el is None or ("=" + (f_el.text or "")) != exp["formula"]:
            raise WriteError(
                f"自检失败：{a1} 公式与预期不符（{exp['formula']!r}）"
            )
        return
    info = _cell_element_info(elem, [])
    if exp["kind"] == "cached":
        if info["formula"] is None:
            raise WriteError(f"自检失败：{a1} 公式 <f> 丢失")
    if not _values_equal(info["value"], exp["value"]):
        raise WriteError(
            f"自检失败：{a1} 实际={info['value']!r} 预期={exp['value']!r}"
        )


def _self_check(tmp: str, original_path: str, book: _Book,
                expected: dict[tuple[str, str], dict], new_parts: dict[str, bytes]) -> None:
    """重新打开临时 zip：zip 完整性、条目清单、工作表数、目标格、非目标条目字节。"""
    with zipfile.ZipFile(tmp) as z:
        bad = z.testzip()
        if bad is not None:
            raise WriteError(f"zip 完整性校验失败：{bad}")
        if z.namelist() != book.names:
            raise WriteError("写回后 ZIP 条目清单发生变化")
        try:
            wb_root = ET.fromstring(z.read(book.workbook_part))
        except (ET.ParseError, KeyError) as exc:
            raise WriteError(f"workbook.xml 重新解析失败：{exc}") from exc
        sheet_count = len(wb_root.findall(f".//{M}sheet"))
        if sheet_count != book.sheet_count:
            raise WriteError(
                f"自检失败：工作表数量变化 {book.sheet_count} → {sheet_count}"
            )
        part_roots: dict[str, object] = {}
        for (part, a1), exp in expected.items():
            if part not in part_roots:
                try:
                    part_roots[part] = ET.fromstring(z.read(part))
                except ET.ParseError as exc:
                    raise WriteError(f"目标工作表 XML 重新解析失败：{exc}") from exc
            _verify_expected(part_roots[part], a1, exp)
        # 非目标条目必须字节完全一致（样式、其他 sheet、workbook.xml、VBA、图片…）
        with zipfile.ZipFile(original_path) as src:
            for name in src.namelist():
                if name in new_parts:
                    continue
                if z.read(name) != src.read(name):
                    raise WriteError(f"自检失败：非目标条目字节被改动（{name}）")


def _atomic_replace(tmp: str, target: str) -> None:
    """临时文件 → 目标文件原子替换；只读文件临时放开权限，替换后恢复。"""
    mode = None
    if os.path.exists(target):
        mode = os.stat(target).st_mode
        if not (mode & _stat.S_IWRITE):
            os.chmod(target, mode | _stat.S_IWRITE)
    try:
        for attempt in range(3):
            try:
                os.replace(tmp, target)
                break
            except PermissionError:
                if attempt == 2:
                    raise
                time.sleep(0.2)
    finally:
        if mode is not None and not (mode & _stat.S_IWRITE) and os.path.exists(target):
            try:
                os.chmod(target, mode)
            except OSError:
                pass


# --------------------------------------------------------------------------
# 写回主流程
# --------------------------------------------------------------------------
def _write_specs(path: str, specs: list[_Spec], *,
                 backup_path: str | None = None,
                 root: str | None = None) -> dict:
    path = os.path.abspath(path)
    if not specs:
        return {"count": 0, "backupPath": backup_path or ""}
    ensure_supported(path)
    for spec in specs:
        if not str(spec.ref.sheet or "").strip():
            raise ValueError("写回坐标缺少工作表名 sheet")
        if int(spec.ref.row) < 1 or int(spec.ref.col) < 1:
            raise ValueError(
                f"坐标是 1 基，row/col 必须 >= 1：{spec.ref.sheet}!r{spec.ref.row}c{spec.ref.col}"
            )
        _validate_spec_path(spec, path)

    # 冲突检测（锁外先快速失败一次）
    _check_conflicts(path, specs)

    with _WRITE_LOCK:
        # 锁内复核，缩小 TOCTOU 窗口
        _check_conflicts(path, specs)
        book = _Book(path)
        try:
            groups: dict[str, list[_Spec]] = {}
            for spec in specs:
                groups.setdefault(spec.ref.sheet, []).append(spec)

            new_parts: dict[str, bytes] = {}
            expected: dict[tuple[str, str], dict] = {}
            for sheet, sheet_specs in groups.items():
                s = book.sheet_part(sheet)  # 不存在/图表表 → ValueError
                part = s["part"]
                text = _decode_xml(book.read(part))
                merges = _parse_merges_text(text)
                # 合并区校验：只允许左上角
                for spec in sheet_specs:
                    rect = _merge_for(merges, spec.ref.row, spec.ref.col)
                    if rect is not None and not (
                        spec.ref.row == rect[0] and spec.ref.col == rect[2]
                    ):
                        raise ValueError("合并区，请编辑左上角单元格")
                for spec in sheet_specs:
                    text, exp = _apply_edit(text, spec)
                    exp["sheet"] = sheet
                    expected[(part, spec.ref.a1)] = exp
                text = _fix_dimension(text, sheet_specs)
                new_parts[part] = _encode_xml(text)

            # 备份（调用方已备份时直接复用其路径；要求顺序：冲突→备份→写）
            created_backup = False
            if backup_path is None:
                backup_path = _audit.backup_file(path, root or _audit.PROJECT_ROOT)
                created_backup = True

            tmp = _write_temp_zip(book, new_parts)
            # Windows 下目标文件仍被本进程 zip 句柄打开时 os.replace 会报 WinError 5，
            # 因此落盘前必须关闭源 zip（book 的属性值仍可安全复用）。
            book.close()
            try:
                _self_check(tmp, path, book, expected, new_parts)
            except BaseException as exc:
                _safe_remove(tmp)
                msg = f"写回自检失败，已回滚（原文件未改动）：{exc}"
                raise WriteError(msg, backup_path=backup_path or "",
                                 result="rolled_back", cause=exc) from exc
            try:
                _atomic_replace(tmp, path)
            except BaseException as exc:
                _safe_remove(tmp)
                raise WriteError(
                    f"写回落盘失败，已回滚（原文件未改动）：{exc}",
                    backup_path=backup_path or "", result="rolled_back", cause=exc,
                ) from exc

            st = os.stat(path)
            return {
                "count": len(specs),
                "backupPath": backup_path,
                "createdBackup": created_backup,
                "current": {"mtime": st.st_mtime, "size": st.st_size},
                "expected": expected,
            }
        finally:
            book.close()


def write_cell(path: str, ref: CellRef, value: Any, kind: str = "value") -> dict:
    """写一个单元格（单文件单格：一次备份、一次落盘）。

    返回 ``{"ref", "oldValue", "newValue", "result": "ok", "backupPath", "sheet", "a1"}``。
    合并区非左上角 → ``ValueError("合并区，请编辑左上角单元格")``。
    """
    p = os.path.abspath(path)
    r = _ref_from_any(ref, p)
    old_info = read_cell(p, r.sheet, r.row, r.col)
    spec = _Spec(r, value, _coerce_kind(kind))
    res = _write_specs(p, [spec])
    new_text = "=" + _formula_body(value) if spec.kind == "formula" else (
        "" if spec.kind == "clear" or value is None else _format_value(value)
    )
    return {
        "ref": r.to_dict(),
        "oldValue": old_info["text"],
        "newValue": new_text,
        "result": "ok",
        "backupPath": res.get("backupPath", ""),
        "current": res.get("current"),
        "sheet": r.sheet,
        "a1": r.a1,
    }


def write_cells(path: str, edits: list, *, backup_path: str | None = None,
                root: str | None = None) -> int:
    """同文件多格写回：一次备份、一次落盘，返回写入条数。"""
    p = os.path.abspath(path)
    specs = [_norm_edit(e, p) for e in edits]
    res = _write_specs(p, specs, backup_path=backup_path, root=root)
    return int(res.get("count", 0))


__all__ = [
    "SUPPORTED_FORMATS",
    "WriteError",
    "ConflictError",
    "UnsupportedFormatError",
    "detect_format",
    "ensure_supported",
    "read_cell",
    "read_cells",
    "scan_cells",
    "write_cell",
    "write_cells",
]
