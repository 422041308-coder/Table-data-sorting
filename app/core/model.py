# -*- coding: utf-8 -*-
"""核心数据契约（CONTRACT v1.0，冻结）。

所有切片只依赖本模块，不得自行增删字段；确需变更由 Lead 升版本号并广播。

坐标约定（极其重要）：
* ``SheetModel.cells`` 的键是 **0 基** ``(row, col)``，与 Excel 内部一致；
* ``CellRef`` 是 **1 基**，``row``/``col`` 与用户在 Excel 里看到的一致；
* 两者换算只有一处：``CellRef.from_index(path, sheet, r0, c0)``。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

CONTRACT_VERSION = "1.0.0"

# 单元格类型
KIND_TEXT = "text"
KIND_NUMBER = "number"
KIND_DATE = "date"
KIND_BOOL = "bool"
KIND_FORMULA = "formula"
KIND_BLANK = "blank"

# 文件格式（按魔数识别后的结果）
FMT_XLSX = "xlsx"
FMT_XLSM = "xlsm"
FMT_XLSB = "xlsb"
FMT_XLS = "xls"
FMT_CSV = "csv"
FMT_ET_ZIP = "et-zip"
FMT_ET_OLE = "et-ole"
FMT_UNKNOWN = "unknown"


# --------------------------------------------------------------------------
# 坐标工具
# --------------------------------------------------------------------------
def col_letter(col: int) -> str:
    """1 基列号 → 列字母。col_letter(1) == 'A'，col_letter(32) == 'AF'。"""
    if col < 1:
        raise ValueError(f"列号必须 >= 1，收到 {col}")
    out = ""
    n = col
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def col_index(letter: str) -> int:
    """列字母 → 1 基列号。大小写不敏感。"""
    letter = str(letter).strip().upper()
    if not letter or not letter.isalpha():
        raise ValueError(f"非法列字母: {letter!r}")
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def split_a1(a1: str) -> tuple[int, int]:
    """'AF12' → (12, 32)。返回 1 基 (row, col)。"""
    s = str(a1).strip().upper()
    i = 0
    while i < len(s) and s[i].isalpha():
        i += 1
    if i == 0 or i == len(s):
        raise ValueError(f"非法 A1 引用: {a1!r}")
    return int(s[i:]), col_index(s[:i])


# --------------------------------------------------------------------------
# 基础模型
# --------------------------------------------------------------------------
@dataclass
class Cell:
    """一个单元格。``text`` 是展示/检索用文本（未归一化）。"""

    raw: Any = None
    kind: str = KIND_BLANK
    text: str = ""
    num: float | None = None
    formula: str | None = None

    def to_dict(self) -> dict:
        return {
            "raw": self.raw if isinstance(self.raw, (str, int, float, bool)) or self.raw is None else str(self.raw),
            "kind": self.kind,
            "text": self.text,
            "num": self.num,
            "formula": self.formula,
        }


@dataclass
class SheetModel:
    """一个工作表。cells 键为 0 基 (row, col)；merges 为 0 基 (r1, r2, c1, c2) 闭区间。"""

    name: str
    n_rows: int = 0
    n_cols: int = 0
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)
    hidden: bool = False

    def get(self, row: int, col: int) -> Cell | None:
        return self.cells.get((row, col))


@dataclass
class WorkbookModel:
    """一个工作簿（统一只读模型）。"""

    path: str
    fmt: str = FMT_UNKNOWN
    sheets: list[SheetModel] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def sheet(self, name: str) -> SheetModel | None:
        for s in self.sheets:
            if s.name == name:
                return s
        return None


@dataclass
class FileRef:
    """索引中的一个文件。"""

    id: str
    path: str
    name: str = ""
    fmt: str = FMT_UNKNOWN
    size: int = 0
    mtime: float = 0.0
    sheets_count: int = 0
    status: str = "ok"  # ok | error | skipped
    error: str = ""
    indexed_at: str = ""

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "path": self.path,
            "name": self.name,
            "fmt": self.fmt,
            "size": self.size,
            "mtime": self.mtime,
            "sheetsCount": self.sheets_count,
            "status": self.status,
            "error": self.error,
            "indexedAt": self.indexed_at,
        }


@dataclass
class Field:
    """识别出的字段（表头列）。"""

    index: int  # 0 基列号
    name: str  # 完整名（多级用 ' / ' 连接）
    path: str = ""  # 与 name 相同，兼容旧名
    level: int = 1
    kind_hint: str = KIND_TEXT  # text|number|date|phone|code|money
    col_span: tuple[int, int] = (0, 0)  # 0 基闭区间

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "name": self.name,
            "level": self.level,
            "kindHint": self.kind_hint,
            "colSpan": list(self.col_span),
        }


@dataclass
class TableModel:
    """规范化后的"逻辑表"。

    rows 中键是 0 基列号；row_numbers 把逻辑行序映射回工作表 0 基行号，
    用于把命中/差异还原成用户看到的"第几行"。
    """

    source: FileRef | None
    sheet: str
    header_rows: list[int] = field(default_factory=list)  # 0 基
    data_start: int = 0  # 0 基
    data_end: int = -1  # 0 基，含
    excluded_rows: list[tuple[int, str]] = field(default_factory=list)  # (0 基行号, 原因)
    fields: list[Field] = field(default_factory=list)
    rows: list[dict[int, Cell]] = field(default_factory=list)
    row_numbers: list[int] = field(default_factory=list)  # 与 rows 对齐，0 基
    parse_notes: list[str] = field(default_factory=list)

    @property
    def path(self) -> str:
        return self.source.path if self.source else ""

    def field_at(self, col: int) -> Field | None:
        for f in self.fields:
            if f.index == col:
                return f
        return None


# --------------------------------------------------------------------------
# 坐标契约（搜索结果 / 对比结果 / 写回请求 / 审计共用）
# --------------------------------------------------------------------------
@dataclass
class CellRef:
    """用户可见坐标：1 基行（= Excel 左侧行标）、1 基列。"""

    file_path: str
    sheet: str
    row: int  # 1 基
    col: int  # 1 基

    @classmethod
    def from_index(cls, file_path: str, sheet: str, r0: int, c0: int) -> "CellRef":
        """0 基 (row, col) → CellRef。"""
        return cls(file_path, sheet, r0 + 1, c0 + 1)

    @property
    def col_letter(self) -> str:
        return col_letter(self.col)

    @property
    def a1(self) -> str:
        return f"{self.col_letter}{self.row}"

    @property
    def display(self) -> str:
        return f"第{self.row}行 · 第{self.col_letter}列(第{self.col}列) · {self.a1}"

    def to_dict(self) -> dict:
        return {
            "filePath": self.file_path,
            "sheet": self.sheet,
            "row": self.row,
            "col": self.col,
            "colLetter": self.col_letter,
            "a1": self.a1,
            "display": self.display,
        }


@dataclass
class Hit:
    """一条检索命中。"""

    ref: CellRef
    field_name: str = ""
    value_text: str = ""  # 单元格完整展示文本
    snippet: str = ""  # 命中片段
    highlight: list[int] = field(default_factory=list)  # snippet 内的 [start, end]
    score: float = 0.0

    def to_dict(self) -> dict:
        return {
            "ref": self.ref.to_dict(),
            "fieldName": self.field_name,
            "valueText": self.value_text,
            "snippet": self.snippet,
            "highlight": self.highlight,
            "score": self.score,
        }


@dataclass
class DiffItem:
    """一条对比差异。added=仅右有；removed=仅左有；changed=两边都有但值不同。"""

    kind: str  # 'added' | 'removed' | 'changed'
    key: str
    field_name: str = ""
    left_ref: CellRef | None = None
    right_ref: CellRef | None = None
    left_value: str = ""
    right_value: str = ""

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "key": self.key,
            "fieldName": self.field_name,
            "leftRef": self.left_ref.to_dict() if self.left_ref else None,
            "rightRef": self.right_ref.to_dict() if self.right_ref else None,
            "leftValue": self.left_value,
            "rightValue": self.right_value,
        }


# --------------------------------------------------------------------------
# 编辑写回
# --------------------------------------------------------------------------
@dataclass
class EditOp:
    """一次写回请求。"""

    ref: CellRef
    new_value: Any = None
    kind: str = "value"  # 'value' | 'formula' | 'clear'
    expect_mtime: float = 0.0
    expect_size: int = 0

    def to_dict(self) -> dict:
        return {
            "ref": self.ref.to_dict(),
            "newValue": self.new_value,
            "kind": self.kind,
            "expectMtime": self.expect_mtime,
            "expectSize": self.expect_size,
        }


@dataclass
class AuditEntry:
    """审计日志条目（追加写 data/audit/*.jsonl）。"""

    ts: str
    ref: CellRef
    old_value: str = ""
    new_value: str = ""
    backup_path: str = ""
    result: str = "ok"  # ok | conflict | rolled_back

    def to_dict(self) -> dict:
        return {
            "ts": self.ts,
            "ref": self.ref.to_dict(),
            "oldValue": self.old_value,
            "newValue": self.new_value,
            "backupPath": self.backup_path,
            "result": self.result,
        }


# --------------------------------------------------------------------------
# 模板与映射
# --------------------------------------------------------------------------
@dataclass
class TemplateDef:
    """模板登记：模板可替换，字段/数据起始格从文件解析 + 用户框选。"""

    id: str
    name: str
    file_path: str
    sheet: str
    title_range: tuple[int, int, int, int] | None = None  # 1 基 (r1, r2, c1, c2)
    header_range: tuple[int, int, int, int] = (1, 1, 1, 1)  # 1 基
    data_start: tuple[int, int] = (2, 1)  # 1 基 (row, col)
    fields: list[Field] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "filePath": self.file_path,
            "sheet": self.sheet,
            "titleRange": list(self.title_range) if self.title_range else None,
            "headerRange": list(self.header_range),
            "dataStart": list(self.data_start),
            "fields": [f.to_dict() for f in self.fields],
        }


@dataclass
class Sel:
    """行/列选择器。

    kind:
      * ``fixed``  —— value 是 1 基行号/列号；
      * ``record`` —— value 是相对当前记录行的偏移（0 = 当前行）；
      * ``header`` —— value 是来源字段名（仅列选择器用）；
      * ``span``   —— 连续 value 个（仅区域映射用）。
    """

    kind: str = "fixed"
    value: Any = 1

    def to_dict(self) -> dict:
        return {"kind": self.kind, "value": self.value}


@dataclass
class MappingRule:
    """一条坐标化映射规则。行/列为 1 基。"""

    source: dict = field(default_factory=dict)  # {"sheet": str|None, "row": Sel, "col": Sel}
    target: dict = field(default_factory=dict)  # {"sheet": str|None, "row": Sel, "col": Sel}
    mode: str = "cell"  # cell | block | transpose | filldown | repeat
    transform: str = "identity"  # identity | number | date | concat | constant
    on_missing: str = "blank"  # blank | skip | error
    constant: Any = None  # transform=constant 时使用
    note: str = ""

    def to_dict(self) -> dict:
        return {
            "source": {
                "sheet": self.source.get("sheet"),
                "row": self.source.get("row", Sel()).to_dict(),
                "col": self.source.get("col", Sel()).to_dict(),
            },
            "target": {
                "sheet": self.target.get("sheet"),
                "row": self.target.get("row", Sel()).to_dict(),
                "col": self.target.get("col", Sel()).to_dict(),
            },
            "mode": self.mode,
            "transform": self.transform,
            "onMissing": self.on_missing,
            "constant": self.constant,
            "note": self.note,
        }


@dataclass
class MappingProfile:
    """映射方案，存 data/profiles/*.json，可复用。"""

    id: str
    name: str
    template_id: str
    source_signature: str = ""
    rules: list[MappingRule] = field(default_factory=list)
    notes: str = ""

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "templateId": self.template_id,
            "sourceSignature": self.source_signature,
            "rules": [r.to_dict() for r in self.rules],
            "notes": self.notes,
        }


@dataclass
class MappedCell:
    """映射预演的一格：来源坐标（可空）→ 目标坐标 → 值。"""

    target: CellRef
    source_ref: CellRef | None = None
    value: Any = None
    rule_index: int = 0

    def to_dict(self) -> dict:
        return {
            "target": self.target.to_dict(),
            "sourceRef": self.source_ref.to_dict() if self.source_ref else None,
            "value": self.value,
            "ruleIndex": self.rule_index,
        }
