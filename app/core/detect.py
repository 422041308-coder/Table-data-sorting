# -*- coding: utf-8 -*-
"""S2 结构识别：把 ``SheetModel`` 还原成逻辑表 ``TableModel``。

识别思路（**结构法**，不靠猜内容）：

1. **大标题行**：表格上方 1~2 行，整行只有 1 个锚点单元格，且跨多列合并、
   或紧随其后（1~2 行内）出现多列表头行；
2. **表头块**：标题之后第一个非空行起，用"纵向合并覆盖"递归扩展出多级表头
   （父级合并到哪一行，那些行就都是表头），最多 6 行；
3. **数据区**：表头之后（允许跳 0~2 行空行）到最后一个非空行；
4. **逐行判定**：小计/合计行、分类占位行、重复表头行、空行 → ``excluded_rows``，
   其余进 ``rows``；``row_numbers`` 与 ``rows`` 一一对应（0 基工作表行号）；
5. **字段名**：多级表头用 ``"父级 / 子级"`` 连接；合并区的跨列父级标题通过
   合并覆盖向下传播到每个子列。
"""
from __future__ import annotations

import os
import re
from typing import Any

from .model import (
    KIND_DATE,
    KIND_FORMULA,
    KIND_NUMBER,
    Cell,
    Field,
    FileRef,
    SheetModel,
    TableModel,
    WorkbookModel,
    col_letter,
)
from .normalize import (
    cell_display,
    guess_kind,
    is_blank,
    normalize_text,
    to_date_text,
    to_number,
)

__all__ = ["detect_table", "detect_workbook"]

MAX_HEADER_ROWS = 6  # 表头最多 6 行
MAX_TITLE_ROWS = 3  # 大标题最多 3 行
MAX_TITLE_SCAN = 5  # 在前 6 行内找标题
SPACER_SKIP = 2  # 表头与数据之间最多跳 2 行空行
MAX_FIELD_SAMPLE = 300  # 猜类型时每列最多取多少行

_SUBTOTAL_WORDS = ("小计", "合计", "总计", "累计", "分计", "共计", "subtotal", "total", "sum")
_SUBTOTAL_SUFFIXES = ("金额", "数量", "数", "值", "元", "（元）", "(元)", "合计", "小计", "总计")
_SUBTOTAL_STRIP_RE = re.compile(r"[\s:：()（）\[\]【】.。,，;；\-—_]+")
# 轻量数值判断（行统计用，避免对每个单元格跑完整的 to_number）
_QUICK_NUM_RE = re.compile(r"^[+-]?[\d,]+(?:\.\d+)?%?$")

# 字段名里出现这些词，数字列更可能是金额
_MONEY_HINTS = (
    "金额", "费用", "费", "价款", "款", "价", "工资", "成本", "收入", "利润", "元", "造价",
)
# 字段名末级出现这些词，即使含"金额/元"也按数量/比例（不是钱）
_COUNT_HINTS = ("数量", "个数", "件数", "次数", "人数", "天数", "户数", "笔数", "项数", "台数", "量")
# 字段名里出现这些词，数字列更可能是日期（Excel 序列号）
_DATE_HINTS = ("时间", "日期", "年月", "生效", "到期", "截止", "开工", "竣工")
# 字段名里出现这些词，即使含"金额/元"也不是钱（税率、比例）
_RATE_HINTS = ("率", "比例", "百分", "占比")
# 首行单元格长文本达到这个长度，基本可以断定是数据行而不是表头
# （表头标签通常很短；工单的标题/诉求内容动辄上百字）
_LONG_TEXT = 40


# --------------------------------------------------------------------------
# 内部数据视图
# --------------------------------------------------------------------------
class _RowInfo:
    """一行的工作表视图（锚点、被合并覆盖的列、数字/公式/文本计数）。"""

    __slots__ = ("row", "anchors", "covered_cols", "numeric", "formulas", "texts")

    def __init__(self, row: int) -> None:
        self.row = row
        self.anchors: dict[int, str] = {}
        self.covered_cols: set[int] = set()
        self.numeric = 0
        self.formulas = 0
        self.texts = 0

    @property
    def anchor_count(self) -> int:
        return len(self.anchors)

    @property
    def anchor_cols(self) -> list[int]:
        return sorted(self.anchors)

    @property
    def covered_count(self) -> int:
        return len(self.covered_cols)


class _Grid:
    """``SheetModel`` 的只读视图：把合并区的值按需传播，缓存每行统计。"""

    def __init__(self, sheet: SheetModel) -> None:
        self.sheet = sheet
        self.cells: dict[tuple[int, int], Cell] = {}
        self.disp: dict[tuple[int, int], str] = {}
        self.by_row: dict[int, dict[int, str]] = {}
        for key, cell in dict(getattr(sheet, "cells", None) or {}).items():
            try:
                r, c = int(key[0]), int(key[1])
            except (TypeError, ValueError, IndexError):
                continue
            if r < 0 or c < 0:
                continue
            text = cell_display(cell)
            if not text or not text.strip():
                continue
            self.cells[(r, c)] = cell
            self.disp[(r, c)] = text
            self.by_row.setdefault(r, {})[c] = text
        self.r_max = max((r for r, _ in self.cells), default=-1)
        self.c_max = max((c for _, c in self.cells), default=-1)
        self.merges: list[tuple[int, int, int, int]] = []
        self.owner: dict[tuple[int, int], tuple[int, int]] = {}
        self.span: dict[tuple[int, int], tuple[int, int, int, int]] = {}
        for item in list(getattr(sheet, "merges", None) or []):
            try:
                r1, r2, c1, c2 = (int(x) for x in item[:4])
            except (TypeError, ValueError):
                continue
            if r1 > r2:
                r1, r2 = r2, r1
            if c1 > c2:
                c1, c2 = c2, c1
            self.merges.append((r1, r2, c1, c2))
            self.span[(r1, c1)] = (r1, r2, c1, c2)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self.owner[(r, c)] = (r1, c1)
        self._row_cache: dict[int, _RowInfo] = {}

    # -- 取值 ---------------------------------------------------------
    def display(self, r: int, c: int) -> str:
        """合并传播后的展示文本：被合并覆盖的位置返回左上角锚点的值。"""
        text = self.disp.get((r, c))
        if text is not None:
            return text
        anchor = self.owner.get((r, c))
        if anchor:
            return self.disp.get(anchor, "")
        return ""

    def anchor_display(self, r: int, c: int) -> str:
        return self.disp.get((r, c), "")

    # -- 行统计 -------------------------------------------------------
    def row_info(self, r: int) -> _RowInfo:
        info = self._row_cache.get(r)
        if info is not None:
            return info
        info = _RowInfo(r)
        for cc, text in self.by_row.get(r, {}).items():
            info.anchors[cc] = text
            cell = self.cells.get((r, cc))
            if getattr(cell, "formula", None) or text.lstrip().startswith("="):
                info.formulas += 1
            kind = getattr(cell, "kind", "") or ""
            if kind == KIND_NUMBER or kind == KIND_DATE:
                info.numeric += 1
            elif kind == KIND_FORMULA:
                # 公式能取到缓存值就算数字，取不到算文本
                if getattr(cell, "num", None) is not None or _QUICK_NUM_RE.match(text.replace(" ", "")):
                    info.numeric += 1
                else:
                    info.texts += 1
            else:
                info.texts += 1
        for (r1, r2, c1, c2) in self.merges:
            if r1 <= r <= r2:
                for c in range(c1, c2 + 1):
                    if (r, c) not in self.disp:
                        info.covered_cols.add(c)
        self._row_cache[r] = info
        return info


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------
def _clean_label(text: Any) -> str:
    """表头/标题文本：换行与多余空白压成一个空格，去掉首尾空白。"""
    if text is None:
        return ""
    s = str(text)
    s = re.sub(r"[\s\u200b\ufeff\u00a0]+", " ", s)
    return s.strip()


def _short(text: str, limit: int = 40) -> str:
    text = _clean_label(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _resolve_source(wb: WorkbookModel, file_ref: FileRef | None) -> FileRef | None:
    if file_ref is not None:
        return file_ref
    path = getattr(wb, "path", "") or ""
    if not path:
        return None
    return FileRef(
        id="",
        path=path,
        name=os.path.basename(path),
        fmt=getattr(wb, "fmt", "") or "",
    )


def _row_is_blank(grid: _Grid, r: int) -> bool:
    """整行既没有锚点单元格、也没有被任何合并区覆盖 → 真空行。"""
    info = grid.row_info(r)
    return info.anchor_count == 0 and info.covered_count == 0


def _row_cols(grid: _Grid, r: int) -> set[int]:
    """该行涉及的所有列（锚点列 + 合并覆盖列）。"""
    cols = set(grid.row_info(r).anchor_cols)
    for (r1, r2, c1, c2) in grid.merges:
        if r1 <= r <= r2:
            cols.update(range(c1, c2 + 1))
    return cols


# --------------------------------------------------------------------------
# 标题行 / 表头块 / 数据区
# --------------------------------------------------------------------------
def _find_title_rows(grid: _Grid) -> list[tuple[int, str]]:
    """找表格上方的大标题行。返回 [(0 基行号, 标题文本), ...]。"""
    titles: list[tuple[int, str]] = []
    limit = min(grid.r_max, MAX_TITLE_SCAN)
    r = 0
    while r <= limit and len(titles) < MAX_TITLE_ROWS:
        if _row_is_blank(grid, r):
            if titles:
                break
            r += 1
            continue
        info = grid.row_info(r)
        if info.anchor_count != 1:
            break
        c = info.anchor_cols[0]
        text = _clean_label(grid.display(r, c))
        if not text:
            break
        # 纯数字短行更像数据，不像标题
        if len(text) <= 12 and to_number(text) is not None:
            break
        span = grid.span.get((r, c))
        merged_wide = span is not None and (span[3] - span[2]) >= 1
        lead_ok = False
        for k in (1, 2):
            rr = r + k
            if rr > grid.r_max:
                break
            n = grid.row_info(rr).anchor_count
            if n >= 3:
                lead_ok = True
                break
            if n != 1:
                break
        first = not titles
        # 首行标题要有"合并跨列"或"较长文本 + 后面紧跟多列表头"的实证；
        # 后续行（副标题，如"（单位：元）"）只要后面紧跟多列表头即可。
        if merged_wide or (lead_ok and (not first or len(text) >= 6)):
            titles.append((r, text))
            r += 1
            continue
        break
    return titles


def _first_content_row(grid: _Grid, start: int) -> int | None:
    for r in range(max(0, start), grid.r_max + 1):
        if not _row_is_blank(grid, r):
            return r
    return None


def _expand_header(grid: _Grid, r0: int) -> list[int]:
    """从表头首行出发，用纵向合并覆盖扩展出所有表头行。"""
    rows = {r0}
    cap = min(grid.r_max, r0 + MAX_HEADER_ROWS - 1)
    changed = True
    while changed:
        changed = False
        for (r1, r2, c1, c2) in grid.merges:
            if r2 <= r1 or r1 not in rows:
                continue
            for r in range(r1 + 1, min(r2, cap) + 1):
                if r not in rows:
                    rows.add(r)
                    changed = True
    return sorted(rows)


def _looks_like_pure_data(info: _RowInfo) -> bool:
    """首行就是数据（而不是表头）的判断。

    证据一：纯数值矩阵（``numeric>=2`` 且没有任何文本）；
    证据二：出现长文本（表头标签通常很短，数据里的标题/内容/描述会很长）。
        很多系统导出的明细页根本没有表头，首行就是数据。
    """
    if info.numeric >= 2 and info.texts == 0 and info.anchor_count >= 2:
        return True
    long_texts = [t for t in info.anchors.values() if len(t) >= _LONG_TEXT]
    return len(long_texts) >= 2 or any(len(t) >= _LONG_TEXT * 2 for t in long_texts)


# --------------------------------------------------------------------------
# 行分类
# --------------------------------------------------------------------------
def _subtotal_label(text: Any) -> bool:
    if not text:
        return False
    t = str(text).strip()
    if not t or len(t) > 14:  # 合计标签都很短
        return False
    low = t.lower()
    if not any(w in low for w in _SUBTOTAL_WORDS):
        return False
    t = normalize_text(t)
    t = _SUBTOTAL_STRIP_RE.sub("", t)
    if not t:
        return False
    for word in _SUBTOTAL_WORDS:
        if t == word:
            return True
        for suffix in _SUBTOTAL_SUFFIXES:
            if t == word + suffix:
                return True
    return False


def _is_subtotal_row(info: _RowInfo, label_max_col: int) -> bool:
    """小计/合计行：左前方有合计标签，或整行以汇总公式为主。"""
    for c in info.anchor_cols:
        if c > label_max_col:
            continue
        if _subtotal_label(info.anchors[c]):
            return True
    if info.formulas >= 3 and info.formulas >= info.anchor_count * 0.6:
        return True
    return False


def _is_repeat_header(
    info: _RowInfo,
    header_labels: set[str],
    header_raw_labels: set[str] | None = None,
) -> bool:
    if not header_labels or info.anchor_count < 3 or info.numeric > 0:
        return False
    # 先用原文快速比对（重复表头是同一批文字复制出来的，原文通常完全一致）
    if header_raw_labels:
        raw_hits = sum(1 for t in info.anchors.values() if t in header_raw_labels)
        if raw_hits < 3:
            return False
    texts = {normalize_text(info.anchors[c]) for c in info.anchor_cols} - {""}
    if len(texts) < 3:
        return False
    hit = len(texts & header_labels)
    return hit >= 3 and hit >= int(len(texts) * 0.8)


def _is_placeholder_row(info: _RowInfo, n_fields: int, placeholder_max: int) -> bool:
    """分类占位行：宽表里只有最前面 1~2 列有文本、没有任何数据。"""
    if n_fields < 10:
        return False
    if info.numeric > 0:
        return False
    if info.anchor_count > placeholder_max:
        return False
    return all(c <= 2 for c in info.anchor_cols)


# --------------------------------------------------------------------------
# 字段类型细化
# --------------------------------------------------------------------------
def _mostly_date_numbers(values: list[Any]) -> bool:
    nums: list[float] = []
    for v in list(values)[:200]:
        if is_blank(v):
            continue
        n = to_number(v)
        if n is None:
            return False
        nums.append(n)
    if not nums:
        return False
    ok = 0
    for n in nums:
        if to_date_text(n) is not None:
            ok += 1
        elif float(n).is_integer() and to_date_text(int(n)) is not None:
            ok += 1
    return ok / len(nums) >= 0.8


def _refine_kind(kind: str, name: str, values: list[Any]) -> str:
    """用字段名把 ``number`` 细化为 ``money`` / ``date``（只改 hint，不改数据）。"""
    norm_name = normalize_text(name)
    if kind != "number" or not norm_name:
        return kind
    last = normalize_text(name.split(" / ")[-1]) or norm_name
    if any(h in last for h in _RATE_HINTS) or any(h in last for h in _COUNT_HINTS):
        return kind
    if any(h in norm_name for h in _MONEY_HINTS):
        return "money"
    if any(h in norm_name for h in _DATE_HINTS) and _mostly_date_numbers(values):
        return "date"
    return kind


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def detect_table(
    wb: WorkbookModel,
    sheet: SheetModel,
    file_ref: FileRef | None = None,
) -> TableModel:
    """识别一个工作表里的逻辑表。空表返回空 ``TableModel``（不抛异常）。"""
    grid = _Grid(sheet)
    table = TableModel(source=_resolve_source(wb, file_ref), sheet=str(getattr(sheet, "name", "") or ""))

    if grid.r_max < 0:
        table.parse_notes.append("空表：没有任何非空单元格，已跳过")
        return table

    # 1) 大标题行
    titles = _find_title_rows(grid)
    for r, text in titles:
        table.parse_notes.append(f"大标题行：第 {r + 1} 行「{_short(text)}」")

    # 2) 表头块
    body_start = _first_content_row(grid, (titles[-1][0] + 1) if titles else 0)
    if body_start is None:
        table.parse_notes.append("只有标题、没有内容区，未生成字段")
        return table

    first_info = grid.row_info(body_start)
    if _looks_like_pure_data(first_info):
        header_rows: list[int] = []
        data_start = body_start
        table.parse_notes.append("未识别到表头行（首行已是数据：长文本/纯数值），按无表头解析")
    else:
        header_rows = _expand_header(grid, body_start)
        data_start = (header_rows[-1] + 1) if header_rows else body_start

    # 表头与数据之间最多跳过 SPACER_SKIP 行空行（不追到很远的其它内容）
    probe = data_start
    skipped = 0
    while probe <= grid.r_max and skipped < SPACER_SKIP and _row_is_blank(grid, probe):
        probe += 1
        skipped += 1
    if skipped and probe <= grid.r_max and not _row_is_blank(grid, probe):
        data_start = probe
        table.parse_notes.append(f"表头与数据之间跳过 {skipped} 行空行")

    # 3) 数据区结束行 = 数据区最后一个非空行
    last_content = -1
    for r in range(max(0, data_start), grid.r_max + 1):
        if not _row_is_blank(grid, r):
            last_content = r
    data_end = last_content if last_content >= data_start else data_start - 1

    # 4) 字段列 = 表头涉及列 ∪ 数据区涉及列
    header_cols: set[int] = set()
    for r in header_rows:
        header_cols |= _row_cols(grid, r)
    data_cols: set[int] = set()
    for r in range(max(0, data_start), data_end + 1):
        data_cols.update(grid.row_info(r).anchor_cols)
    all_cols = sorted(header_cols | data_cols)

    if not all_cols:
        table.parse_notes.append("没有可识别的列，未生成字段")
        return table

    n_fields = len(all_cols)
    label_max_col = min(4, max(2, n_fields // 10))
    placeholder_max = max(2, min(4, n_fields // 10))

    # 5) 逐行分类
    header_labels: set[str] = set()
    header_raw_labels: set[str] = set()
    for r in header_rows:
        for c in _row_cols(grid, r):
            raw = _clean_label(grid.display(r, c))
            if raw:
                header_raw_labels.add(raw)
                header_labels.add(normalize_text(raw))

    rows: list[dict[int, Cell]] = []
    row_numbers: list[int] = []
    excluded: list[tuple[int, str]] = []

    # 表单/标签-值布局（如勘测定界表）几乎每行都只有最前面一两列有字，
    # 如果"分类占位行"占了一半以上，说明这不是分类占位而是表单标签，不能剔除。
    placeholder_ok = True
    if n_fields >= 10:
        nonblank = candidates = 0
        for r in range(max(0, data_start), data_end + 1):
            info = grid.row_info(r)
            if info.anchor_count == 0:
                continue
            nonblank += 1
            if _is_placeholder_row(info, n_fields, placeholder_max):
                candidates += 1
        if candidates and candidates * 2 > nonblank:
            placeholder_ok = False
            table.parse_notes.append(
                f"疑似表单/标签-值布局（{candidates}/{nonblank} 行仅少量列有值），"
                f"不按分类占位行剔除"
            )

    for r in range(max(0, data_start), data_end + 1):
        info = grid.row_info(r)
        if info.anchor_count == 0:
            if info.covered_count:
                excluded.append((r, "合并区延续行（本行无独立单元格）"))
            else:
                excluded.append((r, "空行"))
            continue
        if _is_subtotal_row(info, label_max_col):
            excluded.append((r, "小计/合计行（合计标签或汇总公式）"))
            continue
        if _is_repeat_header(info, header_labels, header_raw_labels):
            excluded.append((r, "重复表头行"))
            continue
        if placeholder_ok and _is_placeholder_row(info, n_fields, placeholder_max):
            excluded.append((r, "分类占位行（仅前几列有文本、无数据）"))
            continue
        rows.append({c: grid.cells[(r, c)] for c in all_cols if (r, c) in grid.cells})
        row_numbers.append(r)

    # 6) 字段（多级表头 "父级 / 子级"，跨列父级通过合并传播到子列）
    fields: list[Field] = []
    for c in all_cols:
        parts: list[str] = []
        for r in header_rows:
            txt = _clean_label(grid.display(r, c))
            if not txt:
                continue
            if parts and normalize_text(parts[-1]) == normalize_text(txt):
                continue  # 纵向合并的父级在子行重复出现 → 折叠
            parts.append(txt)
        name = " / ".join(parts) if parts else f"列{col_letter(c + 1)}"
        values = [row.get(c) for row in rows][:MAX_FIELD_SAMPLE]
        kind = _refine_kind(guess_kind(values), name, values)
        fields.append(
            Field(index=c, name=name, path=name, level=max(1, len(parts)), kind_hint=kind, col_span=(c, c))
        )

    # 7) 落盘到 TableModel + 判据
    table.header_rows = list(header_rows)
    table.data_start = max(0, data_start)
    table.data_end = data_end
    table.fields = fields
    table.rows = rows
    table.row_numbers = row_numbers
    table.excluded_rows = excluded

    if header_rows:
        table.parse_notes.append(
            f"表头行：第 {', '.join(str(r + 1) for r in header_rows)} 行"
            f"（0 基 {header_rows}，共 {len(header_rows)} 行）"
        )
    else:
        table.parse_notes.append("无表头行：字段名按列字母生成")
    if data_end >= data_start:
        table.parse_notes.append(
            f"数据区：第 {data_start + 1} 行–第 {data_end + 1} 行"
            f"（0 基 {data_start}–{data_end}），有效 {len(rows)} 行，排除 {len(excluded)} 行"
        )
    else:
        table.parse_notes.append(f"数据区：未找到数据行（表头结束于 0 基 {data_start - 1}）")
    if excluded:
        counts: dict[str, int] = {}
        for _, reason in excluded:
            counts[reason] = counts.get(reason, 0) + 1
        table.parse_notes.append(
            "排除明细：" + "、".join(f"{k} {v} 行" for k, v in counts.items())
        )
    multi = sum(1 for f in fields if f.level > 1)
    table.parse_notes.append(
        f"字段：{len(fields)} 列" + (f"，多级表头 {multi} 列" if multi else "")
    )
    return table


def detect_workbook(
    wb: WorkbookModel,
    file_ref: FileRef | None = None,
) -> list[TableModel]:
    """识别整本工作簿：逐 sheet 产出 ``TableModel``，跳过空表/无意义表。"""
    tables: list[TableModel] = []
    for sheet in list(getattr(wb, "sheets", None) or []):
        try:
            table = detect_table(wb, sheet, file_ref)
        except Exception as exc:  # 单个工作表失败不影响其它表
            notes = getattr(wb, "notes", None)
            if isinstance(notes, list):
                notes.append(f"工作表 {getattr(sheet, 'name', '?')} 结构识别失败：{exc}")
            continue
        if not table.fields and not table.rows:
            continue  # 空表
        if not table.rows and len(table.fields) <= 1:
            continue  # 只有一个孤零零的单元格，不算表
        if getattr(sheet, "hidden", False):
            table.parse_notes.append("隐藏工作表")
        tables.append(table)
    return tables
