# -*- coding: utf-8 -*-
"""坐标化字段映射引擎（CONTRACT v1.0 · S4 / D12）。

规则模型（``MappingRule``）里的行/列一律是 **1 基**（与 Excel 界面一致）；
``apply_rules`` 返回的键是 **0 基** ``(target_row, target_col)``（与 ``SheetModel`` 一致）。

五种 mode
---------
* ``cell``      —— 一格 → 一格（默认；配 ``record`` 选择器即"列→列、记录逐条下移"）。
* ``block``     —— 源区域块 → 目标区域块（保持方向，尺寸相同）。
* ``transpose`` —— 源区域块 → 目标区域块，**行列互换**。
                   例："A 表第 3 行横排 → B 表 F 列竖排"：
                   ``source.row=fixed:3, source.col=span:{start:2,count:30}``
                   ``target.row=span:{start:4,count:30}, target.col=fixed:6``
* ``filldown``  —— 源列值向下填充（合并单元格只在首行有值）后写入目标格。
* ``repeat``    —— 常量重复写入（``constant`` 或 ``transform="constant"``）。

四种 Sel.kind
-------------
* ``fixed``  —— 1 基绝对行号/列号。
* ``record`` —— 相对当前记录行偏移（0 = 当前记录行；目标侧 0 = 数据起始行）。
* ``header`` —— 列选择器按来源/模板字段名解析（行选择器按表头行偏移）。
* ``span``   —— 区域映射：值可以是 ``N``（个数）或 ``{"start": 起始1基, "count": 个数}``；
                也可在 ``source``/``target`` 字典里用额外的 ``rows``/``cols`` 指定区域尺寸。

规则会在同一方案里按列表顺序执行；``cell/block/transpose`` 中只要有一条选择了
``record``，就按"每条记录执行一次"计算，否则整条规则只执行一次。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from difflib import SequenceMatcher
from typing import Any

from .model import (
    Cell,
    CellRef,
    Field,
    MappedCell,
    MappingProfile,
    MappingRule,
    Sel,
    TableModel,
    TemplateDef,
    col_letter,
)
from . import template as template_mod

MODES = ("cell", "block", "transpose", "filldown", "repeat")
SEL_KINDS = ("fixed", "record", "header", "span")
TRANSFORMS = ("identity", "number", "date", "concat", "constant")
ON_MISSING = ("blank", "skip", "error")


# ==========================================================================
# dict ←→ 数据类
# ==========================================================================
def sel_from_dict(d: Any, default_kind: str = "fixed", default_value: Any = 1) -> Sel:
    if isinstance(d, Sel):
        return d
    if d is None:
        return Sel(kind=default_kind, value=default_value)
    if isinstance(d, (int, str)):
        return Sel(kind=default_kind, value=d)
    if not isinstance(d, dict):
        raise ValueError(f"选择器必须是对象，收到 {type(d).__name__}")
    kind = str(d.get("kind") or default_kind)
    if kind not in SEL_KINDS:
        raise ValueError(f"未知选择器类型 kind={kind!r}（可选：{'/'.join(SEL_KINDS)}）")
    value = d.get("value", d.get("start", default_value))
    return Sel(kind=kind, value=value)


def _side_from_dict(d: Any) -> dict:
    if not isinstance(d, dict):
        d = {}
    return {
        "sheet": d.get("sheet"),
        "row": sel_from_dict(d.get("row"), "fixed", 1),
        "col": sel_from_dict(d.get("col"), "fixed", 1),
        "_rows": d.get("rows"),
        "_cols": d.get("cols"),
        "_raw": d,
    }


def rule_from_dict(d: Any) -> MappingRule:
    if isinstance(d, MappingRule):
        return d
    if not isinstance(d, dict):
        raise ValueError("映射规则必须是对象")
    mode = str(d.get("mode") or "cell")
    if mode not in MODES:
        raise ValueError(f"未知映射模式 mode={mode!r}（可选：{'/'.join(MODES)}）")
    transform = str(d.get("transform") or "identity")
    if transform not in TRANSFORMS:
        raise ValueError(f"未知转换 transform={transform!r}")
    on_missing = str(d.get("onMissing") or d.get("on_missing") or "blank")
    if on_missing not in ON_MISSING:
        raise ValueError(f"未知 onMissing={on_missing!r}")
    src = _side_from_dict(d.get("source"))
    tgt = _side_from_dict(d.get("target"))
    return MappingRule(
        source=src,
        target=tgt,
        mode=mode,
        transform=transform,
        on_missing=on_missing,
        constant=d.get("constant"),
        note=str(d.get("note") or ""),
    )


def profile_from_dict(d: Any) -> MappingProfile:
    if isinstance(d, MappingProfile):
        return d
    if not isinstance(d, dict):
        raise ValueError("映射方案必须是对象")
    rules = [rule_from_dict(r) for r in (d.get("rules") or [])]
    return MappingProfile(
        id=str(d.get("id") or ""),
        name=str(d.get("name") or "未命名方案"),
        template_id=str(d.get("templateId") or d.get("template_id") or ""),
        source_signature=str(d.get("sourceSignature") or d.get("source_signature") or ""),
        rules=rules,
        notes=str(d.get("notes") or ""),
    )


# ==========================================================================
# 单元格取值 / 类型转换
# ==========================================================================
def cell_value(cell: Cell | None) -> Any:
    """从契约 Cell 取"可写入模板"的原值。"""
    if cell is None:
        return None
    raw = cell.raw
    if raw is not None:
        return raw
    if cell.num is not None:
        return cell.num
    if cell.kind in ("blank", ""):
        return None
    if cell.formula:
        return cell.text
    return cell.text if cell.text != "" else None


def _to_number(v: Any) -> tuple[Any, str | None]:
    if isinstance(v, bool):
        return int(v), None
    if isinstance(v, (int, float)):
        return v, None
    if v is None:
        return None, "空值无法转数字"
    s = unicodedata.normalize("NFKC", str(v)).strip()
    if not s:
        return None, "空值无法转数字"
    percent = s.endswith("%")
    for ch in (",", "，", " ", "\u3000", "￥", "¥", "元"):
        s = s.replace(ch, "")
    if percent:
        s = s[:-1] if s.endswith("%") else s
    try:
        f = float(s)
    except ValueError:
        return v, f"无法把 {v!r} 转成数字"
    if percent:
        f = f / 100.0
    return (int(f) if f.is_integer() else f), None


_DATE_RE = re.compile(
    r"^(\d{4})\s*[.\-/年]\s*(\d{1,2})\s*[.\-/月]\s*(\d{1,2})\s*日?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$"
)
_COMPACT_DATE_RE = re.compile(r"^(\d{4})(\d{2})(\d{2})$")


def _to_date(v: Any) -> tuple[Any, str | None]:
    if isinstance(v, (datetime, date)):
        return template_mod.date_to_serial(v), None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v, None  # 已经是 Excel 序列值
    if v is None:
        return None, "空值无法转日期"
    s = unicodedata.normalize("NFKC", str(v)).strip()
    if not s:
        return None, "空值无法转日期"
    m = _DATE_RE.match(s)
    if not m:
        m2 = _COMPACT_DATE_RE.match(s)
        if m2:
            y, mo, d = (int(x) for x in m2.groups())
            hh = mm = ss = 0
        else:
            return v, f"无法把 {v!r} 转成日期（支持 2024.3.9 / 2024-03-09 / 2024年3月9日）"
    else:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        hh = int(m.group(4) or 0)
        mm = int(m.group(5) or 0)
        ss = int(m.group(6) or 0)
    try:
        dt = datetime(y, mo, d, hh, mm, ss)
    except ValueError:
        return v, f"非法日期：{v!r}"
    return template_mod.date_to_serial(dt), None


def _apply_transform(rule: MappingRule, value: Any) -> tuple[Any, str | None]:
    t = rule.transform
    if t == "constant":
        return rule.constant, None
    if t == "number":
        return _to_number(value)
    if t == "date":
        return _to_date(value)
    return value, None


# ==========================================================================
# 选择器解析
# ==========================================================================
def _as_int(v: Any, default: int = 1) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _field_col(source: TableModel, name: str) -> int | None:
    """按字段名找 1 基列号（优先级：全名精确 → 归一化全名 → 叶子名精确 → 叶子名归一 → 叶子包含 → 全名包含）。"""
    return _match_field(source.fields, name)


def _tpl_field_col(template: TemplateDef, name: str) -> int | None:
    """在模板字段里按名字找 1 基列号（同 ``_field_col`` 的优先级）。"""
    return _match_field(template.fields, name)


def _leaf_raw(name: str) -> str:
    """取多级字段名的最后一级（保留原始写法），用于精确匹配。"""
    parts = [p.strip() for p in re.split(r"\s*/\s*", str(name or "").strip()) if p.strip()]
    return parts[-1] if parts else ""


def _match_field(fields: list[Field], name: str) -> int | None:
    """字段名 → 1 基列号。

    关键点：不能"按字段顺序取第一个包含者"。模板第 5 个字段
    ``委托单位/甲方（插入评论合同编码）`` 也包含"合同编码"，按顺序取会错配到 E 列；
    正确做法是优先精确匹配叶子字段名（``合同签订情况 / 合同编码`` → AF 列），
    包含匹配时取"叶子名最短"的候选。
    """
    if not fields:
        return None
    raw = str(name or "").strip()
    if not raw:
        return None
    want = _norm_name(raw)
    want_leaf = _leaf_name(want)
    raw_leaf = _leaf_raw(raw)
    # 1) 全名精确
    for f in fields:
        if (f.name or "").strip() == raw:
            return f.index + 1
    # 2) 归一化全名相等
    for f in fields:
        if _norm_name(f.name) == want:
            return f.index + 1
    # 3) 叶子字段名精确（原始写法）
    for f in fields:
        if raw_leaf and _leaf_raw(f.name) == raw_leaf:
            return f.index + 1
    # 4) 叶子字段名归一化相等
    for f in fields:
        if want_leaf and _leaf_name(_norm_name(f.name)) == want_leaf:
            return f.index + 1
    # 5) 叶子名包含：多个候选时取叶子名最短的（越短越可能是目标字段本身）
    cands: list[tuple[int, int, int]] = []
    for f in fields:
        full = _norm_name(f.name)
        leaf = _leaf_name(full)
        if not want_leaf or not leaf:
            continue
        if want_leaf in leaf or leaf in want_leaf:
            cands.append((len(leaf), len(full), f.index + 1))
    if cands:
        cands.sort()
        return cands[0][2]
    # 6) 全名包含：同样取最短
    cands = []
    for f in fields:
        full = _norm_name(f.name)
        if not full or not want:
            continue
        if want in full or full in want:
            cands.append((len(full), f.index + 1))
    if cands:
        cands.sort()
        return cands[0][1]
    return None


def _source_first_col(source: TableModel) -> int:
    if source.fields:
        return min(f.index for f in source.fields) + 1
    return 1


def _span_parts(sel: Sel, default_start: int) -> tuple[int, int]:
    v = sel.value
    if isinstance(v, dict):
        start = _as_int(v.get("start", v.get("from", default_start)), default_start)
        count = _as_int(v.get("count", v.get("n", 1)), 1)
        return start, max(1, count)
    if isinstance(v, (list, tuple)) and len(v) == 2:
        return _as_int(v[0], default_start), max(1, _as_int(v[1], 1))
    return default_start, max(1, _as_int(v, 1))


def _anchor_row(
    sel: Sel,
    side: dict,
    *,
    is_source: bool,
    record_index: int | None,
    source: TableModel | None,
    template: TemplateDef,
    stride: int,
    warnings: list[str],
    rule_no: int,
) -> int | None:
    kind = sel.kind
    if kind == "fixed":
        return _as_int(sel.value, 1)
    if kind == "record":
        offset = _as_int(sel.value, 0)
        if record_index is None:
            return None
        if is_source:
            assert source is not None
            base = (source.row_numbers[record_index] + 1) if record_index < len(source.row_numbers) else (
                source.data_start + 1 + record_index
            )
            return base + offset
        return template.data_start[0] + record_index * max(1, stride) + offset
    if kind == "header":
        if is_source:
            assert source is not None
            base = (source.header_rows[0] + 1) if source.header_rows else (source.data_start + 1)
        else:
            base = template.header_range[0]
        return base + _as_int(sel.value, 0)
    if kind == "span":
        default = (source.data_start + 1) if (is_source and source is not None) else template.data_start[0]
        return _span_parts(sel, default)[0]
    warnings.append(f"规则{rule_no}：未知行选择器 kind={kind!r}，已跳过")
    return None


def _anchor_col(
    sel: Sel,
    side: dict,
    *,
    is_source: bool,
    source: TableModel | None,
    template: TemplateDef,
    warnings: list[str],
    rule_no: int,
) -> int | None:
    kind = sel.kind
    if kind == "fixed":
        return _as_int(sel.value, 1)
    if kind == "header":
        name = str(sel.value if isinstance(sel.value, str) else "")
        if is_source:
            assert source is not None
            col = _field_col(source, name)
            if col is None:
                warnings.append(f"规则{rule_no}：来源表中找不到字段 {name!r}")
            return col
        col = _tpl_field_col(template, name)
        if col is None:
            warnings.append(f"规则{rule_no}：模板中找不到目标字段 {name!r}")
        return col
    if kind == "record":
        offset = _as_int(sel.value, 0)
        if is_source:
            assert source is not None
            return _source_first_col(source) + offset
        return template.data_start[1] + offset
    if kind == "span":
        default = _source_first_col(source) if (is_source and source is not None) else template.data_start[1]
        return _span_parts(sel, default)[0]
    warnings.append(f"规则{rule_no}：未知列选择器 kind={kind!r}，已跳过")
    return None


def _axis_count(side: dict, sel: Sel, key: str) -> int:
    raw = side.get(key)
    if raw in (None, ""):
        raw = side.get(key.lstrip("_"))  # 兼容直接构造 MappingRule 时的 "rows"/"cols"
    if raw not in (None, ""):
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            pass
    if sel.kind == "span":
        v = sel.value
        if isinstance(v, dict):
            return max(1, _as_int(v.get("count", v.get("n", 1)), 1))
        if isinstance(v, (list, tuple)) and len(v) == 2:
            return max(1, _as_int(v[1], 1))
        return max(1, _as_int(v, 1))
    return 1


# ==========================================================================
# 来源数据访问
# ==========================================================================
def _lookup_record(source: TableModel, abs_row_1based: int) -> int | None:
    for i, r0 in enumerate(source.row_numbers):
        if r0 + 1 == abs_row_1based:
            return i
    if source.row_numbers:
        guess = abs_row_1based - 1 - source.data_start
        if 0 <= guess < len(source.row_numbers):
            return guess
    # 没有 row_numbers 时退化为按 rows 列表下标对齐数据区
    guess = abs_row_1based - 1 - source.data_start
    if 0 <= guess < len(source.rows):
        return guess
    return None


def _read_source(source: TableModel, abs_row: int, abs_col: int) -> tuple[Any, bool]:
    idx = _lookup_record(source, abs_row)
    if idx is None:
        return None, False
    row = source.rows[idx]
    cell = row.get(abs_col - 1)
    if cell is None:
        return None, abs_col - 1 in row
    return cell_value(cell), True


# ==========================================================================
# 规则执行引擎
# ==========================================================================
@dataclass
class _Out:
    target_row: int  # 1 基
    target_col: int  # 1 基
    value: Any
    rule_index: int
    source_row: int | None = None
    source_col: int | None = None


def _iter_record_indexes(source: TableModel, limit: int | None) -> list[int]:
    n = len(source.rows)
    if limit is not None:
        n = min(n, limit)
    return list(range(n))


def _rule_is_per_record(rule: MappingRule) -> bool:
    for side in (rule.source, rule.target):
        for sel in (side.get("row"), side.get("col")):
            if sel is not None and getattr(sel, "kind", "") == "record":
                return True
    return False


def _region_cells(
    rule: MappingRule,
    source: TableModel,
    template: TemplateDef,
    record_index: int | None,
    warnings: list[str],
    rule_no: int,
) -> tuple[list[tuple[int, int, Any, int | None, int | None, bool]], dict]:
    """计算一条规则要写的目标格。

    返回 ([(target_row1, target_col1, value, src_row1|None, src_col1|None, found)],
          {"per_record": bool, "truncated_hint": bool})
    """
    mode = rule.mode
    src = rule.source
    tgt = rule.target
    result: list[tuple[int, int, Any, int | None, int | None, bool]] = []
    per_record = _rule_is_per_record(rule)

    # 区域尺寸（block/transpose 用；cell/filldown/repeat 固定 1×1）
    if mode in ("block", "transpose"):
        src_nrows = _axis_count(src, src["row"], "_rows")
        src_ncols = _axis_count(src, src["col"], "_cols")
        tgt_nrows, tgt_ncols = (src_nrows, src_ncols) if mode == "block" else (src_ncols, src_nrows)
    else:
        src_nrows = src_ncols = 1
        tgt_nrows = tgt_ncols = 1

    stride = tgt_nrows if mode in ("block", "transpose") else 1
    # 目标行锚点依赖 stride；per_record 时 stride>=1
    s_row = _anchor_row(
        src["row"], src, is_source=True, record_index=record_index,
        source=source, template=template, stride=1, warnings=warnings, rule_no=rule_no,
    )
    s_col = _anchor_col(
        src["col"], src, is_source=True, source=source, template=template,
        warnings=warnings, rule_no=rule_no,
    )
    t_row = _anchor_row(
        tgt["row"], tgt, is_source=False, record_index=record_index,
        source=source, template=template, stride=stride, warnings=warnings, rule_no=rule_no,
    )
    t_col = _anchor_col(
        tgt["col"], tgt, is_source=False, source=source, template=template,
        warnings=warnings, rule_no=rule_no,
    )
    if s_row is None or s_col is None or t_row is None or t_col is None:
        return result, {"per_record": per_record, "error": "选择器无法解析（字段名或坐标不存在）"}

    # 跨工作表：v1 只支持单表注入，给出明确警告而不是静默错写
    if src.get("sheet") and source.sheet and str(src["sheet"]) != source.sheet:
        msg = f"规则{rule_no}：来源工作表 {src['sheet']!r} 与当前来源表 {source.sheet!r} 不一致，已跳过"
        warnings.append(msg)
        return result, {"per_record": per_record, "error": msg}
    if tgt.get("sheet") and template.sheet and str(tgt["sheet"]) != template.sheet:
        msg = f"规则{rule_no}：目标工作表 {tgt['sheet']!r} 与模板工作表 {template.sheet!r} 不一致，已跳过"
        warnings.append(msg)
        return result, {"per_record": per_record, "error": msg}
    if t_row < template.data_start[0]:
        warnings.append(
            f"规则{rule_no}：目标行 {t_row} 早于模板数据起始行 {template.data_start[0]}，生成时会被忽略"
        )

    # 读取来源值
    cells: list[tuple[int, int, Any, int, int, bool]] = []  # (dr, dc, value, abs_r, abs_c, found)
    for dr in range(src_nrows):
        for dc in range(src_ncols):
            abs_r = s_row + dr
            abs_c = s_col + dc
            value, found = _read_source(source, abs_r, abs_c)
            cells.append((dr, dc, value, abs_r, abs_c, found))

    transform = rule.transform
    if mode == "filldown":
        # 单个来源格；由调用方跨记录保持 last_seen 状态（见 _evaluate）
        pass
    if transform == "concat":
        joined = " ".join(
            template_mod._as_text(v) for _dr, _dc, v, _ar, _ac, _f in cells if v not in (None, "")
        )
        # concat 只写一个格（目标锚点）
        result.append((t_row, t_col, joined, s_row, s_col, bool(joined)))
        return result, {"per_record": per_record}

    for dr, dc, value, abs_r, abs_c, found in cells:
        if mode == "transpose":
            out_r, out_c = t_row + dc, t_col + dr
        elif mode == "block":
            out_r, out_c = t_row + dr, t_col + dc
        else:
            out_r, out_c = t_row, t_col
        if mode == "filldown":
            # 填充值由 _evaluate 跨记录补全后再做 transform
            result.append((out_r, out_c, value, abs_r, abs_c, found))
        else:
            out_val, err = _apply_transform(rule, value)
            if err:
                warnings.append(f"规则{rule_no}：{err}（来源 {abs_r}行/{col_letter(abs_c)}列），已按原值写入")
                out_val = value
            result.append((out_r, out_c, out_val, abs_r, abs_c, found))
        if mode in ("cell", "filldown", "repeat"):
            break  # 单格模式只写锚点
    return result, {"per_record": per_record}


def _evaluate(
    profile: MappingProfile,
    source: TableModel,
    template: TemplateDef,
    limit: int | None = None,
    strict: bool = False,
) -> tuple[list[_Out], list[str], bool]:
    """执行全部规则；返回 (输出格, 警告, 是否截断)。"""
    warnings: list[str] = []
    outs: list[_Out] = []
    if not isinstance(profile, MappingProfile):
        profile = profile_from_dict(profile)
    if not profile.rules:
        warnings.append("方案里没有映射规则")
        return outs, warnings, False

    total_records = len(source.rows)
    truncated = False
    # filldown 的状态按 (rule, 来源列) 保存，跨记录生效
    fill_state: dict[tuple[int, int], Any] = {}

    for ri, rule in enumerate(profile.rules):
        rule_no = ri + 1
        if not isinstance(rule, MappingRule):
            rule = rule_from_dict(rule)
        if rule.mode not in MODES:
            warnings.append(f"规则{rule_no}：未知模式 {rule.mode!r}，已跳过")
            continue
        per_record = _rule_is_per_record(rule)
        if rule.mode == "transpose" and not per_record:
            # 固定区域转置：整块只算一次
            rec_indexes: list[int | None] = [None]
        elif per_record:
            rec_indexes = list(_iter_record_indexes(source, limit))
            if total_records > len(rec_indexes):
                truncated = True
        else:
            rec_indexes = [None]

        for record_index in rec_indexes:
            cells, meta = _region_cells(rule, source, template, record_index, warnings, rule_no)
            if meta.get("error"):
                if strict:
                    raise ValueError(str(meta["error"]))
                break
            for out_r, out_c, value, abs_r, abs_c, found in cells:
                blank = value is None or value == ""
                # filldown：当前值缺失时沿用同列最近一次非空值
                if rule.mode == "filldown":
                    key = (ri, abs_c)
                    if not blank:
                        fill_state[key] = value
                    elif key in fill_state:
                        value = fill_state[key]
                        blank = False
                        found = True
                    if not blank:
                        value, err = _apply_transform(rule, value)
                        if err:
                            warnings.append(
                                f"规则{rule_no}：{err}（来源 {abs_r}行/{col_letter(abs_c)}列），已按原值写入"
                            )
                if blank:
                    if rule.on_missing == "skip":
                        continue
                    if rule.on_missing == "error":
                        msg = f"规则{rule_no}：缺少来源单元格（第{abs_r}行/{col_letter(abs_c)}列）"
                        if strict:
                            raise ValueError(msg)
                        warnings.append(msg)
                        continue
                    value = ""
                outs.append(
                    _Out(
                        target_row=out_r,
                        target_col=out_c,
                        value=value,
                        rule_index=ri,
                        source_row=abs_r,
                        source_col=abs_c,
                    )
                )
        if per_record and total_records > len(rec_indexes):
            # 后续规则不必重复告警
            pass

    # 目标坐标去重：同一格后写的规则覆盖先写的（与 Excel 粘贴顺序一致）
    merged: dict[tuple[int, int], _Out] = {}
    for out in outs:
        if out.target_row < 1 or out.target_col < 1:
            warnings.append(f"规则{out.rule_index + 1}：目标坐标非法（{out.target_row},{out.target_col}），已跳过")
            continue
        merged[(out.target_row - 1, out.target_col - 1)] = out
    final = list(merged.values())
    final.sort(key=lambda o: (o.target_row, o.target_col))
    # 警告去重并限量（大量行重复同一个问题时不刷屏）
    uniq: list[str] = []
    seen: set[str] = set()
    for w in warnings:
        if w not in seen:
            seen.add(w)
            uniq.append(w)
    if len(uniq) > 40:
        uniq = uniq[:40] + [f"……另有 {len(seen) - 40} 条警告未显示"]
    return final, uniq, truncated


# ==========================================================================
# 公共 API（CONTRACT §4 S4）
# ==========================================================================
def apply_rules(profile: MappingProfile, source: TableModel, template: TemplateDef) -> dict:
    """按规则计算所有目标值 → ``{(target_row_0based, target_col_0based): value}``。"""
    outs, _warnings, _trunc = _evaluate(profile_from_dict(profile), source, template, limit=None, strict=True)
    return {(o.target_row - 1, o.target_col - 1): o.value for o in outs}


def preview(
    profile: MappingProfile,
    source: TableModel,
    template: TemplateDef,
    limit: int = 50,
) -> dict:
    """逐单元格预演：来源坐标 → 目标坐标 → 值（前端确认后才生成）。"""
    try:
        lim = int(limit)
    except (TypeError, ValueError):
        lim = 50
    if lim <= 0:
        lim = 10 ** 9
    outs, warnings, truncated = _evaluate(
        profile_from_dict(profile), source, template, limit=lim, strict=False
    )
    cells: list[dict] = []
    for o in outs:
        target = CellRef(template.file_path, template.sheet, o.target_row, o.target_col)
        source_ref = None
        if o.source_row is not None and o.source_col is not None:
            source_ref = CellRef(source.path, source.sheet, o.source_row, o.source_col)
        cells.append(
            MappedCell(
                target=target,
                source_ref=source_ref,
                value=o.value,
                rule_index=o.rule_index,
            ).to_dict()
        )
    if not source.rows:
        warnings.append("来源表没有识别到数据行")
    if truncated:
        warnings.append(
            f"来源表共 {len(source.rows)} 行，预演只展示前 {lim} 行；生成时会写入全部数据"
        )
    return {
        "cells": cells,
        "truncated": truncated,
        "warnings": warnings,
        "count": len(cells),
        "recordCount": len(source.rows),
        "ruleCount": len(profile.rules),
        "targetSheet": template.sheet,
        "templateId": template.id,
    }


# ==========================================================================
# 自动推荐
# ==========================================================================
# 别名词典：同一组内的词可互相匹配（按最长词优先）
_ALIAS_GROUPS: list[tuple[str, ...]] = [
    ("合同编码", "合同编号", "合同号", "合同代码", "合同编码号"),
    ("联系电话", "手机号", "手机号码", "手机", "电话", "联系方式", "传真"),
    ("项目名称", "工程名称", "项目名", "工程名", "项目"),
    ("合同名称", "合同名"),
    ("中标单位名称", "中标单位", "中标供应商", "施工单位", "供应商", "公司名称", "单位名称"),
    ("项目负责人", "负责人", "项目经理", "联系人"),
    ("经办人", "经办", "办理人"),
    ("实施部门", "部门", "执行部门", "责任部门", "承办部门"),
    ("备注", "说明", "备注说明", "其他"),
    ("含税金额", "金额", "价款", "总金额", "价税合计", "合同金额"),
    ("税率", "税点", "税率%"),
    ("签订时间", "签约时间", "签订日期", "合同日期"),
    ("开工时间", "计划开工时间", "实际开工时间", "开工日期"),
    ("竣工时间", "计划竣工时间", "实际竣工时间", "竣工日期"),
    ("发票金额", "开票金额", "发票额", "开票额"),
    ("已付款金额", "付款金额", "已支付金额", "付款额", "已付款"),
    ("未付款金额", "欠款金额", "未支付金额"),
    ("招标方式", "采购方式", "招标类型"),
    ("序号", "顺序号", "编号", "序"),
    ("资金来源", "资金性质", "资金来源类别"),
]


def _norm_name(s: str) -> str:
    """字段名归一：NFKC、去括号单位、去空白/标点、小写；**保留 "/" 作为多级分隔符**。"""
    s = unicodedata.normalize("NFKC", str(s or "")).strip()
    # 去掉括号里的单位（元/%/天/个/项/人/份/次…）
    s = re.sub(r"[（(][^（()）]{0,10}(元|%|％|天|个|项|人|份|次|月|年|日)[^（()）]{0,10}[)）]", "", s)
    s = re.sub(r"[（(][^（()）]{0,2}[)）]", "", s)
    s = s.replace(" ", "").replace("\u3000", "")
    s = re.sub(r"[\s:：,，;；、.。\\|]+", "", s)
    return s.lower()


def _leaf_name(norm: str) -> str:
    """归一化字段名 → 最后一级（叶子）名；没有 "/" 时就是全名。"""
    if "/" in norm:
        return norm.rsplit("/", 1)[-1]
    return norm


def _has_alias(name: str, alias: str) -> bool:
    """别名命中：整名相等 / 末级字段名相等 / 长度 >=4 的子串。短词只认整名，避免误配。"""
    if not alias or not name:
        return False
    if name == alias or _leaf_name(name) == alias:
        return True
    return len(alias) >= 4 and alias in name


def _alias_label(a: str, b: str) -> str | None:
    for group in _ALIAS_GROUPS:
        aliases = sorted(group, key=len, reverse=True)
        hit_a = next((x for x in aliases if _has_alias(a, x)), None)
        hit_b = next((x for x in aliases if _has_alias(b, x)), None)
        if hit_a and hit_b:
            if hit_a == hit_b:
                return hit_a
            return hit_a if len(hit_a) >= len(hit_b) else hit_b
    return None


def _score(tf: Field, sf: Field) -> tuple[float, str]:
    """字段名相似度打分。优先级与 ``_match_field`` 一致：全名 > 归一全名 > 叶子名 > 别名 > 包含 > 相似度。"""
    t_raw, s_raw = str(tf.name or "").strip(), str(sf.name or "").strip()
    tn = _norm_name(t_raw)
    sn = _norm_name(s_raw)
    if not tn or not sn:
        return 0.0, ""
    tl, sl = _leaf_name(tn), _leaf_name(sn)
    if t_raw == s_raw:
        conf, reason = 0.98, "字段名完全一致"
    elif tn == sn:
        conf, reason = 0.96, "字段名完全一致"
    elif tl and tl == sl:
        conf, reason = 0.94, "叶子字段名一致"
    elif _alias_label(tn, sn):
        conf, reason = 0.85, f"别名词典：{_alias_label(tn, sn)}"
    elif tl and sl and (tl in sl or sl in tl):
        # 叶子名包含：越接近越长越可信；长父级/说明性长名会自然被压低
        shorter, longer = sorted((len(tl), len(sl)))
        ratio = shorter / max(1, longer)
        conf, reason = round(0.58 + 0.24 * ratio, 3), "叶子字段名包含"
    elif tn in sn or sn in tn:
        shorter, longer = sorted((len(tn), len(sn)))
        ratio = shorter / max(1, longer)
        conf, reason = round(0.46 + 0.18 * ratio, 3), "名称包含"
    else:
        ratio = max(SequenceMatcher(None, tn, sn).ratio(), SequenceMatcher(None, tl, sl).ratio())
        if ratio < 0.34:
            return 0.0, ""
        conf, reason = round(0.28 + 0.52 * ratio, 3), f"名称相似度 {ratio:.0%}"
    # 类型一致加分、明显冲突减分
    if tf.kind_hint and sf.kind_hint:
        if tf.kind_hint == sf.kind_hint:
            conf += 0.04
        elif tf.kind_hint != "text" and sf.kind_hint != "text":
            conf -= 0.06
    return max(0.0, min(1.0, round(conf, 3))), reason


def suggest(source: TableModel, template: TemplateDef) -> list[dict]:
    """自动推荐字段映射：名称精确/包含 + 别名词典 + 类型一致 + difflib 相似度。"""
    src_fields = [f for f in (source.fields or []) if (f.name or "").strip()]
    tpl_fields = [f for f in (template.fields or []) if (f.name or "").strip()]
    pairs: list[tuple[float, int, int, str]] = []
    for ti, tf in enumerate(tpl_fields):
        for si, sf in enumerate(src_fields):
            conf, reason = _score(tf, sf)
            if conf >= 0.30:
                pairs.append((conf, ti, si, reason))
    pairs.sort(key=lambda p: (-p[0], p[1], p[2]))
    used_t: set[int] = set()
    used_s: set[int] = set()
    chosen: dict[int, tuple[float, int, str]] = {}
    for conf, ti, si, reason in pairs:
        if ti in used_t or si in used_s:
            continue
        used_t.add(ti)
        used_s.add(si)
        chosen[ti] = (conf, si, reason)
    out: list[dict] = []
    for ti in sorted(chosen):
        conf, si, reason = chosen[ti]
        tf = tpl_fields[ti]
        sf = src_fields[si]
        out.append(
            {
                "targetField": tf.name,
                "sourceField": sf.name,
                "confidence": conf,
                "reason": reason,
                "targetIndex": tf.index,
                "sourceIndex": sf.index,
                "rule": {
                    "source": {
                        "sheet": None,
                        "row": {"kind": "record", "value": 0},
                        "col": {"kind": "header", "value": sf.name},
                    },
                    "target": {
                        "sheet": None,
                        "row": {"kind": "record", "value": 0},
                        "col": {"kind": "fixed", "value": tf.index + 1},
                    },
                    "mode": "cell",
                    "transform": "identity",
                    "onMissing": "blank",
                    "constant": None,
                    "note": f"自动推荐（{reason}，置信度 {conf:.0%}）",
                },
            }
        )
    return out


def source_signature(source: TableModel) -> str:
    """来源表头指纹（用于同类表自动套用方案）。"""
    import hashlib

    names = "|".join(_norm_name(f.name) for f in (source.fields or []) if (f.name or "").strip())
    return hashlib.sha1(names.encode("utf-8")).hexdigest()[:16] if names else ""


def template_signature(template: TemplateDef) -> str:
    import hashlib

    names = "|".join(_norm_name(f.name) for f in (template.fields or []) if (f.name or "").strip())
    return hashlib.sha1(names.encode("utf-8")).hexdigest()[:16] if names else ""
