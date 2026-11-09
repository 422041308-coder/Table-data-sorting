# -*- coding: utf-8 -*-
"""双表对比引擎（切片 S5）。

契约见 ``docs/CONTRACT.md`` 第 4 节 S5，数据类见 ``app/core/model.py``。

坐标约定（务必与契约一致）：

* ``keys`` 是 **1 基列号**（与 Excel 界面一致），支持多列组合键；
* ``TableModel.rows`` 的键是 **0 基列号**，``row_numbers`` 是 0 基工作表行号；
* 每条差异的 ``left_ref`` / ``right_ref`` 一律用
  ``CellRef.from_index(path, sheet, row0, col0)`` 生成（1 基行/列 + A1 + display），
  某侧整行不存在时该侧为 ``None``。

比较规则：

* 文本：先 ``normalize_text``（全角→半角、空白归一、大小写归一）再比；
  ``trim=False`` 时退化为原样比较（仅换行归一）；
* 数值：两侧都是数值单元格时按 ``tolerance`` 比较（默认 0，即精确）；
* ``filldown=True``（默认，符合需求 D 与文档"默认开启"）：空单元格沿用本列上方最近的非空值
  再参与主键匹配与字段比较，差异坐标指向值真正所在的单元格（合并区左上角），
  并在 ``warnings`` 里说明补全数量；
* ``ignoreEmpty=True``：任一侧为空的字段不产生差异（默认 False，空↔有值算修改）；
* 主键为空的行跳过并告警；主键重复的行取第一行参与匹配并告警，绝不静默错配。

返回结构（契约冻结）::

    {"summary": {"matched", "added", "removed", "changed", "leftRows", "rightRows"},
     "items": [DiffItem.to_dict(), ...],
     "keys": [1 基列号...],
     "warnings": ["..."]}

补充：``summary.changed`` 是**有修改的行数**，另附 ``summary.changedCells`` 是修改单元格数
（items 里每个 changed 为一行字段差异）。

本模块只依赖标准库与 ``model.py``；不 import 其他切片（``normalize`` 为契约允许的公共工具，
未就绪时用等价的本地实现兜底）。
"""
from __future__ import annotations

import math
import re
import unicodedata

from .model import CellRef, TableModel, col_letter

__all__ = ["DEFAULT_OPTIONS", "compare"]

DEFAULT_OPTIONS = {
    "trim": True,
    "tolerance": 0.0,
    "filldown": True,
    "ignoreEmpty": False,
}

# --------------------------------------------------------------------------
# 归一化 / 单元格展示：优先用 S2 的 app/core/normalize.py
# 待 S2 就绪后切换：已自动切换；S2 未交付时用本地等价实现兜底。
# --------------------------------------------------------------------------
try:  # pragma: no cover - 取决于 S2 是否已交付
    from .normalize import cell_display as _cell_display_ext
    from .normalize import to_number as _to_number_ext
except Exception:  # noqa: BLE001 - 模块可能尚未交付
    _cell_display_ext = None
    _to_number_ext = None

try:  # pragma: no cover - 取决于 S2 是否已交付
    from .normalize import normalize_text as _normalize_text_ext
except Exception:  # noqa: BLE001 - 模块可能尚未交付
    _normalize_text_ext = None

_WS_RE = re.compile(r"\s+")
_QUOTE_MAP = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"', "`": "'"})


def _local_normalize_text(value) -> str:
    """本地兜底归一：NFKC（全角→半角）+ 空白压缩 + 去首尾 + 小写 + 引号统一。"""
    if value is None:
        return ""
    s = unicodedata.normalize("NFKC", str(value))
    s = s.replace("\u3000", " ").translate(_QUOTE_MAP)
    s = s.strip()
    s = _WS_RE.sub(" ", s)
    return s.lower()


def _normalize(value) -> str:
    if _normalize_text_ext is not None:
        try:
            return _normalize_text_ext(value) or ""
        except Exception:  # noqa: BLE001 - 归一化异常不阻断对比
            return _local_normalize_text(value)
    return _local_normalize_text(value)


def _plain(value) -> str:
    """trim=False 时使用的原样比较文本（仅统一换行）。"""
    if value is None:
        return ""
    return str(value).replace("\r\n", "\n").replace("\r", "\n")


def _text_for_compare(value, trim: bool) -> str:
    return _normalize(value) if trim else _plain(value)


# --------------------------------------------------------------------------
# 单元格取值
# --------------------------------------------------------------------------
def _num_text(num: float) -> str:
    try:
        f = float(num)
    except (TypeError, ValueError):
        return str(num)
    if f != f or f in (float("inf"), float("-inf")):  # NaN / inf
        return str(num)
    if f.is_integer() and abs(f) < 1e16:
        return str(int(f))
    return repr(f)


def _raw_text(raw) -> str:
    if raw is None:
        return ""
    if isinstance(raw, bool):
        return "TRUE" if raw else "FALSE"
    if isinstance(raw, float):
        return _num_text(raw)
    return str(raw)


def _to_number(value):
    """宽松转数值（仅用于文本兜底，不改变数值单元格的判定）。"""
    if _to_number_ext is not None:
        try:
            return _to_number_ext(value)
        except Exception:  # noqa: BLE001
            pass
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("，", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _cell_eff(cell):
    """返回 (展示文本, 数值或 None)；公式有缓存值时用缓存值。

    S2 已交付时用 ``normalize.cell_display`` 统一展示（日期 2024.3.9、数字去 .0），
    否则用本地等价实现。
    """
    if cell is None:
        return "", None
    if _cell_display_ext is not None:
        text = _cell_display_ext(cell)
        num = cell.num
        if num is None and (cell.kind == "number"
                            or (isinstance(cell.raw, (int, float))
                                and not isinstance(cell.raw, bool))):
            num = _to_number(text)
        return text, num
    text = cell.text or ""
    num = cell.num
    if num is not None and (not text or text.startswith("=")):
        text = _num_text(num)
    if not text and num is not None:
        text = _num_text(num)
    if not text:
        text = _raw_text(cell.raw)
    if num is None and (cell.kind == "number" or (isinstance(cell.raw, (int, float))
                                                 and not isinstance(cell.raw, bool))):
        num = _to_number(text)
    return text, num


def _is_empty(text: str) -> bool:
    return _normalize(text) == ""


# --------------------------------------------------------------------------
# 参数校验
# --------------------------------------------------------------------------
def _to_key_col(value) -> int:
    if isinstance(value, bool):
        raise ValueError(f"非法主键列号：{value!r}（应为 1 基整数列号）")
    if isinstance(value, int):
        n = value
    elif isinstance(value, float) and value.is_integer():
        n = int(value)
    elif isinstance(value, str):
        s = value.strip()
        if s.isdigit():
            n = int(s)
        elif s.isalpha():
            from .model import col_index

            n = col_index(s)
        else:
            raise ValueError(f"非法主键列号：{value!r}（应为 1 基整数列号）")
    else:
        raise ValueError(f"非法主键列号：{value!r}（应为 1 基整数列号）")
    if n < 1:
        raise ValueError(f"主键列号必须 >= 1（1 基），收到 {n}")
    if n > 16384:
        raise ValueError(f"主键列号超出 Excel 上限：{n}")
    return n


def _normalize_keys(keys) -> list[int]:
    if not keys:
        raise ValueError("keys 不能为空：请指定 1 基主键列号")
    if not isinstance(keys, (list, tuple)):
        raise ValueError("keys 必须是 1 基列号数组，例如 [32] 或 [1, 2]")
    out: list[int] = []
    seen = set()
    for k in keys:
        n = _to_key_col(k)
        if n in seen:
            continue
        seen.add(n)
        out.append(n)
    if not out:
        raise ValueError("keys 不能为空：请指定 1 基主键列号")
    return out


def _normalize_options(options) -> dict:
    opts = dict(DEFAULT_OPTIONS)
    if options is None:
        return opts
    if not isinstance(options, dict):
        raise ValueError("options 必须是对象，例如 {\"trim\": true, \"tolerance\": 0.001}")
    aliases = {
        "ignore_empty": "ignoreEmpty",
        "ignoreempty": "ignoreEmpty",
        "fill_down": "filldown",
        "fillDown": "filldown",
    }
    for raw_key, value in options.items():
        key = aliases.get(raw_key, raw_key)
        if key not in opts:
            continue
        if key in ("trim", "filldown", "ignoreEmpty"):
            if isinstance(value, str):
                opts[key] = value.strip().lower() not in ("", "0", "false", "no", "off")
            else:
                opts[key] = bool(value)
        elif key == "tolerance":
            try:
                tol = float(value)
            except (TypeError, ValueError):
                raise ValueError(f"tolerance 必须是非负数值，收到 {value!r}")
            if tol < 0:
                raise ValueError(f"tolerance 不能为负，收到 {tol}")
            if not math.isfinite(tol):
                raise ValueError(f"tolerance 必须是有限数值，收到 {value!r}")
            opts[key] = tol
    return opts


# --------------------------------------------------------------------------
# 表结构 / 向下填充
# --------------------------------------------------------------------------
def _table_columns(table: TableModel) -> set[int]:
    cols: set[int] = set()
    for f in table.fields:
        if isinstance(f.index, int) and f.index >= 0:
            cols.add(f.index)
    for row in table.rows:
        cols.update(c for c in row.keys() if isinstance(c, int) and c >= 0)
    return cols


def _safe_row_numbers(table: TableModel, side: str, warnings: list[str]) -> list[int]:
    rn = list(table.row_numbers or [])
    if len(rn) != len(table.rows):
        warnings.append(
            f"{side}表 row_numbers 与 rows 长度不一致（{len(rn)} vs {len(table.rows)}），"
            "坐标退回逻辑行号"
        )
        rn = list(range(len(table.rows)))
    return rn


def _build_effective(table: TableModel, cols: set[int], filldown: bool):
    """构造每行的有效值：``eff[i][c0] = (text, num, origin_row0, origin_col0)``。

    空单元格在 filldown=True 时沿用本列上方最近的非空值，origin 指向值的真实坐标。
    """
    rownums = list(table.row_numbers or [])
    if len(rownums) != len(table.rows):
        rownums = list(range(len(table.rows)))
    eff: list[dict[int, tuple[str, float | None, int, int]]] = []
    filled = 0
    last: dict[int, tuple[str, float | None, int, int]] = {}
    for i, row in enumerate(table.rows):
        cur: dict[int, tuple[str, float | None, int, int]] = {}
        r0 = rownums[i]
        for c0 in cols:
            text, num = _cell_eff(row.get(c0))
            if _is_empty(text):
                if filldown and c0 in last:
                    cur[c0] = last[c0]
                    filled += 1
                else:
                    cur[c0] = ("", None, r0, c0)
            else:
                item = (text, num, r0, c0)
                cur[c0] = item
                last[c0] = item
        eff.append(cur)
    return eff, filled


# --------------------------------------------------------------------------
# 主键
# --------------------------------------------------------------------------
def _key_parts(eff_row, key_cols0: list[int], trim: bool) -> tuple[str, ...]:
    parts = []
    for c0 in key_cols0:
        text = eff_row.get(c0, ("", None, 0, c0))[0]
        parts.append(_text_for_compare(text, trim))
    return tuple(parts)


def _key_display(eff_row, key_cols0: list[int]) -> str:
    parts = []
    for c0 in key_cols0:
        text = eff_row.get(c0, ("", None, 0, c0))[0]
        parts.append(str(text).strip())
    return " | ".join(parts)


def _key_is_empty(parts: tuple[str, ...]) -> bool:
    return all(p == "" for p in parts)


def _build_key_map(left_eff, row_numbers, key_cols0, trim, side, warnings):
    """key tuple → 行下标（重复键取第一行）；同时产出空键/重复键告警。"""
    mapping: dict[tuple[str, ...], int] = {}
    dup: dict[tuple[str, ...], list[int]] = {}
    empty_rows: list[int] = []
    for i, eff_row in enumerate(left_eff):
        parts = _key_parts(eff_row, key_cols0, trim)
        if _key_is_empty(parts):
            empty_rows.append(row_numbers[i])
            continue
        if parts in mapping:
            dup.setdefault(parts, [mapping[parts]]).append(i)
        else:
            mapping[parts] = i
    if empty_rows:
        shown = "、".join(f"第{r0 + 1}行" for r0 in empty_rows[:10])
        more = "等" if len(empty_rows) > 10 else ""
        warnings.append(
            f"{side}表主键为空的行已跳过 {len(empty_rows)} 行（{shown}{more}）"
        )
    for parts, idxs in sorted(dup.items(), key=lambda kv: kv[1][0]):
        first = idxs[0]
        display = _key_display(left_eff[first], key_cols0) or "（空）"
        shown_rows = "、".join(f"第{row_numbers[i] + 1}行" for i in idxs[:10])
        more = "等" if len(idxs) > 10 else ""
        warnings.append(
            f"{side}表主键重复：{display} 出现 {len(idxs)} 次"
            f"（{shown_rows}{more}，取第{row_numbers[first] + 1}行参与匹配）"
        )
    return mapping


# --------------------------------------------------------------------------
# 坐标与字段名
# --------------------------------------------------------------------------
def _eff_ref(table: TableModel, eff_row, i: int, c0: int, row_numbers) -> CellRef:
    item = eff_row.get(c0)
    if item is None:
        return CellRef.from_index(table.path or "", table.sheet or "", row_numbers[i], c0)
    _, _, origin_r0, origin_c0 = item
    return CellRef.from_index(table.path or "", table.sheet or "", origin_r0, origin_c0)


def _field_names(table: TableModel) -> dict[int, str]:
    out: dict[int, str] = {}
    for f in table.fields:
        if isinstance(f.index, int) and f.index >= 0:
            name = (f.name or "").strip()
            if name:
                out[f.index] = name
    return out


def _field_label(c0: int, lnames: dict[int, str], rnames: dict[int, str]) -> str:
    return lnames.get(c0) or rnames.get(c0) or f"第{c0 + 1}列（{col_letter(c0 + 1)}）"


def _diff_item(kind: str, key: str, field_name: str, left_ref, right_ref,
               left_value: str, right_value: str) -> dict:
    from .model import DiffItem

    return DiffItem(
        kind=kind,
        key=key,
        field_name=field_name,
        left_ref=left_ref,
        right_ref=right_ref,
        left_value=left_value,
        right_value=right_value,
    ).to_dict()


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def compare(left: TableModel, right: TableModel, keys, options=None) -> dict:
    """两张逻辑表按主键对比，返回 summary + items + warnings（见模块 docstring）。"""
    if not isinstance(left, TableModel) or not isinstance(right, TableModel):
        raise TypeError("compare() 需要两个 TableModel（左表、右表）")
    keys_1based = _normalize_keys(keys)
    opts = _normalize_options(options)
    trim = opts["trim"]
    tolerance = opts["tolerance"]
    filldown = opts["filldown"]
    ignore_empty = opts["ignoreEmpty"]
    key_cols0 = [k - 1 for k in keys_1based]

    warnings: list[str] = []
    row_numbers_l = _safe_row_numbers(left, "左", warnings)
    row_numbers_r = _safe_row_numbers(right, "右", warnings)
    if not left.path:
        warnings.append("左表缺少文件路径（TableModel.source 为空），差异坐标 filePath 将为空")
    if not right.path:
        warnings.append("右表缺少文件路径（TableModel.source 为空），差异坐标 filePath 将为空")

    left_cols = _table_columns(left)
    right_cols = _table_columns(right)
    all_cols = left_cols | right_cols

    for c0 in key_cols0:
        letter = col_letter(c0 + 1)
        in_left = c0 in left_cols
        in_right = c0 in right_cols
        if not in_left and not in_right:
            warnings.append(f"主键列 {c0 + 1}（{letter}）在左右两表中都不存在")
        elif not in_left:
            warnings.append(f"主键列 {c0 + 1}（{letter}）在左表中不存在")
        elif not in_right:
            warnings.append(f"主键列 {c0 + 1}（{letter}）在右表中不存在")

    lnames = _field_names(left)
    rnames = _field_names(right)
    diff_cols = sorted(all_cols - set(key_cols0))

    # 字段名/字段集差异告警（同列号进行比较，不猜测列顺序）
    renamed: list[str] = []
    for c0 in diff_cols:
        ln, rn = lnames.get(c0), rnames.get(c0)
        if ln and rn and _normalize(ln) != _normalize(rn):
            renamed.append(f"第{c0 + 1}列（{col_letter(c0 + 1)}）左“{ln}”/右“{rn}”")
    if renamed:
        shown = "；".join(renamed[:5])
        more = f" 等 {len(renamed)} 处" if len(renamed) > 5 else ""
        warnings.append(f"左右表字段名不一致（按同列号比较）：{shown}{more}")
    only_left = [(c0, lnames[c0]) for c0 in sorted(lnames) if c0 not in rnames and c0 not in key_cols0]
    only_right = [(c0, rnames[c0]) for c0 in sorted(rnames) if c0 not in lnames and c0 not in key_cols0]
    if only_left:
        names = "、".join(name for _, name in only_left[:8])
        more = f" 等 {len(only_left)} 个" if len(only_left) > 8 else ""
        warnings.append(f"右表缺少左表字段：{names}{more}（这些列右侧按空值参与比较）")
    if only_right:
        names = "、".join(name for _, name in only_right[:8])
        more = f" 等 {len(only_right)} 个" if len(only_right) > 8 else ""
        warnings.append(f"左表缺少右表字段：{names}{more}（这些列左侧按空值参与比较）")

    left_eff, lfilled = _build_effective(left, all_cols, filldown)
    right_eff, rfilled = _build_effective(right, all_cols, filldown)
    if filldown and (lfilled or rfilled):
        warnings.append(
            f"已按向下填充补全空单元格：左表 {lfilled} 个、右表 {rfilled} 个"
            "（合并单元格/分类列按上方最近非空值补全，差异坐标指向值真实所在单元格）"
        )

    left_map = _build_key_map(left_eff, row_numbers_l, key_cols0, trim, "左", warnings)
    right_map = _build_key_map(right_eff, row_numbers_r, key_cols0, trim, "右", warnings)

    removed: list[dict] = []
    added: list[dict] = []
    changed: list[dict] = []
    matched = 0
    changed_rows = 0
    changed_cells = 0

    # ---- removed / matched（按左表行序） ----
    for i, eff_row in enumerate(left_eff):
        parts = _key_parts(eff_row, key_cols0, trim)
        if _key_is_empty(parts):
            continue
        key_display = _key_display(eff_row, key_cols0)
        if parts not in right_map:
            left_ref = _eff_ref(left, eff_row, i, key_cols0[0], row_numbers_l)
            removed.append(_diff_item("removed", key_display, "", left_ref, None, "", ""))
            continue
        if left_map.get(parts) != i:  # 重复键：只有第一行参与匹配
            continue
        j = right_map[parts]
        matched += 1
        row_changed = False
        right_eff_row = right_eff[j]
        for c0 in diff_cols:
            litem = eff_row.get(c0)
            ritem = right_eff_row.get(c0)
            if litem is None and ritem is None:
                continue
            ltext, lnum = (litem[0], litem[1]) if litem else ("", None)
            rtext, rnum = (ritem[0], ritem[1]) if ritem else ("", None)
            l_empty = _is_empty(ltext)
            r_empty = _is_empty(rtext)
            if not l_empty and not r_empty:
                if lnum is not None and rnum is not None:
                    is_diff = abs(float(lnum) - float(rnum)) > tolerance
                else:
                    is_diff = _text_for_compare(ltext, trim) != _text_for_compare(rtext, trim)
            elif l_empty and r_empty:
                continue
            else:  # 一侧为空
                if ignore_empty:
                    continue
                is_diff = True
            if not is_diff:
                continue
            left_ref = _eff_ref(left, eff_row, i, c0, row_numbers_l)
            right_ref = _eff_ref(right, right_eff_row, j, c0, row_numbers_r)
            changed.append(_diff_item(
                "changed", key_display, _field_label(c0, lnames, rnames),
                left_ref, right_ref, ltext, rtext,
            ))
            changed_cells += 1
            row_changed = True
        if row_changed:
            changed_rows += 1

    # ---- added（仅有右表，按右表行序） ----
    for j, eff_row in enumerate(right_eff):
        parts = _key_parts(eff_row, key_cols0, trim)
        if _key_is_empty(parts):
            continue
        if parts in left_map:
            continue
        key_display = _key_display(eff_row, key_cols0)
        right_ref = _eff_ref(right, eff_row, j, key_cols0[0], row_numbers_r)
        added.append(_diff_item("added", key_display, "", None, right_ref, "", ""))

    # 排序：仅左（左行序）→ 仅右（右行序）→ 修改（左行序、列号）
    items = removed + added + changed

    if not left.rows:
        warnings.append("左表没有数据行，无法对比")
    if not right.rows:
        warnings.append("右表没有数据行，无法对比")

    return {
        "summary": {
            "matched": matched,
            "added": len(added),
            "removed": len(removed),
            "changed": changed_rows,
            "leftRows": len(left.rows),
            "rightRows": len(right.rows),
            "changedCells": changed_cells,
        },
        "items": items,
        "keys": keys_1based,
        "warnings": warnings,
    }
