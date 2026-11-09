# -*- coding: utf-8 -*-
"""双表对比服务层（切片 S5）。

对外接口（契约冻结，见 docs/CONTRACT.md 第 4 节 S5）：

* ``compare(body, root=None) -> dict``
    body = ``{"left": {"path", "sheet"}, "right": {"path", "sheet"},
              "keys": [32], "options": {"trim": true, "tolerance": 0.001, "filldown": true}}``
    keys 是 **1 基列号**；sheet 可为工作表名或 0 基下标，缺省取第一个有数据的表。
* ``export(body, root=None, out_dir=None) -> dict``
    导出 xlsx 差异报告，返回 ``{"path", "count", ...}``。

加载两表走 S1（``app.core.readers.read_workbook``）+ S2（``app.core.detect.detect_table``），
对比引擎（``app/core/compare.py``）只接收真实 ``TableModel``；读取/识别失败统一抛
``ApiError``（中文原因 + 稳定错误码），不会静默返回空表。

坐标：差异项里的左右 CellRef 由核心引擎生成（文件绝对路径 + 工作表 + 1 基行 + 1 基列 + A1）；
本服务层负责把相对路径按 ``root`` 解析成绝对路径。

导出报告列（顺序固定）：
    差异类型｜主键｜字段名｜左值｜左文件｜左工作表｜左行｜左列｜左A1｜
    右值｜右文件｜右工作表｜右行｜右列｜右A1
样式：修改=黄底、仅左（removed）=红底、仅右（added）=绿底。
``app/core/xlsxwriter.py`` 对外只有表头加粗样式，因此本模块在写完后对 xlsx 包做一次
最小后处理（追加 3 个填充样式并给明细行单元格设置 s 索引）；后处理失败时保留数据正确的
无样式报告，并在返回值的 ``styleNote`` 里说明。

本模块只依赖标准库 + ``app/core/model.py`` / ``app/core/compare.py`` / ``app/core/xlsxwriter.py``
/ S1 读取层 / S2 识别层 / ``app/server.ApiError``。
"""
from __future__ import annotations

import datetime as _dt
import os
import re
import zipfile
import xml.etree.ElementTree as ET

from ..core.compare import DEFAULT_OPTIONS, compare as compare_tables
from ..core.model import (
    FMT_UNKNOWN,
    FileRef,
    SheetModel,
    TableModel,
    WorkbookModel,
)
from ..core.xlsxwriter import auto_widths, write_xlsx
from ..server import ApiError

__all__ = ["compare", "export"]

_REPORT_HEADER = [
    "差异类型", "主键", "字段名", "左值", "左文件", "左工作表", "左行", "左列", "左A1",
    "右值", "右文件", "右工作表", "右行", "右列", "右A1",
]
_KIND_LABEL = {"changed": "修改", "removed": "仅左（删除）", "added": "仅右（新增）"}
_KIND_COLOR = {"changed": "FFFFF2CC", "removed": "FFFFC7CE", "added": "FFC6EFCE"}


# --------------------------------------------------------------------------
# 参数解析
# --------------------------------------------------------------------------
def _project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _resolve_path(path: str, root) -> str:
    path = os.path.expanduser(str(path).strip())
    if not os.path.isabs(path) and root:
        path = os.path.join(str(root), path)
    return os.path.abspath(path)


def _fail(message: str, code: str = "BAD_PARAM", status: int = 400):
    """抛出带稳定错误码的 ApiError（api.py 统一转成 {"ok":false,...} 响应）。"""
    raise ApiError(code, message, status)


def _parse_spec(spec, side: str, root) -> dict:
    if isinstance(spec, str):
        path = spec
        sheet = None
    elif isinstance(spec, dict):
        path = spec.get("path") or spec.get("filePath") or spec.get("file")
        sheet = spec.get("sheet")
    else:
        _fail(f"{side}表参数必须是对象 {{\"path\",\"sheet\"}} 或文件路径字符串")
    if not path:
        _fail(f"{side}表缺少 path（文件地址）")
    path = _resolve_path(path, root)
    if not os.path.isfile(path):
        _fail(f"{side}表文件不存在：{path}", "NOT_FOUND", 404)
    return {"path": path, "sheet": sheet}


def _parse_body(body, root):
    if not isinstance(body, dict):
        _fail("对比参数必须是对象")
    left_spec = _parse_spec(body.get("left"), "左", root)
    right_spec = _parse_spec(body.get("right"), "右", root)
    keys = body.get("keys")
    if keys is None:
        _fail("缺少 keys：请指定 1 基主键列号，例如 [32] 或 [1,2]")
    options = body.get("options")
    if options is None:
        options = {}
    if not isinstance(options, dict):
        _fail("options 必须是对象")
    merged = dict(DEFAULT_OPTIONS)
    merged.update({k: v for k, v in options.items() if k in merged})
    return left_spec, right_spec, keys, merged


# --------------------------------------------------------------------------
# 加载层：S1 read_workbook + S2 detect_table
# --------------------------------------------------------------------------
def _pick_sheet(wb: WorkbookModel, sheet_spec, side: str) -> SheetModel:
    sheets = [s for s in wb.sheets]
    if not sheets:
        _fail(f"{side}表没有工作表：{wb.path}", "READ_FAILED")
    if sheet_spec is None or sheet_spec == "":
        for s in sheets:
            if s.cells:
                return s
        return sheets[0]
    if isinstance(sheet_spec, int) and not isinstance(sheet_spec, bool):
        if sheet_spec < 0 or sheet_spec >= len(sheets):
            _fail(f"{side}表工作表下标越界：{sheet_spec}（共 {len(sheets)} 个）")
        return sheets[sheet_spec]
    name = str(sheet_spec)
    for s in sheets:
        if s.name == name:
            return s
    options = "、".join(s.name for s in sheets[:10])
    _fail(f"{side}表工作表不存在：{name}（可选：{options}）")


def _file_ref(path: str, wb: WorkbookModel) -> FileRef:
    try:
        st = os.stat(path)
        size, mtime = st.st_size, st.st_mtime
    except OSError:
        size, mtime = 0, 0.0
    return FileRef(
        id=os.path.basename(path),
        path=path,
        name=os.path.basename(path),
        fmt=wb.fmt or FMT_UNKNOWN,
        size=size,
        mtime=mtime,
        sheets_count=len(wb.sheets),
        status="ok",
    )


def _load_table(spec: dict, side: str) -> TableModel:
    path = spec["path"]
    sheet_spec = spec["sheet"]
    try:
        from app.core.detect import detect_table
        from app.core.readers import read_workbook
    except Exception as exc:  # noqa: BLE001 - 依赖未交付（正常集成时不会发生）
        _fail(f"读取/结构识别模块未就绪：{exc}", "DEPENDENCY_MISSING", 500)
    try:
        wb = read_workbook(path)
    except Exception as exc:  # noqa: BLE001 - 统一转中文错误
        _fail(f"{side}表读取失败：{path}（{exc}）", "READ_FAILED")
    sheet = _pick_sheet(wb, sheet_spec, side)
    file_ref = _file_ref(path, wb)
    try:
        try:
            table = detect_table(wb, sheet, file_ref=file_ref)
        except TypeError:  # 兼容只接受两个位置参数的实现
            table = detect_table(wb, sheet)
    except Exception as exc:  # noqa: BLE001
        _fail(f"{side}表结构识别失败：{path} / {sheet.name}（{exc}）", "DETECT_FAILED")
    if table is None or not table.rows:
        _fail(f"{side}表没有识别到数据行：{path} / {sheet.name}", "EMPTY_TABLE")
    if table.source is None:
        table.source = file_ref
    if table.source and table.source.path:
        table.source.path = os.path.abspath(table.source.path)
    if not table.sheet:
        table.sheet = sheet.name
    return table




# --------------------------------------------------------------------------
# 对比
# --------------------------------------------------------------------------
def compare(body: dict, root=None) -> dict:
    """加载左右表并对比，返回核心引擎结果 + 左右表元信息。"""
    left_spec, right_spec, keys, options = _parse_body(body, root)
    left = _load_table(left_spec, "左")
    right = _load_table(right_spec, "右")
    try:
        result = compare_tables(left, right, keys, options)
    except ValueError as exc:  # 主键/选项参数错误 → 400
        _fail(str(exc), "BAD_PARAM")
    result["warnings"] = list(result.get("warnings") or [])
    result["left"] = {"path": left.path, "sheet": left.sheet, "rows": len(left.rows)}
    result["right"] = {"path": right.path, "sheet": right.sheet, "rows": len(right.rows)}
    result["options"] = options
    return result


# --------------------------------------------------------------------------
# 导出 xlsx 报告
# --------------------------------------------------------------------------
def _report_rows(items: list[dict]) -> list[list]:
    rows = []
    for item in items:
        left = item.get("leftRef") or {}
        right = item.get("rightRef") or {}

        def pick(ref, key):
            return ref.get(key, "") if ref else ""

        rows.append([
            _KIND_LABEL.get(item.get("kind"), item.get("kind") or ""),
            item.get("key") or "",
            item.get("fieldName") or "",
            item.get("leftValue") or "",
            pick(left, "filePath"),
            pick(left, "sheet"),
            left.get("row") if left else "",
            left.get("col") if left else "",
            pick(left, "a1"),
            item.get("rightValue") or "",
            pick(right, "filePath"),
            pick(right, "sheet"),
            right.get("row") if right else "",
            right.get("col") if right else "",
            pick(right, "a1"),
        ])
    return rows


def _summary_rows(result: dict) -> list[list]:
    s = result.get("summary") or {}
    left = result.get("left") or {}
    right = result.get("right") or {}
    keys = result.get("keys") or []
    rows = [
        ["左表文件", left.get("path", "")],
        ["左表工作表", left.get("sheet", "")],
        ["右表文件", right.get("path", "")],
        ["右表工作表", right.get("sheet", "")],
        ["主键（1 基列号）", "、".join(str(k) for k in keys)],
        ["匹配行数", s.get("matched", 0)],
        ["仅右有（新增）行数", s.get("added", 0)],
        ["仅左有（删除）行数", s.get("removed", 0)],
        ["有修改的行数", s.get("changed", 0)],
        ["修改单元格数", s.get("changedCells", 0)],
        ["左表数据行数", s.get("leftRows", 0)],
        ["右表数据行数", s.get("rightRows", 0)],
        ["样式说明", "修改=黄底、仅左（删除）=红底、仅右（新增）=绿底"],
        ["生成时间", _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ]
    for idx, msg in enumerate(result.get("warnings") or [], start=1):
        rows.append([f"告警 {idx}", msg])
    return rows


def _unique_path(out_dir: str) -> str:
    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = f"对比报告_{ts}"
    path = os.path.join(out_dir, base + ".xlsx")
    n = 2
    while os.path.exists(path):
        path = os.path.join(out_dir, f"{base}_{n}.xlsx")
        n += 1
    return path


def _resolve_out_dir(root, out_dir) -> str:
    if out_dir:
        out = os.path.dirname(os.path.abspath(out_dir)) if out_dir.lower().endswith(".xlsx") else out_dir
        out = _resolve_path(out, root)
    else:
        out = os.path.join(_resolve_path(str(root), None) if root else _project_root(), "data", "output")
    os.makedirs(out, exist_ok=True)
    return out


def _patch_styles(z: zipfile.ZipFile, sheet_xml: str, kinds: list[str]) -> tuple[str, str]:
    """在 styles.xml 追加 3 个填充样式，并给明细行加 s 索引。返回 (styles_xml, sheet_xml)。"""
    styles = z.read("xl/styles.xml").decode("utf-8")

    def bump(tag: str, add: int) -> int:
        nonlocal styles
        m = re.search(rf'<{tag} count="(\d+)">', styles)
        if not m:
            raise ValueError(f"styles.xml 缺少 <{tag}>")
        base = int(m.group(1))
        styles = styles[:m.start()] + f'<{tag} count="{base + add}">' + styles[m.end():]
        return base

    fill_base = bump("fills", len(_KIND_COLOR))
    xf_base = bump("cellXfs", len(_KIND_COLOR))
    fills_xml = "".join(
        f'<fill><patternFill patternType="solid"><fgColor rgb="{color}"/>'
        f'<bgColor indexed="64"/></patternFill></fill>'
        for color in _KIND_COLOR.values()
    )
    styles = styles.replace("</fills>", fills_xml + "</fills>", 1)
    xf_xml = "".join(
        f'<xf numFmtId="0" fontId="0" fillId="{fill_base + i}" borderId="0" xfId="0" applyFill="1"/>'
        for i in range(len(_KIND_COLOR))
    )
    styles = styles.replace("</cellXfs>", xf_xml + "</cellXfs>", 1)
    style_index = {kind: xf_base + i for i, kind in enumerate(_KIND_COLOR)}

    def repl(match):
        row_num = int(match.group(1))
        body = match.group(2)
        idx = row_num - 2  # 数据行从第 2 行开始（第 1 行是表头）
        if 0 <= idx < len(kinds):
            s = style_index.get(kinds[idx])
            if s is not None:
                # xlsxwriter 会给每个 <c> 写 s="0"（表头为 s="1"），先替换 s="0"；
                # 若某实现省略了 s 属性，再为缺 s 的单元格补上，避免重复属性。
                body = re.sub(r'\bs="0"', f's="{s}"', body)
                body = re.sub(r'<c (?![^>]*\bs=")', f'<c s="{s}" ', body)
        return f'<row r="{row_num}">' + body + "</row>"

    sheet_xml = re.sub(r'<row r="(\d+)">(.*?)</row>', repl, sheet_xml, flags=re.S)
    return styles, sheet_xml


def _rewrite_zip(path: str, replacements: dict[str, str]) -> None:
    """就地重写 xlsx 包（先整体读入内存，避免沙箱下 os.replace 覆盖已存在文件被拒）。"""
    with zipfile.ZipFile(path) as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w") as zout:
        for info, data in entries:
            if info.filename in replacements:
                data = replacements[info.filename].encode("utf-8")
            zout.writestr(info, data)


def _parse_xlsx_or_raise(path: str) -> None:
    with zipfile.ZipFile(path) as z:
        for name in ("[Content_Types].xml", "xl/workbook.xml", "xl/styles.xml",
                     "xl/worksheets/sheet1.xml"):
            if name not in z.namelist():
                raise ValueError(f"报告缺少包内条目：{name}")
            ET.fromstring(z.read(name))
        styles = z.read("xl/styles.xml").decode("utf-8")
        if 'count="5"' not in styles and '<fills count="5">' not in styles:
            raise ValueError("报告样式表未包含 5 个填充")


def _apply_kind_styles(path: str, kinds: list[str]) -> bool:
    if not kinds:
        return False
    try:
        with zipfile.ZipFile(path) as z:
            if "xl/styles.xml" not in z.namelist():
                return False
            sheet_xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            styles, sheet_xml = _patch_styles(z, sheet_xml, kinds)
        _rewrite_zip(path, {"xl/styles.xml": styles, "xl/worksheets/sheet1.xml": sheet_xml})
        _parse_xlsx_or_raise(path)
        return True
    except Exception:  # noqa: BLE001 - 样式失败不影响数据正确性
        return False


def export(body: dict, root=None, out_dir=None) -> dict:
    """执行对比并导出 xlsx 差异报告。返回 {"path","count","styled","summary","warnings"}。"""
    result = compare(body, root)
    items = result.get("items") or []
    rows = _report_rows(items)
    header = list(_REPORT_HEADER)
    widths = auto_widths(rows, header, min_w=8.0, max_w=60.0)
    out_dir_abs = _resolve_out_dir(root, out_dir)
    path = _unique_path(out_dir_abs)
    sheets = [
        {
            "name": "差异明细",
            "header": header,
            "rows": rows,
            "widths": {i + 1: w for i, w in enumerate(widths)},
            "freeze": "A2",
        },
        {
            "name": "汇总",
            "header": ["项目", "值"],
            "rows": _summary_rows(result),
            "widths": {1: 18.0, 2: 80.0},
            "freeze": "A2",
        },
    ]
    write_xlsx(path, sheets)
    styled = _apply_kind_styles(path, [str(it.get("kind") or "") for it in items])
    warnings = list(result.get("warnings") or [])
    style_note = ""
    if items and not styled:
        # 样式后处理失败时重写一份无样式但数据完整的报告（避免半损坏的 zip）
        try:
            write_xlsx(path, sheets)
        except Exception:  # noqa: BLE001
            pass
        style_note = "样式标注未应用（xlsxwriter 现有能力不足 / 后处理失败），数据列完整可直接使用"
        warnings.append(style_note)
    elif styled:
        style_note = "已应用样式：修改=黄底、仅左=红底、仅右=绿底"
    return {
        "path": path,
        "count": len(items),
        "styled": styled,
        "styleNote": style_note,
        "summary": result.get("summary") or {},
        "warnings": warnings,
    }
