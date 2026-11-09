# -*- coding: utf-8 -*-
"""极简 xlsx 写入器（零依赖，Lead 维护，契约冻结）。

用途：检索结果导出、对比报告导出、映射转换的"新建表格"场景。
不做样式保真——模板改造走 ``app/core/template.py`` 的克隆注入路径。

对外接口（冻结）：
    write_xlsx(path, sheets=[...], default_width=14.0) -> str
    sheet 字典字段：name, rows, header, widths, merges, freeze
"""
from __future__ import annotations

import datetime as _dt
import math
import os
import re
import zipfile
from xml.sax.saxutils import escape

_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xml_text(value) -> str:
    s = str(value)
    s = _BAD_XML.sub("", s)
    return escape(s)


def _q(s: str) -> str:
    """XML 属性值转义。"""
    return escape(s, {'"': "&quot;"})


def _cell_xml(ref: str, value, style: int = 0) -> str:
    """生成一个 <c> 节点。支持的 Python 类型：None/str/int/float/bool/datetime。"""
    if value is None:
        return ""
    if isinstance(value, bool):
        return f'<c r="{ref}" s="{style}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (_dt.datetime, _dt.date)):
        # 日期以文本写入（导出场景足够；模板场景走 template.py）
        return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{_xml_text(value.isoformat(sep=" "))}</t></is></c>'
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{value}</t></is></c>'
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    text = str(value)
    if len(text) > 32000:  # Excel 单元格上限 32767
        text = text[:32000]
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{_xml_text(text)}</t></is></c>'


def _sheet_xml(sheet: dict, default_width: float) -> str:
    from .model import col_letter

    name = sheet.get("name") or "Sheet1"
    rows: list[list] = list(sheet.get("rows") or [])
    header = sheet.get("header")
    widths = sheet.get("widths")
    merges = sheet.get("merges") or []
    freeze = sheet.get("freeze")

    all_rows = ([header] if header else []) + rows
    body = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            style = 1 if (header and r_idx == 1) else 0
            xml = _cell_xml(f"{col_letter(c_idx)}{r_idx}", value, style)
            if xml:
                cells.append(xml)
        if cells:
            body.append(f'<row r="{r_idx}">' + "".join(cells) + "</row>")

    cols_xml = ""
    if widths:
        parts = []
        if isinstance(widths, dict):
            items = sorted(widths.items())
        else:
            items = list(enumerate(widths, start=1))
        for idx, w in items:
            try:
                parts.append(f'<col min="{int(idx)}" max="{int(idx)}" width="{float(w)}" customWidth="1"/>')
            except (TypeError, ValueError):
                continue
        if parts:
            cols_xml = "<cols>" + "".join(parts) + "</cols>"

    pane = ""
    view = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    if freeze and re.fullmatch(r"[A-Z]+\d+", str(freeze)):
        from .model import split_a1

        fr, fc = split_a1(freeze)
        x = col_letter(fc)
        if fr > 1 and fc > 1:
            pane = (f'<pane xSplit="{fc - 1}" ySplit="{fr - 1}" topLeftCell="{x}{fr}" '
                    f'activePane="bottomRight" state="frozen"/>')
        elif fr > 1:
            pane = f'<pane ySplit="{fr - 1}" topLeftCell="A{fr}" activePane="bottomLeft" state="frozen"/>'
        elif fc > 1:
            pane = f'<pane xSplit="{fc - 1}" topLeftCell="{x}1" activePane="topRight" state="frozen"/>'
        view = f'<sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews>'

    merge_xml = ""
    if merges:
        refs = []
        for m in merges:
            r1, r2, c1, c2 = [int(x) for x in m]
            refs.append(f"{col_letter(c1)}{r1}:{col_letter(c2)}{r2}")
        merge_xml = f'<mergeCells count="{len(refs)}">' + "".join(
            f'<mergeCell ref="{r}"/>' for r in refs) + "</mergeCells>"

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + view
        + cols_xml
        + "<sheetData>" + "".join(body) + "</sheetData>"
        + merge_xml
        + "</worksheet>"
    )


_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="2">'
    '<font><sz val="11"/><name val="宋体"/><charset val="134"/></font>'
    '<font><b/><sz val="11"/><name val="宋体"/><charset val="134"/></font>'
    '</fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="2">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)


def write_xlsx(path: str, sheets: list[dict], default_width: float = 14.0) -> str:
    """写出一个 xlsx 文件，返回路径。"""
    from .model import col_letter  # noqa: F401

    if not sheets:
        sheets = [{"name": "Sheet1", "rows": []}]
    total = len(sheets)
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)

    ct = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for i in range(1, total + 1):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append("</Types>")

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    wb = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        "<sheets>",
    ]
    for i, sheet in enumerate(sheets, start=1):
        nm = _q(str(sheet.get("name") or f"Sheet{i}")[:31])
        wb.append(f'<sheet name="{nm}" sheetId="{i}" r:id="rId{i}"/>')
    wb.append("</sheets></workbook>")

    wbrels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for i in range(1, total + 1):
        wbrels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    wbrels.append(f'<Relationship Id="rId{total + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    wbrels.append("</Relationships>")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wbrels))
        z.writestr("xl/styles.xml", _STYLES)
        for i, sheet in enumerate(sheets, start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(sheet, default_width))
    return path


def write_rows_xlsx(path: str, rows: list[list], header: list | None = None,
                    sheet_name: str = "Sheet1", widths: list | dict | None = None) -> str:
    """单表快捷写法。"""
    return write_xlsx(path, [{
        "name": sheet_name,
        "rows": rows,
        "header": header,
        "widths": widths,
    }])


def auto_widths(rows: list[list], header: list | None = None, min_w: float = 8.0,
                max_w: float = 60.0) -> list[float]:
    """按内容估算列宽（中文字符按 2 计）。"""
    all_rows = ([header] if header else []) + rows
    n = max((len(r) for r in all_rows), default=0)
    out = []
    for c in range(n):
        best = min_w
        for r in all_rows:
            if c >= len(r):
                continue
            s = "" if r[c] is None else str(r[c])
            w = sum(2 if ord(ch) > 127 else 1 for ch in s) + 2
            best = max(best, min(w, max_w))
        out.append(best)
    return out
