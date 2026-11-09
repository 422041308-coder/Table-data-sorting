# -*- coding: utf-8 -*-
"""Dump the structure of the sample workbook to a UTF-8 text file."""
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

path, out_path = sys.argv[1], sys.argv[2]
wb = load_workbook(path, data_only=False)
lines = [f"FILE: {path}", f"SHEETS: {wb.sheetnames}"]

for ws in wb.worksheets:
    lines.append("=" * 100)
    lines.append(f"SHEET: {ws.title!r}  max_row={ws.max_row} max_col={ws.max_column}")
    lines.append(f"MERGED: {[str(r) for r in ws.merged_cells.ranges]}")
    lines.append(f"FREEZE: {ws.freeze_panes}  FILTER: {ws.auto_filter.ref}")
    lines.append("-- cell-by-cell, first 10 rows --")
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            v = cell.value
            if v is None:
                continue
            lines.append(f"{get_column_letter(c)}{r} [{type(v).__name__}] = {v!r}")
    # 2nd/3rd row header labels paired with column letter
    lines.append("-- column -> header labels row2/row3 --")
    for c in range(1, ws.max_column + 1):
        v2 = ws.cell(row=2, column=c).value
        v3 = ws.cell(row=3, column=c).value
        lines.append(f"{get_column_letter(c)}: r2={v2!r} r3={v3!r}")
    lines.append("-- data row 4 raw --")
    for c in range(1, ws.max_column + 1):
        lines.append(f"{get_column_letter(c)}4 = {ws.cell(row=4, column=c).value!r}")

with open(out_path, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))
print(f"written {out_path} ({len(lines)} lines)")
