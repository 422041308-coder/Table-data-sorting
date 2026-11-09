# -*- coding: utf-8 -*-
"""S5 双表对比单测：增/删/改、多键、重复键、空键、容差、归一化、向下填充、导出报告。

运行：
    python -m unittest tests.test_compare -v
    （或 python -m unittest discover -s tests -v）
"""
from __future__ import annotations

import os
import shutil
import unittest
import uuid
import zipfile
import xml.etree.ElementTree as ET

from app.core.compare import DEFAULT_OPTIONS, compare
from app.core.model import Cell, Field, FileRef, TableModel, col_letter
from app.core.xlsxwriter import write_xlsx
from app.server import ApiError
from app.services import compare_service

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# --------------------------------------------------------------------------
# fixture 工具
# --------------------------------------------------------------------------
def _mk_cell(value) -> Cell:
    if value is None or value == "":
        return Cell(raw=None, kind="blank", text="")
    if isinstance(value, bool):
        return Cell(raw=value, kind="bool", text="TRUE" if value else "FALSE")
    if isinstance(value, (int, float)):
        num = float(value)
        text = str(int(num)) if num.is_integer() else str(num)
        return Cell(raw=value, kind="number", text=text, num=num)
    return Cell(raw=value, kind="text", text=str(value))


def make_table(path: str, sheet: str, fields: dict[int, str], rows: list[dict[int, object]],
               row_numbers: list[int] | None = None, start_row0: int = 3) -> TableModel:
    """用 dict（0 基列号 → 值）构造 TableModel fixture。"""
    src = FileRef(id=os.path.basename(path), path=os.path.abspath(path),
                  name=os.path.basename(path), status="ok")
    fields_list = [Field(index=c0, name=name) for c0, name in sorted(fields.items())]
    data = [{c0: _mk_cell(v) for c0, v in row.items()} for row in rows]
    if row_numbers is None:
        row_numbers = list(range(start_row0, start_row0 + len(rows)))
    return TableModel(
        source=src,
        sheet=sheet,
        header_rows=[1],
        data_start=start_row0,
        data_end=row_numbers[-1] if row_numbers else -1,
        fields=fields_list,
        rows=data,
        row_numbers=list(row_numbers),
    )


def _items_of(items, kind):
    return [it for it in items if it["kind"] == kind]


class CompareCoreTest(unittest.TestCase):
    LEFT = r"C:\对比样本\左表.xlsx"
    RIGHT = r"C:\对比样本\右表.xlsx"
    SHEET = "台账"

    def assert_ref(self, ref, path, sheet, row, col):
        self.assertIsNotNone(ref, "差异必须带该侧 CellRef")
        self.assertEqual(ref["filePath"], os.path.abspath(path),
                         "CellRef 必须是文件绝对路径")
        self.assertEqual(ref["sheet"], sheet)
        self.assertEqual(ref["row"], row, "CellRef.row 应为 1 基")
        self.assertEqual(ref["col"], col, "CellRef.col 应为 1 基")
        self.assertEqual(ref["colLetter"], col_letter(col))
        self.assertEqual(ref["a1"], col_letter(col) + str(row))
        self.assertIn(f"第{row}行", ref["display"])

    # ---- 增 / 删 / 改 + 完整坐标 ----
    def test_added_removed_changed_and_refs(self):
        left = make_table(
            self.LEFT, self.SHEET,
            {0: "合同编号", 1: "项目", 2: "金额"},
            [
                {0: "A", 1: "甲", 2: 100},
                {0: "B", 1: "乙", 2: 200},
                {0: "D", 1: "丁", 2: 400},
            ],
            row_numbers=[3, 4, 5],
        )
        right = make_table(
            self.RIGHT, self.SHEET,
            {0: "合同编号", 1: "项目", 2: "金额"},
            [
                {0: "A", 1: "甲", 2: 100},
                {0: "B", 1: "乙", 2: 250},
                {0: "C", 1: "丙", 2: 300},
            ],
            row_numbers=[3, 4, 5],
        )
        res = compare(left, right, [1], {"filldown": False})
        self.assertEqual(
            {k: res["summary"][k] for k in
             ("matched", "added", "removed", "changed", "leftRows", "rightRows", "changedCells")},
            {"matched": 2, "added": 1, "removed": 1, "changed": 1,
             "leftRows": 3, "rightRows": 3, "changedCells": 1},
        )
        self.assertEqual(res["keys"], [1])
        # added=仅右有，removed=仅左有，changed=逐字段
        self.assertEqual([it["kind"] for it in res["items"]],
                         ["removed", "added", "changed"])

        removed = _items_of(res["items"], "removed")[0]
        self.assertEqual(removed["key"], "D")
        self.assert_ref(removed["leftRef"], self.LEFT, self.SHEET, 6, 1)
        self.assertIsNone(removed["rightRef"])

        added = _items_of(res["items"], "added")[0]
        self.assertEqual(added["key"], "C")
        self.assertIsNone(added["leftRef"])
        self.assert_ref(added["rightRef"], self.RIGHT, self.SHEET, 6, 1)  # 右表第 3 行 = 第 6 行

        changed = _items_of(res["items"], "changed")[0]
        self.assertEqual(changed["key"], "B")
        self.assertEqual(changed["fieldName"], "金额")
        self.assertEqual(changed["leftValue"], "200")
        self.assertEqual(changed["rightValue"], "250")
        self.assert_ref(changed["leftRef"], self.LEFT, self.SHEET, 5, 3)  # 第 4 个工作表行 = 第 5 行
        self.assert_ref(changed["rightRef"], self.RIGHT, self.SHEET, 5, 3)
        # 每条差异的 ref 字段完整（含 A1/display，供前端直接展示）
        for it in res["items"]:
            for side in ("leftRef", "rightRef"):
                if it[side] is None:
                    continue
                for key in ("filePath", "sheet", "row", "col", "colLetter", "a1", "display"):
                    self.assertIn(key, it[side])

    # ---- 多列组合键 ----
    def test_multi_column_keys(self):
        left = make_table(
            self.LEFT, self.SHEET, {0: "编号", 1: "版次", 2: "值"},
            [{0: "K", 1: "A", 2: 1}, {0: "K", 1: "B", 2: 2}, {0: "M", 1: "A", 2: 3}],
            row_numbers=[3, 4, 5],
        )
        right = make_table(
            self.RIGHT, self.SHEET, {0: "编号", 1: "版次", 2: "值"},
            [{0: "K", 1: "A", 2: 1}, {0: "K", 1: "B", 2: 9}],
            row_numbers=[3, 4],
        )
        res = compare(left, right, [1, 2], {"filldown": False})
        self.assertEqual(res["summary"]["matched"], 2)
        self.assertEqual(res["summary"]["removed"], 1)
        self.assertEqual(res["summary"]["added"], 0)
        self.assertEqual(res["summary"]["changed"], 1)
        changed = _items_of(res["items"], "changed")[0]
        self.assertEqual(changed["key"], "K | B")
        removed = _items_of(res["items"], "removed")[0]
        self.assertEqual(removed["key"], "M | A")

    # ---- 重复主键：告警且不静默错配 ----
    def test_duplicate_keys_warning_no_mispair(self):
        left = make_table(
            self.LEFT, self.SHEET, {0: "合同编号", 1: "金额"},
            [{0: "A", 1: 1}, {0: "A", 1: 2}, {0: "B", 1: 3}],
            row_numbers=[4, 5, 6],
        )
        right = make_table(
            self.RIGHT, self.SHEET, {0: "合同编号", 1: "金额"},
            [{0: "A", 1: 1}, {0: "B", 1: 9}],
            row_numbers=[4, 5],
        )
        res = compare(left, right, [1], {"filldown": False})
        self.assertEqual(res["summary"]["matched"], 2)
        self.assertEqual(res["summary"]["changed"], 1)
        text = "\n".join(res["warnings"])
        self.assertIn("主键重复：A 出现 2 次", text)
        self.assertIn("参与匹配", text)
        # 重复的第二行不会与唯一行错配：差异只有 B 的金额
        self.assertEqual([it["fieldName"] for it in _items_of(res["items"], "changed")],
                         ["金额"])

    def test_empty_key_warning_and_skip(self):
        left = make_table(
            self.LEFT, self.SHEET, {0: "合同编号", 1: "金额"},
            [{0: "", 1: 1}, {0: "   ", 1: 2}, {0: "A", 1: 3}],
            row_numbers=[4, 5, 6],
        )
        right = make_table(
            self.RIGHT, self.SHEET, {0: "合同编号", 1: "金额"},
            [{0: "A", 1: 3}],
            row_numbers=[4],
        )
        res = compare(left, right, [1], {"filldown": False})
        self.assertEqual(res["summary"]["matched"], 1)
        self.assertEqual(res["summary"]["removed"], 0, "空主键行不能算作仅左有")
        self.assertEqual(res["summary"]["added"], 0)
        self.assertIn("主键为空的行已跳过 2 行", "\n".join(res["warnings"]))

    # ---- 数值容差 ----
    def test_number_tolerance(self):
        left = make_table(self.LEFT, self.SHEET, {0: "ID", 1: "金额"},
                          [{0: "A", 1: 100.0}], row_numbers=[3])
        right = make_table(self.RIGHT, self.SHEET, {0: "ID", 1: "金额"},
                           [{0: "A", 1: 100.0005}], row_numbers=[3])
        exact = compare(left, right, [1])
        self.assertEqual(exact["summary"]["changedCells"], 1)
        loose = compare(left, right, [1], {"tolerance": 0.001})
        self.assertEqual(loose["summary"]["changedCells"], 0)
        tighter = compare(left, right, [1], {"tolerance": "0.0001"})
        self.assertEqual(tighter["summary"]["changedCells"], 1)

    # ---- 文本归一化（全角/空白/大小写） ----
    def test_text_normalize_and_trim_option(self):
        left = make_table(self.LEFT, self.SHEET, {0: "ID", 1: "名称"},
                          [{0: "A", 1: " ＡＢＣ "}, {0: "B", 1: "甲  乙"}], row_numbers=[3, 4])
        right = make_table(self.RIGHT, self.SHEET, {0: "ID", 1: "名称"},
                           [{0: "A", 1: "abc"}, {0: "B", 1: "甲 乙"}], row_numbers=[3, 4])
        norm = compare(left, right, [1])
        self.assertEqual(norm["summary"]["changedCells"], 0,
                         "normalize_text 应抹平全角/大小写/多余空白")
        strict = compare(left, right, [1], {"trim": False})
        self.assertEqual(strict["summary"]["changedCells"], 2)

    # ---- 向下填充（合并单元格/分类列只在首行） ----
    def test_filldown_merged_style_column(self):
        # 序号唯一（可作主键），“项目/金额”只在首行有值（视觉合并）
        left = make_table(
            self.LEFT, self.SHEET, {0: "序号", 1: "项目", 2: "金额"},
            [{0: 1, 1: "P1", 2: 100}, {0: 2, 1: None, 2: None}],
            row_numbers=[3, 4],
        )
        right = make_table(
            self.RIGHT, self.SHEET, {0: "序号", 1: "项目", 2: "金额"},
            [{0: 1, 1: "P1", 2: 100}, {0: 2, 1: "P1", 2: 100}],
            row_numbers=[3, 4],
        )
        filled = compare(left, right, [1])  # 默认 filldown=True，主键用唯一序号列
        self.assertEqual(filled["summary"]["matched"], 2)
        self.assertEqual(filled["summary"]["changedCells"], 0,
                         "向下填充后第 4 行的项目/金额应与右表一致")
        self.assertTrue(any("向下填充" in w for w in filled["warnings"]))
        # 被填充的值应指向真实单元格（合并区左上角），不是空单元格
        changed_right = make_table(
            self.RIGHT, self.SHEET, {0: "序号", 1: "项目", 2: "金额"},
            [{0: 1, 1: "P1", 2: 100}, {0: 2, 1: "P1", 2: 200}],
            row_numbers=[3, 4],
        )
        res = compare(left, changed_right, [1])
        changed = _items_of(res["items"], "changed")[0]
        self.assertEqual(changed["fieldName"], "金额")
        self.assertEqual(changed["leftValue"], "100")
        self.assertEqual(changed["leftRef"]["row"], 4)  # 值真实所在行（第 1 条记录）
        self.assertEqual(changed["leftRef"]["col"], 3)
        # 关闭 filldown：空值不再补全，第 4 行的项目/金额为空 vs 有值 → 2 处修改
        off = compare(left, right, [1], {"filldown": False})
        self.assertEqual(off["summary"]["matched"], 2)
        self.assertEqual(off["summary"]["changedCells"], 2)
        # 若把“只在首行有值”的项目列当主键：填充后必然出现重复键，必须告警而不是错配
        key_filled = compare(left, right, [2])  # 1 基第 2 列 = 项目
        self.assertIn("主键重复：P1 出现 2 次", "\n".join(key_filled["warnings"]))
        self.assertEqual(key_filled["summary"]["matched"], 1)

    # ---- ignoreEmpty ----
    def test_ignore_empty(self):
        left = make_table(self.LEFT, self.SHEET, {0: "ID", 1: "备注"},
                          [{0: "A", 1: ""}], row_numbers=[3])
        right = make_table(self.RIGHT, self.SHEET, {0: "ID", 1: "备注"},
                           [{0: "A", 1: "有值"}], row_numbers=[3])
        default = compare(left, right, [1])
        self.assertEqual(default["summary"]["changedCells"], 1)
        ignored = compare(left, right, [1], {"ignoreEmpty": True})
        self.assertEqual(ignored["summary"]["changedCells"], 0)

    # ---- 字段名/缺列告警 ----
    def test_field_name_mismatch_warning(self):
        left = make_table(self.LEFT, self.SHEET, {0: "编码", 1: "金额"},
                          [{0: "A", 1: 1}], row_numbers=[3])
        right = make_table(self.RIGHT, self.SHEET, {0: "编码", 1: "金额（元）"},
                           [{0: "A", 1: 1}], row_numbers=[3])
        res = compare(left, right, [1])
        self.assertIn("字段名不一致", "\n".join(res["warnings"]))
        self.assertEqual(res["summary"]["changedCells"], 0)

    def test_missing_key_column_warning(self):
        left = make_table(self.LEFT, self.SHEET, {0: "ID"}, [{0: "A"}], row_numbers=[3])
        right = make_table(self.RIGHT, self.SHEET, {0: "ID"}, [{0: "A"}], row_numbers=[3])
        res = compare(left, right, [5])
        self.assertIn("主键列 5（E）在左右两表中都不存在", "\n".join(res["warnings"]))
        self.assertEqual(res["summary"]["matched"], 0)

    def test_invalid_params(self):
        left = make_table(self.LEFT, self.SHEET, {0: "ID"}, [{0: "A"}])
        right = make_table(self.RIGHT, self.SHEET, {0: "ID"}, [{0: "A"}])
        with self.assertRaises(ValueError):
            compare(left, right, [])
        with self.assertRaises(ValueError):
            compare(left, right, [0])
        with self.assertRaises(ValueError):
            compare(left, right, [1], {"tolerance": -0.1})
        with self.assertRaises(ValueError):
            compare(left, right, [1], {"tolerance": "abc"})
        self.assertTrue(DEFAULT_OPTIONS["filldown"], "文档约定向下填充默认开启")


class CompareServiceTest(unittest.TestCase):
    """服务层端到端：xlsx 写入 → 加载 → 对比 → 导出报告 → 重新解析。"""

    def setUp(self):
        base = os.path.join(_PROJECT, "data", "tmp")
        os.makedirs(base, exist_ok=True)
        # 注意：不用 tempfile.mkdtemp（它按 0700 建目录，在本项目沙箱下会拒绝写入），
        # 改用普通目录 + uuid 保证唯一。
        self.tmp = os.path.join(base, "test_compare_" + uuid.uuid4().hex)
        os.makedirs(self.tmp, exist_ok=True)
        self.left_path = os.path.join(self.tmp, "左表.xlsx")
        self.right_path = os.path.join(self.tmp, "右表.xlsx")
        write_xlsx(self.left_path, [{
            "name": "台账",
            "header": ["合同编号", "项目", "金额"],
            "rows": [["A", "甲", 100], ["B", "乙", 200], ["D", "丁", 400]],
        }])
        write_xlsx(self.right_path, [{
            "name": "台账",
            "header": ["合同编号", "项目", "金额"],
            "rows": [["A", "甲", 100], ["B", "乙", 250], ["C", "丙", 300]],
        }])
        self.body = {
            "left": {"path": self.left_path, "sheet": "台账"},
            "right": {"path": self.right_path, "sheet": "台账"},
            "keys": [1],
            "options": {"trim": True, "tolerance": 0.001, "filldown": False},
        }

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_service_compare_end_to_end(self):
        res = compare_service.compare(self.body)
        self.assertEqual(res["summary"]["matched"], 2)
        self.assertEqual(res["summary"]["added"], 1)
        self.assertEqual(res["summary"]["removed"], 1)
        self.assertEqual(res["summary"]["changed"], 1)
        self.assertEqual(os.path.abspath(res["left"]["path"]), os.path.abspath(self.left_path))
        for it in res["items"]:
            for side in ("leftRef", "rightRef"):
                if it[side] is None:
                    continue
                self.assertEqual(it[side]["filePath"], os.path.abspath(self.left_path)
                                 if side == "leftRef" else os.path.abspath(self.right_path))
                for key in ("filePath", "sheet", "row", "col", "colLetter", "a1", "display"):
                    self.assertIn(key, it[side])
        # 修改的是金额（第 3 列）
        changed = _items_of(res["items"], "changed")[0]
        self.assertEqual((changed["leftRef"]["row"], changed["leftRef"]["col"]), (3, 3))
        self.assertEqual((changed["leftRef"]["a1"], changed["rightRef"]["a1"]), ("C3", "C3"))

    def test_service_export_report_reparsable(self):
        out = compare_service.export(self.body, out_dir=self.tmp)
        self.assertTrue(os.path.isfile(out["path"]))
        self.assertEqual(out["count"], 3)
        self.assertTrue(out["styled"], "应成功写入黄/红/绿样式")
        with zipfile.ZipFile(out["path"]) as z:
            names = set(z.namelist())
            self.assertIn("xl/styles.xml", names)
            styles = z.read("xl/styles.xml").decode("utf-8")
            self.assertIn('<fills count="5">', styles)
            self.assertIn('<cellXfs count="5">', styles)
            for color in ("FFFFF2CC", "FFFFC7CE", "FFC6EFCE"):
                self.assertIn(color, styles)
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            sheet_names = [s.get("name") for s in wb.find(_NS + "sheets")]
            self.assertEqual(sheet_names, ["差异明细", "汇总"])
            sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
            rows = sheet.find(_NS + "sheetData")
            header_cells = list(rows[0])
            header_text = []
            for c in header_cells:
                header_text.append("".join(t.text or "" for t in c.iter(_NS + "t")))
            for col in compare_service._REPORT_HEADER:
                self.assertIn(col, header_text, f"报告必须包含列：{col}")
            data_rows = list(rows)[1:]
            self.assertEqual(len(data_rows), 3)
            row2_cells = list(data_rows[0])
            self.assertTrue(all(c.get("s") for c in row2_cells),
                            "差异明细数据行应带样式索引（修改黄/仅左红/仅右绿）")
            # 第一行是“仅左（删除）”的 D 行
            first_text = "".join("".join(t.text or "" for t in c.iter(_NS + "t"))
                                 for c in row2_cells)
            self.assertIn("仅左", first_text)
            self.assertIn("D", first_text)
            # 汇总表包含汇总与告警信息
            summary = ET.fromstring(z.read("xl/worksheets/sheet2.xml"))
            summary_text = "".join(t.text or "" for t in summary.iter(_NS + "t"))
            self.assertIn("匹配行数", summary_text)
            self.assertIn("有修改的行数", summary_text)

    def test_service_rejects_bad_body(self):
        cases = [
            ({"right": {"path": self.right_path}, "keys": [1]}, "BAD_PARAM"),
            ({"left": {"path": self.left_path}, "right": {"path": self.right_path},
              "keys": []}, "BAD_PARAM"),
            ({"left": {"path": os.path.join(self.tmp, "无.xlsx")},
              "right": {"path": self.right_path}, "keys": [1]}, "NOT_FOUND"),
            ({"left": {"path": self.left_path, "sheet": "不存在"},
              "right": {"path": self.right_path}, "keys": [1]}, "BAD_PARAM"),
            ({"left": {"path": self.left_path}, "right": {"path": self.right_path},
              "keys": [1], "options": {"tolerance": -1}}, "BAD_PARAM"),
        ]
        for body, code in cases:
            with self.assertRaises(ApiError) as ctx:
                compare_service.compare(body)
            self.assertEqual(ctx.exception.code, code, ctx.exception.message)


if __name__ == "__main__":
    unittest.main()
