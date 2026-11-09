# -*- coding: utf-8 -*-
"""S3 检索单元测试：索引 / 查询语义 / 高亮 / 增量更新。"""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app.core.index import SearchIndex  # noqa: E402
from app.core.model import Cell, CellRef, Field, FileRef, TableModel  # noqa: E402
from app.core.search import highlight, parse_query  # noqa: E402


def cell(text, kind="text"):
    return Cell(raw=text, kind=kind, text=str(text))


def make_table(path=r"C:\data\a.xlsx", sheet="台账"):
    fields = [
        Field(0, "序号"), Field(1, "项目名称"), Field(2, "合同名称"),
        Field(3, "合同编码"), Field(4, "金额"),
    ]
    rows = [
        {0: cell("1"), 1: cell("示例双创改造工程"), 2: cell("设计合同"),
         3: cell("DEMO-ZS-24001"), 4: cell("52800")},
        {0: cell("2"), 1: cell("乙南侧环境整治工程"), 2: cell("施工合同"),
         3: cell("DEMO-DX-24002"), 4: cell("2788335")},
        {0: cell("3"), 1: cell("丙土方清运"), 2: cell("监理合同"),
         3: cell("DEMO-YG-24003"), 4: cell("27883")},
    ]
    return TableModel(
        source=FileRef(id="x", path=path, name="a.xlsx"),
        sheet=sheet,
        header_rows=[1, 2],
        data_start=3,
        fields=fields,
        rows=rows,
        row_numbers=[3, 4, 5],
    )


class TestParseQuery(unittest.TestCase):
    def test_terms(self):
        p = parse_query("示例 合同")
        self.assertEqual(p["terms"], ["示例", "合同"])
        self.assertEqual(p["phrases"], [])

    def test_quotes(self):
        p = parse_query('"设计合同" 示例')
        self.assertEqual(p["phrases"], ["设计合同"])
        self.assertEqual(p["terms"], ["示例"])

    def test_chinese_quotes(self):
        p = parse_query("“施工合同”")
        self.assertEqual(p["phrases"], ["施工合同"])


class TestHighlight(unittest.TestCase):
    def test_basic(self):
        snippet, marks = highlight("示例双创改造工程设计合同", "设计合同")
        self.assertIn("设计合同", snippet)
        self.assertEqual(snippet[marks[0]:marks[1]], "设计合同")

    def test_not_found(self):
        snippet, marks = highlight("abc", "xyz")
        self.assertEqual(marks, [])
        self.assertEqual(snippet, "abc")

    def test_truncate(self):
        text = "前" * 100 + "目标" + "后" * 100
        snippet, marks = highlight(text, "目标", context=5)
        self.assertTrue(snippet.startswith("…"))
        self.assertTrue(snippet.endswith("…"))
        self.assertEqual(snippet[marks[0]:marks[1]], "目标")


class TestIndex(unittest.TestCase):
    def setUp(self):
        self.index = SearchIndex()
        self.index.add(make_table())

    def test_size(self):
        size = self.index.size()
        self.assertEqual(size["tables"], 1)
        self.assertEqual(size["rows"], 3)
        self.assertEqual(size["cells"], 15)

    def test_single_keyword(self):
        hits, total = self.index.search("24001")
        self.assertEqual(total, 1)
        ref = hits[0].ref
        self.assertEqual(ref.row, 4)       # row_numbers[0]=3 → 1 基 4
        self.assertEqual(ref.col, 4)       # 0 基 3 → 1 基 4
        self.assertEqual(ref.col_letter, "D")
        self.assertEqual(ref.a1, "D4")
        self.assertEqual(ref.sheet, "台账")
        self.assertEqual(ref.file_path, r"C:\data\a.xlsx")
        self.assertEqual(hits[0].field_name, "合同编码")
        self.assertIn("24001", hits[0].snippet)

    def test_multi_term_all_same_row(self):
        hits, total = self.index.search("示例 合同")
        self.assertGreaterEqual(total, 2)  # 项目名称格 + 合同名称格
        rows = {(h.ref.row) for h in hits}
        self.assertEqual(rows, {4})

    def test_all_no_match_across_rows(self):
        hits, total = self.index.search("示例 乙")
        self.assertEqual(total, 0)

    def test_any_mode(self):
        hits, total = self.index.search("示例 乙", mode="any")
        rows = {h.ref.row for h in hits}
        self.assertEqual(rows, {4, 5})  # 第 4 行(示例)、第 5 行(乙)

    def test_phrase(self):
        hits, total = self.index.search('"设计合同"')
        self.assertEqual(total, 1)
        self.assertEqual(hits[0].ref.row, 4)

    def test_field_filter(self):
        hits, total = self.index.search("24001", field="合同编码")
        self.assertEqual(total, 1)
        self.assertEqual(hits[0].field_name, "合同编码")
        # 该字段里没有的词不应命中
        hits, total = self.index.search("24001", field="项目名称")
        self.assertEqual(total, 0)

    def test_value_text_and_highlight(self):
        hits, _ = self.index.search("DEMO-DX")
        self.assertEqual(len(hits), 1)
        snippet, marks = hits[0].snippet, hits[0].highlight
        self.assertEqual(snippet[marks[0]:marks[1]], "DEMO-DX")

    def test_pagination(self):
        hits1, total = self.index.search("合同", limit=2, offset=0)
        hits2, _ = self.index.search("合同", limit=2, offset=2)
        self.assertEqual(total, len(hits1) + len(hits2))
        self.assertNotEqual([h.ref.a1 for h in hits1], [h.ref.a1 for h in hits2])

    def test_update_cell(self):
        ref = CellRef(r"C:\data\a.xlsx", "台账", 4, 4)
        self.assertTrue(self.index.update_cell(ref, "ZZZ-NEW"))
        hits, total = self.index.search("24001")
        self.assertEqual(total, 0)
        hits, total = self.index.search("ZZZ-NEW")
        self.assertEqual(total, 1)
        self.assertEqual(hits[0].ref.a1, "D4")

    def test_unknown_cell_update(self):
        ref = CellRef(r"C:\data\none.xlsx", "台账", 1, 1)
        self.assertFalse(self.index.update_cell(ref, "x"))

    def test_clear(self):
        self.index.clear()
        hits, total = self.index.search("24001")
        self.assertEqual(total, 0)
        self.assertEqual(self.index.size()["rows"], 0)


class TestIndexNumbers(unittest.TestCase):
    def test_numeric_cell(self):
        t = make_table()
        t.rows[0][4] = Cell(raw=52800, kind="number", text="52800", num=52800.0)
        idx = SearchIndex()
        idx.add(t)
        hits, total = idx.search("52800")
        self.assertGreaterEqual(total, 1)
        self.assertEqual(hits[0].ref.a1, "E4")


if __name__ == "__main__":
    unittest.main()
