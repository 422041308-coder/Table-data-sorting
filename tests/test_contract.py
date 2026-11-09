# -*- coding: utf-8 -*-
"""契约冒烟测试：数据模型 / xlsx 写入 / 路由 / mock 形状。

运行：python -m unittest tests.test_contract -v
"""
import os
import sys
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app.core import xlsxwriter  # noqa: E402
from app.core.model import (  # noqa: E402
    CONTRACT_VERSION,
    Cell,
    CellRef,
    col_index,
    col_letter,
    split_a1,
)


class TestCoords(unittest.TestCase):
    def test_col_letter(self):
        self.assertEqual(col_letter(1), "A")
        self.assertEqual(col_letter(26), "Z")
        self.assertEqual(col_letter(27), "AA")
        self.assertEqual(col_letter(32), "AF")
        self.assertEqual(col_letter(58), "BF")

    def test_col_index(self):
        self.assertEqual(col_index("A"), 1)
        self.assertEqual(col_index("af"), 32)
        self.assertEqual(col_index("BF"), 58)

    def test_cellref(self):
        ref = CellRef(r"C:\a\b.xlsx", "Sheet0", 12, 32)
        self.assertEqual(ref.col_letter, "AF")
        self.assertEqual(ref.a1, "AF12")
        self.assertIn("第12行", ref.display)
        d = ref.to_dict()
        for k in ("filePath", "sheet", "row", "col", "colLetter", "a1", "display"):
            self.assertIn(k, d)
        self.assertEqual(d["row"], 12)

    def test_from_index(self):
        ref = CellRef.from_index("x.xlsx", "S", 3, 31)
        self.assertEqual((ref.row, ref.col), (4, 32))
        self.assertEqual(ref.a1, "AF4")

    def test_split_a1(self):
        self.assertEqual(split_a1("AF12"), (12, 32))
        self.assertEqual(split_a1("B3"), (3, 2))

    def test_contract_version(self):
        self.assertTrue(CONTRACT_VERSION.startswith("1."))


class TestXlsxWriter(unittest.TestCase):
    def test_roundtrip(self):
        td = os.path.join(ROOT, "data", "tmp")  # 工作区内可写；不依赖系统 Temp
        os.makedirs(td, exist_ok=True)
        path = os.path.join(td, "out.xlsx")
        xlsxwriter.write_xlsx(path, [
            {"name": "检索结果", "header": ["文件地址", "第几行", "A1"],
             "rows": [[r"C:\a\b.xlsx", 4, "AF4"], ["x", 5, "A5"]],
             "widths": [40, 8, 10], "merges": [(1, 1, 1, 3)], "freeze": "A2"},
            {"name": "第二表", "rows": [["中文", 1.5, True]]},
        ])
        self.assertTrue(os.path.isfile(path))
        self.assertTrue(zipfile.is_zipfile(path))
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            self.assertIn("xl/workbook.xml", names)
            self.assertIn("xl/worksheets/sheet1.xml", names)
            self.assertIn("xl/worksheets/sheet2.xml", names)
            wb = z.read("xl/workbook.xml").decode("utf-8")
            self.assertIn("检索结果", wb)
            sheet1 = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("文件地址", sheet1)
            self.assertIn("AF4", sheet1)
            self.assertIn("mergeCell", sheet1)
            # XML 必须可解析
            ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
            ET.fromstring(z.read("xl/styles.xml"))
            ET.fromstring(z.read("[Content_Types].xml"))

    def test_autowidth(self):
        widths = xlsxwriter.auto_widths([["中文内容", 1]], header=["列", "数字"])
        self.assertEqual(len(widths), 2)
        self.assertGreater(widths[0], widths[1])


class TestRoutes(unittest.TestCase):
    def test_app_and_routes(self):
        from app.server import create_app

        app = create_app(ROOT)
        self.assertGreater(len(app.routes), 20)
        fn, match = app.match("GET", "/api/health")
        self.assertIsNotNone(fn)
        self.assertIsNotNone(match)
        fn2, _ = app.match("GET", "/api/no-such")
        self.assertIsNone(fn2)

    def test_no_duplicate_routes(self):
        from app.server import create_app

        app = create_app(ROOT)
        seen = set()
        for method, rx, _fn in app.routes:
            key = (method, rx.pattern)
            self.assertNotIn(key, seen, f"重复路由: {key}")
            seen.add(key)


class TestMockShape(unittest.TestCase):
    def test_search_shape(self):
        from app import mock

        res = mock.search("合同")
        self.assertIn("total", res)
        self.assertIn("hits", res)
        if res["hits"]:
            hit = res["hits"][0]
            for k in ("ref", "fieldName", "valueText", "snippet", "highlight", "score"):
                self.assertIn(k, hit)
            for k in ("filePath", "sheet", "row", "col", "colLetter", "a1"):
                self.assertIn(k, hit["ref"])

    def test_compare_shape(self):
        from app import mock

        res = mock.compare({})
        self.assertIn("summary", res)
        self.assertIn("items", res)
        for item in res["items"]:
            self.assertIn(item["kind"], ("added", "removed", "changed"))
            self.assertIn("leftRef", item)
            self.assertIn("rightRef", item)

    def test_preview_shape(self):
        from app import mock

        res = mock.preview("f1", 0, 5)
        self.assertIn("rows", res)
        self.assertIn("merges", res)
        self.assertIn("headerRows", res)


class TestCell(unittest.TestCase):
    def test_cell_dict(self):
        c = Cell(raw=123, kind="number", text="123", num=123.0)
        d = c.to_dict()
        self.assertEqual(d["kind"], "number")
        self.assertEqual(d["num"], 123.0)


if __name__ == "__main__":
    unittest.main()
