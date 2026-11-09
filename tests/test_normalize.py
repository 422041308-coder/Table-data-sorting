# -*- coding: utf-8 -*-
"""S2 归一化工具单测：normalize_text / cell_display / guess_kind / to_number / to_date_text。"""
import datetime as dt
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app.core.model import (  # noqa: E402
    KIND_BLANK,
    KIND_BOOL,
    KIND_DATE,
    KIND_FORMULA,
    KIND_NUMBER,
    KIND_TEXT,
    Cell,
)
from app.core.normalize import (  # noqa: E402
    cell_display,
    format_number,
    guess_kind,
    is_blank,
    normalize_text,
    to_date_text,
    to_number,
)


class TestNormalizeText(unittest.TestCase):
    def test_fullwidth_to_halfwidth(self):
        self.assertEqual(normalize_text("ＡＢＣ１２３"), "abc123")
        self.assertEqual(normalize_text("（元）"), "(元)")

    def test_whitespace_collapsed(self):
        self.assertEqual(normalize_text("  合同\n\n名称\t\t"), "合同 名称")
        self.assertEqual(normalize_text("序号\u3000\u3000名称"), "序号 名称")

    def test_case_lowered(self):
        self.assertEqual(normalize_text("DEMO-ZS-24001"), "demo-zs-24001")

    def test_quotes_unified(self):
        self.assertEqual(normalize_text("“合同”"), '"合同"')
        self.assertEqual(normalize_text("‘A’"), "'a'")

    def test_empty(self):
        self.assertEqual(normalize_text(None), "")
        self.assertEqual(normalize_text("   "), "")

    def test_idempotent(self):
        once = normalize_text("  Ａ  B\t（元） ")
        self.assertEqual(normalize_text(once), once)


class TestFormatAndDisplay(unittest.TestCase):
    def test_format_number(self):
        self.assertEqual(format_number(52800.0), "52800")
        self.assertEqual(format_number(0.06), "0.06")
        self.assertEqual(format_number(2245976.1), "2245976.1")

    def test_number_cell_drops_dot_zero(self):
        self.assertEqual(cell_display(Cell(kind=KIND_NUMBER, num=52800.0, text="52800")), "52800")
        self.assertEqual(cell_display(Cell(kind=KIND_NUMBER, num=0.06, text="0.06")), "0.06")

    def test_number_cell_without_num(self):
        self.assertEqual(cell_display(Cell(kind=KIND_NUMBER, text="1,234.50")), "1234.5")

    def test_formula_uses_cached_value(self):
        cell = Cell(kind=KIND_FORMULA, formula="=SUM(A1:A3)", num=2245976.1, text="")
        self.assertEqual(cell_display(cell), "2245976.1")
        # 拿不到缓存值才退回公式文本
        cell2 = Cell(kind=KIND_FORMULA, formula="=SUM(A1:A3)", text="=SUM(A1:A3)")
        self.assertEqual(cell_display(cell2), "=SUM(A1:A3)")

    def test_date_cell(self):
        self.assertEqual(cell_display(Cell(kind=KIND_DATE, text="2024.3.9")), "2024.3.9")
        self.assertEqual(cell_display(Cell(kind=KIND_DATE, num=45360.0)), "2024.3.9")

    def test_bool_cell(self):
        self.assertEqual(cell_display(Cell(kind=KIND_BOOL, raw=True)), "TRUE")
        self.assertEqual(cell_display(Cell(kind=KIND_BOOL, raw=False)), "FALSE")

    def test_text_passthrough(self):
        self.assertEqual(cell_display(Cell(kind=KIND_TEXT, text="2024年台账")), "2024年台账")

    def test_plain_values(self):
        self.assertEqual(cell_display(None), "")
        self.assertEqual(cell_display(3), "3")
        self.assertEqual(cell_display(3.0), "3")
        self.assertEqual(cell_display("abc"), "abc")

    def test_is_blank(self):
        self.assertTrue(is_blank(None))
        self.assertTrue(is_blank(Cell(kind=KIND_BLANK)))
        self.assertTrue(is_blank("   "))
        self.assertFalse(is_blank(Cell(kind=KIND_NUMBER, num=0.0)))
        self.assertFalse(is_blank("0"))


class TestToNumber(unittest.TestCase):
    def test_plain(self):
        self.assertEqual(to_number("123"), 123.0)
        self.assertEqual(to_number(123), 123.0)
        self.assertIsNone(to_number(True))

    def test_amounts(self):
        self.assertEqual(to_number("1,234.50"), 1234.5)
        self.assertEqual(to_number("￥52800"), 52800.0)
        self.assertEqual(to_number("52800元"), 52800.0)
        self.assertEqual(to_number(" 2 788 335 "), 2788335.0)

    def test_percent_and_negative(self):
        self.assertAlmostEqual(to_number("6%"), 0.06)
        self.assertEqual(to_number("(123)"), -123.0)
        self.assertEqual(to_number("-4.5"), -4.5)

    def test_fullwidth(self):
        self.assertEqual(to_number("１２３"), 123.0)

    def test_not_numbers(self):
        self.assertIsNone(to_number("abc"))
        self.assertIsNone(to_number("2024.3.9"))
        self.assertIsNone(to_number(""))
        self.assertIsNone(to_number(None))
        self.assertIsNone(to_number("-"))
        self.assertIsNone(to_number("DEMO-ZS-24001"))

    def test_cell(self):
        self.assertEqual(to_number(Cell(kind=KIND_NUMBER, num=10.0)), 10.0)
        self.assertEqual(to_number(Cell(kind=KIND_TEXT, text="7")), 7.0)


class TestToDateText(unittest.TestCase):
    def test_text_forms(self):
        for raw in ("2024.3.9", "2024-03-09", "2024/3/9", "2024年3月9日", "2024.03.09"):
            self.assertEqual(to_date_text(raw), "2024.3.9", raw)

    def test_datetime_forms(self):
        self.assertEqual(to_date_text("2024-03-09 12:00:00"), "2024.3.9")
        self.assertEqual(to_date_text("2026-11-08 13:01:00.0"), "2026.11.8")

    def test_compact(self):
        self.assertEqual(to_date_text("20240309"), "2024.3.9")
        self.assertEqual(to_date_text(20240309), "2024.3.9")

    def test_objects(self):
        self.assertEqual(to_date_text(dt.date(2024, 3, 9)), "2024.3.9")
        self.assertEqual(to_date_text(dt.datetime(2024, 3, 9, 12, 30)), "2024.3.9")

    def test_excel_serial(self):
        self.assertEqual(to_date_text(45360), "2024.3.9")
        self.assertEqual(to_date_text(45360.75), "2024.3.9")

    def test_invalid(self):
        for raw in ("2024.13.40", "abc", "", None, "13800000001", True):
            self.assertIsNone(to_date_text(raw), raw)


class TestGuessKind(unittest.TestCase):
    def test_number(self):
        self.assertEqual(guess_kind([1, 2, 3]), "number")
        self.assertEqual(guess_kind(["52800", "2788335"]), "number")

    def test_phone_from_text_and_number(self):
        self.assertEqual(guess_kind(["13800000001", "13800000002"]), "phone")
        self.assertEqual(guess_kind([13800000003, 13800000004]), "phone")

    def test_date(self):
        self.assertEqual(guess_kind(["2024.3.9", "2024-03-10"]), "date")

    def test_code(self):
        self.assertEqual(guess_kind(["DEMO-ZS-24001", "DEMO-JS-24002"]), "code")

    def test_money(self):
        self.assertEqual(guess_kind(["￥1,000.00", "￥2,000.00"]), "money")

    def test_text(self):
        self.assertEqual(guess_kind(["合同A", "合同B"]), "text")

    def test_empty_and_mixed(self):
        self.assertEqual(guess_kind([]), "text")
        self.assertEqual(guess_kind([None, "", "  "]), "text")
        self.assertEqual(guess_kind(["合同A", None, "合同B"]), "text")

    def test_cell_values(self):
        cells = [
            Cell(kind=KIND_NUMBER, num=1.0, text="1"),
            Cell(kind=KIND_NUMBER, num=2.0, text="2"),
        ]
        self.assertEqual(guess_kind(cells), "number")


if __name__ == "__main__":
    unittest.main()
