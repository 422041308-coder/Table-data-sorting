# -*- coding: utf-8 -*-
"""S1 采集层单测：格式嗅探 / xlsx / xls / csv 读取。

运行：``python -m unittest tests.test_readers -v``

* 项目代码零第三方依赖；测试里允许用 DSH 自带的 openpyxl 做 **对拍**（仅对 xlsx），
  没有 openpyxl 时相关用例自动 skip；
* 真实语料用例在语料目录不存在时自动 skip。
"""
import os
import re
import shutil
import struct
import sys
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app.core.model import (  # noqa: E402
    KIND_BOOL,
    KIND_DATE,
    KIND_FORMULA,
    KIND_NUMBER,
    KIND_TEXT,
    Cell,
    CellRef,
)
from app.core.readers import read_biff, read_csv, read_ooxml, read_workbook  # noqa: E402
from app.core.readers import biff as biff_mod  # noqa: E402
from app.core.readers import ooxml as ooxml_mod  # noqa: E402
from app.core.sniff import decode_text_bytes, detect_format  # noqa: E402

SAMPLE = os.path.join(ROOT, "samples", "项目台账统计表.xlsx")
CORPUS = os.environ.get("EXCEL_TOOLKIT_CORPUS", "")
TMP_DIR = os.path.join(ROOT, "data", "tmp", "s1_tests")

try:
    import openpyxl  # noqa: F401

    HAS_OPENPYXL = True
except ImportError:  # pragma: no cover - 目标机没有 openpyxl 时走这里
    HAS_OPENPYXL = False

CJK_RE = re.compile(r"[\u4e00-\u9fff]")


# --------------------------------------------------------------------------
# 测试辅助
# --------------------------------------------------------------------------
def _corpus_files(exts):
    out = []
    if not os.path.isdir(CORPUS):
        return out
    for base, _dirs, files in os.walk(CORPUS):
        for name in files:
            if name.lower().endswith(tuple(exts)):
                out.append(os.path.join(base, name))
    return sorted(out)


def _corpus_find(*names):
    for path in _corpus_files((".xls", ".et", ".xlsx", ".csv")):
        if os.path.basename(path) in names:
            return path
    return None


def _prepare_tmp():
    if os.path.isdir(TMP_DIR):
        shutil.rmtree(TMP_DIR, ignore_errors=True)
    os.makedirs(TMP_DIR, exist_ok=True)
    return TMP_DIR


def _write_zip(path, entries):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in entries.items():
            if isinstance(data, str):
                data = data.encode("utf-8")
            zf.writestr(name, data)
    return path


_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_WORKBOOK_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<workbook xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}"><bookViews><workbookView/></bookViews>'
    "<sheets>"
    '<sheet name="数据" sheetId="1" r:id="rId1"/>'
    '<sheet name="隐藏" sheetId="2" state="hidden" r:id="rId2"/>'
    "</sheets></workbook>"
)
_SHEET1_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<worksheet xmlns="{_NS_MAIN}"><dimension ref="A1:N3"/><sheetData>'
    '<row r="1">'
    '<c r="A1" t="s"><v>0</v></c>'
    '<c r="B1" t="s"><v>1</v></c>'
    '<c r="C1" t="inlineStr"><is><t>内联</t></is></c>'
    '<c r="D1"><v>3.14</v></c>'
    '<c r="E1"><v>42</v></c>'
    '<c r="F1" t="b"><v>1</v></c>'
    '<c r="G1" t="e"><v>#DIV/0!</v></c>'
    '<c r="H1" s="1"><v>45292</v></c>'
    '<c r="I1"><f>1+2</f><v>3</v></c>'
    '<c r="J1"><f>SUM(A1:A3)</f></c>'
    '<c r="K1"><f t="shared" ref="K1:K3" si="0">A1+1</f><v>2</v></c>'
    '<c r="L1" s="2"><v>45292</v></c>'
    "</row>"
    '<row r="2"><c r="K2"><f t="shared" si="0"/><v>3</v></c></row>'
    '<row r="3"><c r="A3"><v>7</v></c>'
    '<c r="K3"><f t="shared" si="0"/></c></row>'
    "</sheetData>"
    '<mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells>'
    "</worksheet>"
)
_SHEET2_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<worksheet xmlns="{_NS_MAIN}"><dimension ref="A1"/><sheetData/></worksheet>'
)
_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<styleSheet xmlns="{_NS_MAIN}">'
    '<numFmts count="1">'
    '<numFmt numFmtId="164" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/>'
    "</numFmts>"
    '<cellXfs count="3">'
    '<xf numFmtId="0"/>'
    '<xf numFmtId="14" applyNumberFormat="1"/>'
    '<xf numFmtId="164" applyNumberFormat="1"/>'
    "</cellXfs></styleSheet>"
)
_SST_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<sst xmlns="{_NS_MAIN}" count="2" uniqueCount="2">'
    "<si><t>共享</t></si>"
    "<si><r><rPr><b/></rPr><t>富</t></r><r><t>文本</t></r>"
    '<rPh sb="0" eb="1"><t>ふりがな</t></rPh></si>'
    "</sst>"
)


def _write_feature_xlsx(path, macro=False, ext=".xlsx"):
    """构造一个覆盖共享/内联字符串、日期、公式、合并、隐藏表的 xlsx。"""
    main_type = ("application/vnd.ms-excel.sheet.macroEnabled.main+xml"
                 if macro else
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{_CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{main_type}"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_REL_NS}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/worksheet" Target="worksheets/sheet2.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>"
    )
    _write_zip(path, {
        "[Content_Types].xml": content_types,
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{_REL_NS}">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": _WORKBOOK_XML,
        "xl/_rels/workbook.xml.rels": rels,
        "xl/worksheets/sheet1.xml": _SHEET1_XML,
        "xl/worksheets/sheet2.xml": _SHEET2_XML,
        "xl/styles.xml": _STYLES_XML,
        "xl/sharedStrings.xml": _SST_XML,
    })
    return path if ext == ".xlsx" else _rename_ext(path, ext)


def _rename_ext(path, ext):
    if os.path.splitext(path)[1].lower() == ext:
        return path
    new_path = os.path.splitext(path)[0] + ext
    os.replace(path, new_path)
    return new_path


def _write_minimal_xlsx(path, ext=".xlsx", macro=False):
    """最小可读 xlsx（用于 xlsm / et-zip 嗅探与分发测试）。"""
    return _write_feature_xlsx(path, macro=macro, ext=ext)


class _TmpMixin(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = _prepare_tmp()

    def tmp_path(self, name):
        return os.path.join(self.tmp, name)


# --------------------------------------------------------------------------
# 1. 格式嗅探
# --------------------------------------------------------------------------
class TestSniff(_TmpMixin):
    def test_sample(self):
        self.assertEqual(detect_format(SAMPLE), "xlsx")

    def test_missing_file(self):
        with self.assertRaises(ValueError) as ctx:
            detect_format(os.path.join(self.tmp, "不存在.xlsx"))
        self.assertRegex(str(ctx.exception), CJK_RE)

    def test_unknown_binary(self):
        path = self.tmp_path("随机.bin")
        with open(path, "wb") as fp:
            fp.write(b"\x89PNG\r\n\x1a\n" + b"\x00\x01\x02" * 100)
        self.assertEqual(detect_format(path), "unknown")

    def test_zip_without_workbook_is_unknown(self):
        path = self.tmp_path("假docx.xlsx")
        _write_zip(path, {"word/document.xml": "<doc/>"})
        self.assertEqual(detect_format(path), "unknown")

    def test_xlsm_zip(self):
        path = _write_minimal_xlsx(self.tmp_path("宏表.xlsx"), macro=True)
        self.assertEqual(detect_format(path), "xlsm")

    def test_xlsb_zip(self):
        path = self.tmp_path("二进制.xlsb")
        _write_zip(path, {
            "[Content_Types].xml": f'<Types xmlns="{_CT}"/>',
            "xl/workbook.bin": b"\x00\x01\x02\x03",
        })
        self.assertEqual(detect_format(path), "xlsb")

    def test_et_zip(self):
        path = _write_minimal_xlsx(self.tmp_path("WPS表格.et"), ext=".et")
        self.assertEqual(detect_format(path), "et-zip")

    def test_ole_is_xls(self):
        # 只有魔数、内容无效：嗅探层只看容器类型
        path = self.tmp_path("老表格.xls")
        with open(path, "wb") as fp:
            fp.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504)
        self.assertEqual(detect_format(path), "xls")

    def test_ole_et(self):
        path = self.tmp_path("老WPS.et")
        with open(path, "wb") as fp:
            fp.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504)
        self.assertEqual(detect_format(path), "et-ole")

    @unittest.skipUnless(os.path.isdir(CORPUS), "真实语料目录不存在")
    def test_real_corpus_formats(self):
        """真实语料 65 个表格：格式识别必须全部正确。"""
        expected = {".xls": "xls", ".xlsx": "xlsx", ".csv": "csv", ".et": "et-ole"}
        files = _corpus_files((".xls", ".xlsx", ".csv", ".et"))
        self.assertGreaterEqual(len(files), 60)
        for path in files:
            ext = os.path.splitext(path)[1].lower()
            self.assertEqual(detect_format(path), expected[ext],
                             f"格式识别错误：{path}")

    def test_decode_text_bytes(self):
        self.assertEqual(decode_text_bytes("中文".encode("utf-8"))[0], "utf-8")
        self.assertEqual(decode_text_bytes(b"\xef\xbb\xbfA,B")[0], "utf-8-sig")
        gbk = decode_text_bytes("姓名,金额\n张三,100\n".encode("gbk"))
        self.assertEqual(gbk[0], "gbk")
        self.assertIn("张三", gbk[1])
        utf16 = decode_text_bytes("甲\t乙\n1\t2\n".encode("utf-16"))
        self.assertIn(utf16[0], ("utf-16", "utf-16-le"))
        self.assertIn("甲", utf16[1])
        self.assertIsNone(decode_text_bytes(b"\xff\xff\xff\xff"))


# --------------------------------------------------------------------------
# 2. 样本 xlsx（2 表 / 58 列 / 22 处合并）
# --------------------------------------------------------------------------
class TestSampleXlsx(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.wb = read_workbook(SAMPLE)
        cls.sheet = cls.wb.sheets[0]

    def test_format_and_sheets(self):
        self.assertEqual(self.wb.fmt, "xlsx")
        self.assertEqual(len(self.wb.sheets), 2)
        self.assertEqual([s.name for s in self.wb.sheets],
                         ["2024年项目台账（自筹-甲代-施工-日常经营）", "Sheet1"])

    def test_used_range(self):
        # 样本 A1:BF63，即 63 行 58 列
        self.assertEqual(self.sheet.n_rows, 63)
        self.assertEqual(self.sheet.n_cols, 58)

    def test_merges(self):
        # 结构导出 samples/structure.txt 与 sheet1.xml 均为 24 处合并
        # （任务书里写的 22 处应为笔误，这里按真实数据断言）
        expected = {
            (0, 0, 0, 57),   # A1:BF1
            (1, 1, 7, 22),   # H2:W2
            (1, 1, 23, 27),  # X2:AB2
            (1, 1, 28, 29),  # AC2:AD2
            (1, 1, 30, 37),  # AE2:AL2
            (1, 1, 38, 40),  # AM2:AO2
            (1, 1, 41, 44),  # AP2:AS2
            (1, 1, 45, 46),  # AT2:AU2
            (1, 1, 49, 50),  # AX2:AY2
            (1, 1, 51, 54),  # AZ2:BC2
            (1, 2, 0, 0),    # A2:A3
            (3, 6, 0, 0),    # A4:A7
            (1, 2, 1, 1),    # B2:B3
            (3, 6, 1, 1),    # B4:B7
            (1, 2, 2, 2),    # C2:C3
            (3, 6, 2, 2),    # C4:C7
            (1, 2, 3, 3),    # D2:D3
            (1, 2, 4, 4),    # E2:E3
            (1, 2, 5, 5),    # F2:F3
            (1, 2, 6, 6),    # G2:G3
            (3, 6, 37, 37),  # AL4:AL7
            (1, 2, 55, 55),  # BD2:BD3
            (1, 2, 56, 56),  # BE2:BE3
            (1, 2, 57, 57),  # BF2:BF3
        }
        self.assertEqual(len(self.sheet.merges), 24)
        self.assertEqual(set(self.sheet.merges), expected)
        for r1, r2, c1, c2 in self.sheet.merges:
            self.assertLessEqual(r1, r2)
            self.assertLessEqual(c1, c2)
            self.assertLess(r2, self.sheet.n_rows)
            self.assertLess(c2, self.sheet.n_cols)

    def test_cells_are_zero_based(self):
        key = next(iter(self.sheet.cells))
        self.assertIsInstance(key, tuple)
        self.assertGreaterEqual(key[0], 0)
        self.assertGreaterEqual(key[1], 0)
        # (0,0)=A1, (3,1)=B4, (3,28)=AC4
        self.assertIn((0, 0), self.sheet.cells)
        self.assertIn((3, 1), self.sheet.cells)
        self.assertIn((3, 28), self.sheet.cells)

    def test_text_cells(self):
        self.assertEqual(self.sheet.cells[(0, 0)].text,
                         "2024年示例区示例公司项目台账（单位：元）")
        self.assertEqual(self.sheet.cells[(1, 0)].text, "序号")
        self.assertEqual(self.sheet.cells[(3, 1)].text, "示例项目甲—园区改造工程")
        self.assertIn("DEMO-ZS-24001", self.sheet.cells[(3, 31)].text)

    def test_number_cells(self):
        ac4 = self.sheet.cells[(3, 28)]
        self.assertEqual(ac4.kind, KIND_NUMBER)
        self.assertEqual(ac4.num, 52800.0)
        self.assertEqual(ac4.text, "52800")
        ad4 = self.sheet.cells[(3, 29)]
        self.assertAlmostEqual(ad4.num, 0.06)
        self.assertEqual(ad4.text, "0.06")

    def test_formula_cached_value(self):
        # AN8 = SUM(AN4:INDEX(AN:AN,ROW()-1))，缓存值 2495529
        an8 = self.sheet.cells[(7, 39)]
        self.assertEqual(an8.formula, "=SUM(AN4:INDEX(AN:AN,ROW()-1))")
        self.assertEqual(an8.num, 2495529.0)
        self.assertEqual(an8.text, "2495529")
        # AV4 是数组公式，缓存值 0
        av4 = self.sheet.cells[(3, 47)]
        self.assertTrue(av4.formula.startswith("=_xlfn.IFS("))
        self.assertEqual(av4.num, 0.0)

    def test_json_serializable(self):
        import json

        json.dumps([cell.to_dict() for cell in list(self.sheet.cells.values())[:50]],
                   ensure_ascii=False)


@unittest.skipUnless(HAS_OPENPYXL, "目标机没有 openpyxl，跳过对拍用例")
class TestSampleXlsxAgainstOpenpyxl(unittest.TestCase):
    """用 openpyxl 对拍样本：值 / 合并区 / 公式 / 使用范围。"""

    @classmethod
    def setUpClass(cls):
        import openpyxl

        cls.wb = read_workbook(SAMPLE)
        cls.sheet = cls.wb.sheets[0]
        cls.o_values = openpyxl.load_workbook(SAMPLE, data_only=True).worksheets[0]
        cls.o_formulas = openpyxl.load_workbook(SAMPLE, data_only=False).worksheets[0]

    def test_used_range_matches(self):
        self.assertEqual(self.o_values.max_row, 63)
        self.assertEqual(self.o_values.max_column, 58)

    def test_merges_match(self):
        theirs = {(r.min_row - 1, r.max_row - 1, r.min_col - 1, r.max_col - 1)
                  for r in self.o_values.merged_cells.ranges}
        self.assertEqual(theirs, set(self.sheet.merges))

    def test_values_match(self):
        mismatches = []
        for row in self.o_values.iter_rows():
            for oc in row:
                if oc.value is None:
                    continue
                our = self.sheet.cells.get((oc.row - 1, oc.column - 1))
                if our is None:
                    mismatches.append(f"{oc.coordinate} 缺失")
                    continue
                formula = self.o_formulas.cell(row=oc.row, column=oc.column).value
                if isinstance(formula, str) and formula.startswith("="):
                    continue  # 公式单元格下面单独校验
                if isinstance(oc.value, bool):
                    ok = bool(our.raw) == oc.value
                elif isinstance(oc.value, (int, float)):
                    ok = our.num is not None and abs(our.num - float(oc.value)) < 1e-9
                else:
                    ok = our.text == str(oc.value)
                if not ok:
                    mismatches.append(f"{oc.coordinate}: {oc.value!r} != {our!r}")
        self.assertEqual(mismatches, [])

    def test_formula_text_matches(self):
        for row in self.o_formulas.iter_rows():
            for oc in row:
                formula = oc.value
                if not (isinstance(formula, str) and formula.startswith("=")):
                    continue  # ArrayFormula 对象跳过
                our = self.sheet.cells.get((oc.row - 1, oc.column - 1))
                self.assertIsNotNone(our, f"{oc.coordinate} 缺失")
                self.assertEqual(our.formula, formula, f"{oc.coordinate} 公式不一致")


# --------------------------------------------------------------------------
# 3. 合成 xlsx：共享串 / 内联串 / 布尔 / 错误 / 日期 / 公式 / 合并 / 隐藏表
# --------------------------------------------------------------------------
class TestOoxmlFeatures(_TmpMixin):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.path = _write_feature_xlsx(os.path.join(cls.tmp, "特征.xlsx"))
        cls.wb = read_workbook(cls.path)
        cls.sheet = cls.wb.sheets[0]

    def test_two_sheets_and_hidden(self):
        self.assertEqual(len(self.wb.sheets), 2)
        self.assertFalse(self.sheet.hidden)
        self.assertTrue(self.wb.sheets[1].hidden)
        self.assertEqual(self.wb.sheets[1].n_rows, 1)
        self.assertEqual(self.wb.sheets[1].n_cols, 1)

    def test_shared_and_inline_strings(self):
        self.assertEqual(self.sheet.cells[(0, 0)].text, "共享")
        self.assertEqual(self.sheet.cells[(0, 1)].text, "富文本")  # rPh 注音被忽略
        self.assertEqual(self.sheet.cells[(0, 2)].text, "内联")
        self.assertEqual(self.sheet.cells[(0, 2)].kind, KIND_TEXT)

    def test_number_bool_error(self):
        d1 = self.sheet.cells[(0, 3)]
        self.assertEqual(d1.kind, KIND_NUMBER)
        self.assertAlmostEqual(d1.num, 3.14)
        e1 = self.sheet.cells[(0, 4)]
        self.assertEqual(e1.raw, 42)
        self.assertEqual(e1.text, "42")
        f1 = self.sheet.cells[(0, 5)]
        self.assertEqual(f1.kind, KIND_BOOL)
        self.assertTrue(f1.raw)
        self.assertEqual(f1.text, "TRUE")
        g1 = self.sheet.cells[(0, 6)]
        self.assertEqual(g1.text, "#DIV/0!")

    def test_dates(self):
        h1 = self.sheet.cells[(0, 7)]  # 内置日期格式 14
        self.assertEqual(h1.kind, KIND_DATE)
        self.assertEqual(h1.text, "2024.1.1")
        self.assertEqual(h1.num, 45292.0)
        l1 = self.sheet.cells[(0, 11)]  # 自定义格式 yyyy"年"m"月"d"日"
        self.assertEqual(l1.kind, KIND_DATE)
        self.assertEqual(l1.text, "2024.1.1")

    def test_formulas(self):
        i1 = self.sheet.cells[(0, 8)]  # 有缓存值
        self.assertEqual(i1.formula, "=1+2")
        self.assertEqual(i1.num, 3.0)
        j1 = self.sheet.cells[(0, 9)]  # 无缓存值
        self.assertEqual(j1.kind, KIND_FORMULA)
        self.assertEqual(j1.formula, "=SUM(A1:A3)")
        self.assertEqual(j1.text, "=SUM(A1:A3)")
        k1 = self.sheet.cells[(0, 10)]  # 共享公式主格
        self.assertEqual(k1.formula, "=A1+1")
        k2 = self.sheet.cells[(1, 10)]  # 共享公式子格：有缓存值，公式文本不还原
        self.assertIsNone(k2.formula)
        self.assertEqual(k2.kind, KIND_NUMBER)
        self.assertEqual(k2.num, 3.0)
        k3 = self.sheet.cells[(2, 10)]  # 共享公式子格：无缓存值
        self.assertEqual(k3.kind, KIND_FORMULA)
        self.assertEqual(k3.text, "=<共享公式>")
        self.assertIsNone(k3.formula)

    def test_merges_and_extent(self):
        self.assertEqual(self.sheet.merges, [(1, 2, 0, 1)])
        self.assertEqual(self.sheet.n_rows, 3)
        self.assertEqual(self.sheet.n_cols, 12)

    def test_et_zip_dispatch(self):
        path = _write_feature_xlsx(self.tmp_path("WPS表.et"), ext=".et")
        wb = read_workbook(path)
        self.assertEqual(wb.fmt, "et-zip")
        self.assertEqual(len(wb.sheets), 2)

    def test_xlsm_dispatch(self):
        path = _write_feature_xlsx(self.tmp_path("宏表2.xlsx"), macro=True)
        self.assertEqual(detect_format(path), "xlsm")
        wb = read_workbook(path)
        self.assertEqual(wb.fmt, "xlsm")

    def test_xlsb_unsupported(self):
        path = self.tmp_path("二进制.xlsb")
        _write_zip(path, {
            "[Content_Types].xml": f'<Types xmlns="{_CT}"/>',
            "xl/workbook.bin": b"\x00",
        })
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertIn("xlsb", str(ctx.exception))
        self.assertRegex(str(ctx.exception), CJK_RE)


# --------------------------------------------------------------------------
# 4. 日期 / RK / 字符串工具
# --------------------------------------------------------------------------
class TestDateHelpers(unittest.TestCase):
    def test_serial_to_datetime(self):
        self.assertEqual(ooxml_mod.serial_to_datetime(45292), ooxml_mod._dt.datetime(2024, 1, 1))
        self.assertEqual(ooxml_mod.serial_to_datetime(1), ooxml_mod._dt.datetime(1900, 1, 1))
        self.assertEqual(ooxml_mod.serial_to_datetime(59), ooxml_mod._dt.datetime(1900, 2, 28))
        self.assertEqual(ooxml_mod.serial_to_datetime(61), ooxml_mod._dt.datetime(1900, 3, 1))
        self.assertEqual(ooxml_mod.serial_to_datetime(0, date1904=True),
                         ooxml_mod._dt.datetime(1904, 1, 1))

    def test_format_serial_date(self):
        self.assertEqual(ooxml_mod.format_serial_date(45292, False, True, False), "2024.1.1")
        self.assertEqual(ooxml_mod.format_serial_date(0.5, False, False, True), "12:00")
        self.assertEqual(ooxml_mod.format_serial_date(45292.25, False, True, True),
                         "2024.1.1 06:00")
        self.assertIsNone(ooxml_mod.format_serial_date(0.5, False, True, False))

    def test_classify_date_format(self):
        self.assertEqual(ooxml_mod.classify_date_format(14, None), (True, False))
        self.assertEqual(ooxml_mod.classify_date_format(22, None), (True, True))
        self.assertEqual(ooxml_mod.classify_date_format(45, None), (False, True))
        self.assertIsNone(ooxml_mod.classify_date_format(0, "#,##0.00"))
        self.assertIsNone(ooxml_mod.classify_date_format(164, "General"))
        self.assertIsNone(ooxml_mod.classify_date_format(164, "0.00%"))
        self.assertEqual(ooxml_mod.classify_date_format(164, "yyyy-mm-dd"), (True, False))
        self.assertEqual(ooxml_mod.classify_date_format(164, "hh:mm:ss"), (False, True))
        self.assertEqual(ooxml_mod.classify_date_format(164, "yyyy/m/d h:mm"), (True, True))


class TestBiffHelpers(unittest.TestCase):
    def test_rk_integer(self):
        rk = (123 << 2) | 0x02
        self.assertEqual(biff_mod._rk_value(rk), 123.0)

    def test_rk_negative(self):
        value = (-5) & 0x3FFFFFFF
        rk = (value << 2) | 0x02
        self.assertEqual(biff_mod._rk_value(rk), -5.0)

    def test_rk_divide_by_100(self):
        rk = ((12345 << 2) | 0x02) | 0x01
        self.assertAlmostEqual(biff_mod._rk_value(rk), 123.45)

    def test_rk_float(self):
        # RK 浮点只保存 IEEE754 高 32 位（低 32 位被丢弃），会有精度损失
        number = 12345.678
        bits = struct.unpack("<Q", struct.pack("<d", number))[0]
        rk = (bits >> 32) & 0xFFFFFFFC
        self.assertAlmostEqual(biff_mod._rk_value(rk), 12345.65625, places=6)

    def test_read_unicode_compressed(self):
        body = struct.pack("<H", 2) + b"\x00" + b"AB"
        text, off = biff_mod._read_unicode(body, 0, 2)
        self.assertEqual(text, "AB")
        self.assertEqual(off, 5)

    def test_read_unicode_high_byte(self):
        raw = "汉字".encode("utf-16-le")
        body = struct.pack("<H", 2) + b"\x01" + raw
        text, off = biff_mod._read_unicode(body, 0, 2)
        self.assertEqual(text, "汉字")
        self.assertEqual(off, 3 + len(raw))

    def test_sst_continue_split(self):
        """SST 字符串跨 CONTINUE 时，续记录首字节是新的编码标志。"""
        first = "AB".encode("latin-1")
        # 第一条字符串：cch=4, flags=0（压缩），字符在 SST 记录里放 3 个，跨到 CONTINUE
        sst_body = struct.pack("<II", 1, 1) + struct.pack("<H", 4) + b"\x00" + b"ABC"
        continue_body = b"\x00" + b"D"  # 首字节 0x00 = 仍然压缩，剩余 "D"
        strings = biff_mod._parse_sst([sst_body, continue_body], "测试.xls")
        self.assertEqual(strings, ["ABCD"])

    def test_read_unicode_lenbytes1(self):
        body = b"\x02\x00" + b"OK"
        text, off = biff_mod._read_unicode(body, 0, 1)
        self.assertEqual(text, "OK")
        self.assertEqual(off, 4)


# --------------------------------------------------------------------------
# 5. 真实 .xls / .et
# --------------------------------------------------------------------------
@unittest.skipUnless(os.path.isdir(CORPUS), "真实语料目录不存在")
class TestRealXls(unittest.TestCase):
    def _read(self, *names):
        path = _corpus_find(*names)
        if path is None:
            self.skipTest(f"语料中找不到：{names}")
        return path, read_workbook(path)

    def test_three_xls(self):
        cases = [
            ("即将超期工单截止3月16日18点.xls", 2, 186,
             ["Sheet0", "筛选分析-处理部门 (计数)"]),
            ("截至3月13日已扣分工单及到期即将扣分工单明细.xls", 4, 8800,
             ["截至3月13日本考核期已扣分工单", "3月13日至3月18日即将到期工单",
              "2024年未办结工单", "Sheet1"]),
            ("截至9月20日18时到期9件.xls", 1, 70, ["Sheet0"]),
        ]
        for name, n_sheets, n_cells, sheet_names in cases:
            with self.subTest(name=name):
                _path, wb = self._read(name)
                self.assertEqual(wb.fmt, "xls")
                self.assertEqual(len(wb.sheets), n_sheets)
                self.assertEqual([s.name for s in wb.sheets], sheet_names)
                self.assertEqual(sum(len(s.cells) for s in wb.sheets), n_cells)
                for sheet in wb.sheets:
                    self.assertGreaterEqual(sheet.n_rows, 1)
                    self.assertGreaterEqual(sheet.n_cols, 1)
                    for r1, r2, c1, c2 in sheet.merges:
                        self.assertLessEqual(r1, r2)
                        self.assertLessEqual(c1, c2)
                        self.assertLess(r2, sheet.n_rows)
                        self.assertLess(c2, sheet.n_cols)

    def test_first_cells(self):
        _path, wb = self._read("即将超期工单截止3月16日18点.xls")
        sheet = wb.sheets[0]
        self.assertEqual(sheet.cells[(0, 0)].text, "服务单编号")
        self.assertEqual(sheet.cells[(1, 0)].text, "DEMO-GD-24005")
        self.assertIn("昭乌达南路街道", sheet.cells[(1, 2)].text)

    def test_sst_continue_long_text(self):
        """大 .xls 的 SST 跨 CONTINUE：长文本必须完整解析（不能截断）。"""
        _path, wb = self._read("截至3月13日已扣分工单及到期即将扣分工单明细.xls")
        texts = [c.text for s in wb.sheets for c in s.cells.values()
                 if c.kind == KIND_TEXT and len(c.text) > 100]
        self.assertGreater(len(texts), 100)
        self.assertTrue(any(t.endswith("。") for t in texts))
        self.assertTrue(any("诉求" in t for t in texts))

    def test_biff_dates(self):
        """真实 .xls 里的日期序列号要转成 2024.3.9 风格的展示文本。"""
        path = _corpus_find("示例表格C.xls")
        if path is None:
            self.skipTest("语料中找不到日期样本 .xls")
        wb = read_workbook(path)
        dates = [c for s in wb.sheets for c in s.cells.values() if c.kind == KIND_DATE]
        self.assertGreater(len(dates), 0)
        self.assertTrue(any(re.match(r"^\d{4}\.\d{1,2}\.\d{1,2}", c.text) for c in dates))
        self.assertTrue(all(c.num and c.num > 40000 for c in dates))

    def test_et_ole(self):
        cases = [("示例表格E.et", 2, 26), ("示例表格F.et", 3, 23)]
        for name, n_sheets, n_cells in cases:
            with self.subTest(name=name):
                _path, wb = self._read(name)
                self.assertEqual(wb.fmt, "et-ole")
                self.assertEqual(len(wb.sheets), n_sheets)
                self.assertEqual(sum(len(s.cells) for s in wb.sheets), n_cells)
                self.assertTrue(any(s.merges for s in wb.sheets))


# --------------------------------------------------------------------------
# 6. 真实 CSV
# --------------------------------------------------------------------------
@unittest.skipUnless(os.path.isdir(CORPUS), "真实语料目录不存在")
class TestRealCsv(unittest.TestCase):
    def test_gis_csv_utf16_tab(self):
        path = _corpus_find("示例数据.csv")
        if path is None:
            self.skipTest("语料中找不到 CSV")
        wb = read_workbook(path)
        self.assertEqual(wb.fmt, "csv")
        self.assertEqual(len(wb.sheets), 1)
        sheet = wb.sheets[0]
        self.assertEqual(sheet.name, "园地复垦耕地")
        self.assertEqual(sheet.n_rows, 407)
        self.assertEqual(sheet.n_cols, 10)
        self.assertEqual(sheet.cells[(0, 0)].text, "文件夹")
        self.assertEqual(sheet.cells[(0, 1)].text, "名称")
        self.assertIn("utf-16", " ".join(wb.notes))
        self.assertIn("分隔符", " ".join(wb.notes))


class TestCsvReader(_TmpMixin):
    def test_utf8_sig(self):
        path = self.tmp_path("带BOM.csv")
        with open(path, "wb") as fp:
            fp.write("姓名,金额\n张三,100\n".encode("utf-8-sig"))
        wb = read_workbook(path)
        sheet = wb.sheets[0]
        self.assertEqual(sheet.cells[(0, 0)].text, "姓名")
        self.assertEqual(sheet.cells[(1, 0)].text, "张三")
        self.assertEqual(sheet.cells[(1, 1)].text, "100")

    def test_gbk(self):
        path = self.tmp_path("国标.csv")
        with open(path, "wb") as fp:
            fp.write("部门,人数\r\n工程部,12\r\n".encode("gbk"))
        wb = read_csv(path)
        self.assertEqual(wb.sheets[0].cells[(1, 0)].text, "工程部")
        self.assertIn("gbk", " ".join(wb.notes))

    def test_quoted_newline_and_trailing_blank(self):
        path = self.tmp_path("多行.csv")
        with open(path, "wb") as fp:
            fp.write('a,b\n"第一行\n第二行",2\n\n'.encode("utf-8"))
        wb = read_csv(path)
        sheet = wb.sheets[0]
        self.assertEqual(sheet.n_rows, 2)  # 末尾空行被裁掉
        self.assertEqual(sheet.cells[(1, 0)].text, "第一行\n第二行")

    def test_header_is_data_row(self):
        path = self.tmp_path("首行数据.csv")
        with open(path, "wb") as fp:
            fp.write("1,2,3\n4,5,6\n".encode("utf-8"))
        sheet = read_csv(path).sheets[0]
        self.assertEqual(sheet.n_rows, 2)
        self.assertEqual(sheet.cells[(0, 0)].text, "1")
        self.assertEqual(sheet.cells[(1, 2)].text, "6")

    def test_empty_csv(self):
        path = self.tmp_path("空.csv")
        with open(path, "wb") as fp:
            fp.write(b"")
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertRegex(str(ctx.exception), CJK_RE)


# --------------------------------------------------------------------------
# 7. 失败路径：必须抛 ValueError（中文原因），不许静默返回空表
# --------------------------------------------------------------------------
class TestFailures(_TmpMixin):
    def test_missing_file(self):
        with self.assertRaises(ValueError) as ctx:
            read_workbook(os.path.join(self.tmp, "不存在.xls"))
        self.assertRegex(str(ctx.exception), CJK_RE)

    def test_corrupt_ole(self):
        path = self.tmp_path("坏OLE.xls")
        with open(path, "wb") as fp:
            fp.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504)
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertRegex(str(ctx.exception), CJK_RE)

    def test_unknown_zip(self):
        path = self.tmp_path("普通压缩包.zip")
        _write_zip(path, {"readme.txt": "hello"})
        self.assertEqual(detect_format(path), "unknown")
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertRegex(str(ctx.exception), CJK_RE)

    def test_zip_without_sheets(self):
        path = self.tmp_path("没有表.xlsx")
        _write_zip(path, {
            "[Content_Types].xml": f'<Types xmlns="{_CT}"/>',
            "xl/workbook.xml": (
                f'<?xml version="1.0"?><workbook xmlns="{_NS_MAIN}"><sheets/></workbook>'),
        })
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertIn("工作表", str(ctx.exception))
        self.assertRegex(str(ctx.exception), CJK_RE)

    def test_corrupt_sheet_xml(self):
        path = _write_feature_xlsx(self.tmp_path("坏表.xlsx"))
        with zipfile.ZipFile(path) as zf:
            entries = {name: zf.read(name) for name in zf.namelist()}
        entries["xl/worksheets/sheet1.xml"] = b"<worksheet><sheetData><row>"
        _write_zip(path, entries)
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertRegex(str(ctx.exception), CJK_RE)
        self.assertIn("损坏", str(ctx.exception))

    def test_missing_shared_strings(self):
        path = _write_feature_xlsx(self.tmp_path("缺SST.xlsx"))
        with zipfile.ZipFile(path) as zf:
            entries = {name: zf.read(name) for name in zf.namelist()}
        entries.pop("xl/sharedStrings.xml", None)
        _write_zip(path, entries)
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertIn("共享字符串", str(ctx.exception))
        self.assertRegex(str(ctx.exception), CJK_RE)

    def test_binary_csv(self):
        path = self.tmp_path("其实是二进制.csv")
        with open(path, "wb") as fp:
            fp.write(b"\x00\x01\x02\xff" * 200)
        self.assertEqual(detect_format(path), "unknown")
        with self.assertRaises(ValueError) as ctx:
            read_workbook(path)
        self.assertRegex(str(ctx.exception), CJK_RE)


# --------------------------------------------------------------------------
# 8. 代码零第三方依赖
# --------------------------------------------------------------------------
class TestNoThirdPartyImports(unittest.TestCase):
    def test_reader_sources_are_stdlib_only(self):
        import inspect

        modules = [ooxml_mod, biff_mod, sys.modules["app.core.sniff"],
                   sys.modules["app.core.readers.csvr"]]
        pattern = re.compile(r"^\s*(?:import|from)\s+(openpyxl|pandas|xlrd|numpy)\b", re.M)
        for module in modules:
            source = inspect.getsource(module)
            self.assertIsNone(pattern.search(source),
                              f"{module.__name__} 引入了第三方依赖")


if __name__ == "__main__":
    unittest.main(verbosity=2)
