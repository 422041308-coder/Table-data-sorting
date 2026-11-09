# -*- coding: utf-8 -*-
"""S4 字段映射与模板改造测试：模板解析 / 克隆注入保样式 / 转置 / 规则引擎 / 服务层。

运行：python -m unittest tests.test_convert -v
只用标准库 + 项目自身模块；样本用 samples/项目台账统计表.xlsx。
"""
import json
import os
import re
import shutil
import sys
import unittest
import uuid
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app.core import mapping, template  # noqa: E402
from app.core.model import (  # noqa: E402
    Cell,
    CellRef,
    Field,
    FileRef,
    MappingProfile,
    MappingRule,
    Sel,
    TableModel,
    TemplateDef,
)
from app.core.xlsxwriter import write_xlsx  # noqa: E402
from app.services import mapping_service as svc  # noqa: E402

SAMPLE = os.path.join(ROOT, "samples", "项目台账统计表.xlsx")
TEST_ROOT = os.path.join(ROOT, "data", "tmp")
os.makedirs(TEST_ROOT, exist_ok=True)
# 每次运行独立临时目录（工作区内自建；本沙箱下 tempfile.mkdtemp 造出的目录不可写）
TMP = os.path.join(TEST_ROOT, "test_convert_" + uuid.uuid4().hex[:8])
os.makedirs(TMP, exist_ok=True)
OUTPUT = os.path.join(ROOT, "data", "output")
SAMPLE_SHEET = "2024年项目台账（自筹-甲代-施工-日常经营）"


def tearDownModule():
    shutil.rmtree(TMP, ignore_errors=True)


# ==========================================================================
# 测试夹具
# ==========================================================================
def _ensure_tmp():
    os.makedirs(TMP, exist_ok=True)
    os.makedirs(OUTPUT, exist_ok=True)


def _copy_file(src: str, dst: str) -> str:
    """复制文件且不继承源文件的只读属性（样本文件可能是只读的）。"""
    with open(src, "rb") as f:
        data = f.read()
    if os.path.exists(dst):
        try:
            os.chmod(dst, 0o666)
        except OSError:
            pass
        try:
            os.remove(dst)
        except OSError:
            pass
    with open(dst, "wb") as f:
        f.write(data)
    try:
        os.chmod(dst, 0o666)
    except OSError:
        pass
    return dst


def _write_simple_template(path: str) -> None:
    write_xlsx(
        path,
        [
            {
                "name": "台账",
                "header": ["项目名称", "合同编码", "金额", "备注", "固定列", "转置目标"],
                "rows": [],
                "widths": [18, 16, 12, 20, 12, 14],
                "freeze": "A2",
            }
        ],
    )


def _write_horizontal_source(path: str) -> None:
    """A 表：第 1 行表头，第 2 行占位数据（带数字，便于识别为数据行），
    第 3 行是"字段名→值"横排（B..G 列）。"""
    header = ["姓名", "姓名值", "性别", "性别值", "年龄", "年龄值", "部门", "部门值", "电话", "电话值", "备注", "备注值"]
    row2 = [1, "x", 2, "y", 3, "z", 4, "w", 5, "v", 6, "u"]
    row3 = [None, "姓名", "张三", "性别", "男", "年龄", 30, "部门", "工程部", "电话", "13800000000", "备注", "借调"]
    write_xlsx(path, [{"name": "A表", "header": header, "rows": [row2, row3]}])


def _write_records_source(path: str, n: int = 20) -> None:
    header = ["序号", "项目名称", "合同编码", "金额", "实施部门"]
    rows = [
        [i + 1, f"测试项目{i + 1:02d}", f"HT-2024-{i + 1:03d}", 10000 + i * 100, "工程部" if i % 2 == 0 else "计划统筹部"]
        for i in range(n)
    ]
    write_xlsx(path, [{"name": "过程表", "header": header, "rows": rows}])


def _simple_template(path: str | None = None) -> TemplateDef:
    path = path or os.path.join(TMP, "simple_template.xlsx")
    if not os.path.isfile(path):
        _write_simple_template(path)
    return template.load_template(path)


def _horizontal_source() -> TableModel:
    """手造 TableModel：Excel 第 3 行（0 基 2）是横排的字段名/值。"""
    row3 = {
        1: Cell(raw="姓名", kind="text", text="姓名"),
        2: Cell(raw="张三", kind="text", text="张三"),
        3: Cell(raw="性别", kind="text", text="性别"),
        4: Cell(raw="男", kind="text", text="男"),
        5: Cell(raw="年龄", kind="text", text="年龄"),
        6: Cell(raw=30, kind="number", text="30", num=30.0),
        7: Cell(raw="部门", kind="text", text="部门"),
        8: Cell(raw="工程部", kind="text", text="工程部"),
    }
    fields = [
        Field(index=1, name="姓名", kind_hint="text"),
        Field(index=2, name="姓名值", kind_hint="text"),
    ]
    return TableModel(
        source=FileRef(id="src", path=os.path.join(TMP, "A表.xlsx"), name="A表.xlsx", fmt="xlsx"),
        sheet="A表",
        header_rows=[0],
        data_start=2,
        data_end=2,
        fields=fields,
        rows=[row3],
        row_numbers=[2],
    )


def _records_source() -> TableModel:
    fields = [
        Field(index=0, name="项目名称", kind_hint="text"),
        Field(index=1, name="合同编码", kind_hint="code"),
        Field(index=2, name="金额", kind_hint="money"),
        Field(index=3, name="联系电话", kind_hint="phone"),
    ]
    rows = [
        {0: Cell(raw="示例项目", kind="text", text="示例项目"), 1: Cell(raw="HT-001", kind="text", text="HT-001"),
         2: Cell(raw=1000, kind="number", text="1000", num=1000.0), 3: Cell(raw="13800000001", kind="text", text="13800000001")},
        {1: Cell(raw="HT-002", kind="text", text="HT-002"), 2: Cell(raw="2000", kind="text", text="2000"),
         3: Cell(raw="13800000002", kind="text", text="13800000002")},
        {1: Cell(raw="HT-003", kind="text", text="HT-003"), 2: Cell(raw="3000", kind="text", text="3000"),
         3: Cell(raw="13800000003", kind="text", text="13800000003")},
    ]
    return TableModel(
        source=FileRef(id="src2", path=os.path.join(TMP, "B表.xlsx"), name="B表.xlsx", fmt="xlsx"),
        sheet="B表",
        header_rows=[0],
        data_start=1,
        data_end=3,
        fields=fields,
        rows=rows,
        row_numbers=[1, 2, 3],
    )


def _transpose_rule() -> MappingRule:
    return MappingRule(
        source={"row": Sel("fixed", 3), "col": Sel("span", {"start": 2, "count": 6})},
        target={"row": Sel("span", {"start": 2, "count": 6}), "col": Sel("fixed", 6)},
        mode="transpose",
        note="A 表第 3 行横排 → B 表 F 列竖排",
    )


def _rows_from_cells(cells: dict, data_start: tuple[int, int]) -> list[dict]:
    """{(r0,c0): value} → inject(rows=[{1基列: 值}])。"""
    ds_row = data_start[0]
    by_row: dict[int, dict[int, object]] = {}
    for (r0, c0), v in cells.items():
        r1, c1 = r0 + 1, c0 + 1
        if r1 < ds_row:
            continue
        by_row.setdefault(r1, {})[c1] = v
    if not by_row:
        return []
    return [by_row.get(r, {}) for r in range(ds_row, max(by_row) + 1)]


# ==========================================================================
# 1. 模板解析
# ==========================================================================
class TestLoadTemplate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _ensure_tmp()

    def test_sample_layout(self):
        tpl = template.load_template(SAMPLE)
        self.assertEqual(tpl.sheet, SAMPLE_SHEET)
        self.assertIsNotNone(tpl.title_range)
        self.assertEqual(tuple(tpl.header_range), (2, 3, 1, 58))  # 两级表头
        self.assertEqual(tuple(tpl.data_start), (4, 1))
        self.assertEqual(len(tpl.fields), 58)
        names = [f.name for f in tpl.fields]
        self.assertEqual(names[0], "序号")
        self.assertEqual(names[1], "项目名称")
        # 多级表头用 " / " 连接，父级跨列标题向下传播
        self.assertIn("计划部前期相关资料 / 项目决策部门", names)
        self.assertIn("合同签订情况 / 合同编码", names)
        self.assertIn("乙方中标金额 / 含税金额（元）", names)
        # 字段列号（0 基）：AF=31、BE=56
        af = next(f for f in tpl.fields if f.name == "合同签订情况 / 合同编码")
        self.assertEqual(af.index, 31)
        self.assertEqual(af.kind_hint, "code")
        be = next(f for f in tpl.fields if f.name == "实施部门")
        self.assertEqual(be.index, 56)

    def test_simple_template(self):
        tpl = _simple_template()
        self.assertEqual(tpl.sheet, "台账")
        self.assertEqual(tuple(tpl.header_range), (1, 1, 1, 6))
        self.assertEqual(tuple(tpl.data_start), (2, 1))
        self.assertEqual([f.name for f in tpl.fields], ["项目名称", "合同编码", "金额", "备注", "固定列", "转置目标"])

    def test_sheet_switch_and_missing_sheet(self):
        self.assertEqual(len(template.list_sheets(SAMPLE)), 2)
        tpl2 = template.load_template(SAMPLE, "Sheet1")
        self.assertEqual(tpl2.sheet, "Sheet1")
        self.assertLessEqual(len(tpl2.fields), 1)
        with self.assertRaises(ValueError):
            template.load_template(SAMPLE, "不存在的表")

    def test_not_xlsx(self):
        p = os.path.join(TMP, "notzip.xlsx")
        with open(p, "wb") as f:
            f.write(b"not a zip file")
        with self.assertRaises(ValueError):
            template.load_template(p)


# ==========================================================================
# 2. 克隆注入：保样式 / 值 / 公式 / 合并 / dimension / calcChain
# ==========================================================================
class TestInject(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _ensure_tmp()
        cls.out_dir = os.path.join(TMP, "inject")
        os.makedirs(cls.out_dir, exist_ok=True)

    def _make_rows(self, n=20):
        rows = []
        for i in range(n):
            rows.append(
                {
                    1: i + 1,
                    2: f"测试项目{i + 1:02d}",
                    32: f"HT-2024-{i + 1:03d}",
                    29: 10000 + i * 100,
                    31: "已签订",
                }
            )
        return rows

    def test_sample_template_style_preserved(self):
        out = os.path.join(self.out_dir, "台账_20行.xlsx")
        rows = self._make_rows(20)
        path = template.inject(SAMPLE, out, (4, 1), rows)
        self.assertTrue(os.path.isfile(path))
        self.assertTrue(zipfile.is_zipfile(path))

        with zipfile.ZipFile(SAMPLE) as a, zipfile.ZipFile(path) as b:
            self.assertEqual(a.namelist(), b.namelist())
            # 只允许 workbook.xml 与目标 sheet XML 变化，其余条目字节完全一致
            changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
            self.assertEqual(sorted(changed), ["xl/workbook.xml", "xl/worksheets/sheet1.xml"])
            # 表头/标题行原样保留（逐字节）
            sheet_old = a.read("xl/worksheets/sheet1.xml").decode("utf-8")
            sheet_new = b.read("xl/worksheets/sheet1.xml").decode("utf-8")
            for r in (1, 2, 3):
                old_row = re.search(r'<row r="%d"[^>]*>.*?</row>' % r, sheet_old, re.S).group(0)
                new_row = re.search(r'<row r="%d"[^>]*>.*?</row>' % r, sheet_new, re.S).group(0)
                self.assertEqual(old_row, new_row, f"第 {r} 行（标题/表头）被改动")
            # dimension 更新
            self.assertIn('<dimension ref="A1:BF23"/>', sheet_new)
            # 数据区样板合并（A4:A7 / B4:B7 / C4:C7 / AL4:AL7）被移除，表头合并保留
            self.assertNotIn('A4:A7', sheet_new)
            self.assertIn('A1:BF1', sheet_new)
            self.assertIn('H2:W2', sheet_new)
            # 数组公式被克隆并平移行引用（第 4 行是第一条数据，行引用保持 4；第 5 行递增到 5）
            self.assertIn('ref="AV5"', sheet_new)
            self.assertIn('AN5-AT5', sheet_new)
            self.assertIn('ref="AV23"', sheet_new)
            self.assertIn('AN23-AT23', sheet_new)
            # 强制重算
            wb = b.read("xl/workbook.xml").decode("utf-8")
            self.assertIn('fullCalcOnLoad="1"', wb)

        g = template.read_grid(path)
        self.assertEqual(g["dimension"], "A1:BF23")
        self.assertEqual(g["cells"].get((3, 0)), 1)
        self.assertEqual(g["cells"].get((3, 1)), "测试项目01")
        self.assertEqual(g["cells"].get((3, 31)), "HT-2024-001")
        self.assertEqual(g["cells"].get((3, 28)), 10000)
        self.assertEqual(g["cells"].get((22, 1)), "测试项目20")
        # 超出数据区之后不应再有模板旧数据（原第 8 行小计）
        self.assertIsNone(g["cells"].get((23, 1)))
        self.assertEqual(template.verify_injected(path, (4, 1), rows), [])

    def test_value_types_and_escaping(self):
        out = os.path.join(self.out_dir, "types.xlsx")
        rows = [
            {
                1: 1,
                2: '文本 <&>"引号" 换行\n第二行',
                3: 0.06,
                4: True,
                5: date(2024, 3, 9),
                6: datetime(2024, 3, 9, 12, 30),
            }
        ]
        path = template.inject(SAMPLE, out, (4, 1), rows)
        g = template.read_grid(path)
        self.assertEqual(g["cells"].get((3, 1)), '文本 <&>"引号" 换行\n第二行')
        self.assertAlmostEqual(g["cells"].get((3, 2)), 0.06)
        self.assertTrue(g["cells"].get((3, 3)))
        self.assertAlmostEqual(g["cells"].get((3, 4)), 45360.0)  # 2024-03-09 序列值
        self.assertAlmostEqual(g["cells"].get((3, 5)), 45360.5208333, places=4)
        # 语法/结构仍合法
        with zipfile.ZipFile(path) as z:
            ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

    def test_empty_and_missing_sample_row(self):
        out1 = os.path.join(self.out_dir, "empty.xlsx")
        p1 = template.inject(SAMPLE, out1, (4, 1), [])
        g1 = template.read_grid(p1)
        self.assertEqual(g1["dimension"], "A1:BF3")

        tpl = _simple_template(os.path.join(TMP, "simple2.xlsx"))
        out2 = os.path.join(self.out_dir, "no_sample_row.xlsx")
        p2 = template.inject(
            os.path.join(TMP, "simple2.xlsx"), out2, tuple(tpl.data_start), [{1: "名称", 2: "HT-1"}]
        )
        g2 = template.read_grid(p2)
        self.assertEqual(g2["cells"].get((1, 0)), "名称")
        self.assertEqual(g2["cells"].get((1, 1)), "HT-1")
        with zipfile.ZipFile(p2) as z:
            ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

    def test_raise_on_same_path_and_bad_args(self):
        with self.assertRaises(ValueError):
            template.inject(SAMPLE, SAMPLE, (4, 1), [])
        with self.assertRaises(ValueError):
            template.inject(SAMPLE, os.path.join(self.out_dir, "x.xlsx"), (0, 1), [])
        with self.assertRaises(ValueError):
            template.inject(SAMPLE, os.path.join(self.out_dir, "x.xlsx"), (4, 1), ["不是dict"])
        with self.assertRaises(ValueError):
            template.inject(SAMPLE, os.path.join(self.out_dir, "x.xlsx"), (4, 1), [{1: 1}], sheet="无此表")

    def test_calcchain_removed(self):
        """模板若带 calcChain，输出必须删除条目、Content_Types 与 rels 引用。"""
        src = os.path.join(TMP, "with_calcchain.xlsx")
        _copy_file(SAMPLE, src)
        # 注入一个假的 calcChain（内容可以极简，inject 只要识别并删除它）
        with zipfile.ZipFile(src) as z:
            entries = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
        entries["xl/calcChain.xml"] = (
            None,
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
        )
        for n in ("[Content_Types].xml", "xl/_rels/workbook.xml.rels"):
            info, data = entries[n]
            text = data.decode("utf-8")
            if n.endswith("Content_Types].xml"):
                text = text.replace(
                    "</Types>",
                    '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
                )
            else:
                text = text.replace(
                    "</Relationships>",
                    '<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
                )
            entries[n] = (info, text.encode("utf-8"))
        with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as zo:
            for name, (info, data) in entries.items():
                zo.writestr(info or name, data)

        out = os.path.join(self.out_dir, "calcchain_out.xlsx")
        template.inject(src, out, (4, 1), self._make_rows(2))
        with zipfile.ZipFile(out) as z:
            names = z.namelist()
            self.assertNotIn("xl/calcChain.xml", names)
            self.assertNotIn("calcChain", z.read("[Content_Types].xml").decode("utf-8"))
            self.assertNotIn("calcChain", z.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
            ET.fromstring(z.read("xl/worksheets/sheet1.xml"))


# ==========================================================================
# 3. 规则引擎：cell / transpose / filldown / repeat / block / Sel
# ==========================================================================
class TestMappingEngine(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _ensure_tmp()
        cls.tpl = _simple_template()

    def test_header_sel_leaf_priority(self):
        """回归：字段名"合同编码"不能按包含匹配错抢到「委托单位/甲方（插入评论合同编码）」的 E 列。"""
        tpl = template.load_template(SAMPLE)
        src = TableModel(
            source=FileRef(id="s", path="", name=""),
            sheet="来源",
            header_rows=[0],
            data_start=3,
            data_end=3,
            fields=[
                Field(index=4, name="委托单位/甲方（插入评论合同编码）", kind_hint="code"),
                Field(index=31, name="合同签订情况 / 合同编码", kind_hint="code"),
            ],
            rows=[{4: Cell(raw="E4-值", kind="text", text="E4-值"),
                   31: Cell(raw="AF4-值", kind="text", text="AF4-值")}],
            row_numbers=[3],
        )
        rule = MappingRule(
            source={"row": Sel("record", 0), "col": Sel("header", "合同编码")},
            target={"row": Sel("record", 0), "col": Sel("header", "合同编码")},
            mode="cell",
        )
        prof = MappingProfile(id="p-leaf", name="叶子名优先", template_id=tpl.id, rules=[rule])
        cells = mapping.apply_rules(prof, src, tpl)
        # 来源 AF（列 32，0 基 31）→ 目标 AF4（0 基行 3、列 31）
        self.assertEqual(cells, {(3, 31): "AF4-值"})
        pv = mapping.preview(prof, src, tpl, limit=5)
        self.assertEqual(pv["cells"][0]["sourceRef"]["a1"], "AF4")
        self.assertEqual(pv["cells"][0]["target"]["a1"], "AF4")

    def test_transpose_row_to_column(self):
        """核心验收用例：A 表第 3 行横排 → B 表 F 列竖排。"""
        src = _horizontal_source()
        prof = MappingProfile(id="p1", name="转置", template_id=self.tpl.id, rules=[_transpose_rule()])
        cells = mapping.apply_rules(prof, src, self.tpl)
        self.assertEqual(
            {(r, c): v for (r, c), v in cells.items()},
            {
                (1, 5): "姓名",
                (2, 5): "张三",
                (3, 5): "性别",
                (4, 5): "男",
                (5, 5): "年龄",
                (6, 5): 30,
            },
        )
        pv = mapping.preview(prof, src, self.tpl, limit=10)
        self.assertFalse(pv["truncated"])
        self.assertEqual(len(pv["cells"]), 6)
        self.assertEqual([c["target"]["a1"] for c in pv["cells"]], ["F2", "F3", "F4", "F5", "F6", "F7"])
        self.assertEqual([c["value"] for c in pv["cells"]], ["姓名", "张三", "性别", "男", "年龄", 30])
        # 来源坐标必须可追溯：B3..G3
        self.assertEqual([c["sourceRef"]["a1"] for c in pv["cells"]], ["B3", "C3", "D3", "E3", "F3", "G3"])
        self.assertEqual([c["ruleIndex"] for c in pv["cells"]], [0] * 6)

    def test_transpose_output_matches_preview(self):
        """预演清单与实际 inject 输出逐格一致。"""
        src = _horizontal_source()
        tpl_path = os.path.join(TMP, "simple_template.xlsx")
        prof = MappingProfile(id="p1", name="转置", template_id=self.tpl.id, rules=[_transpose_rule()])
        pv = mapping.preview(prof, src, self.tpl, limit=100)
        cells = mapping.apply_rules(prof, src, self.tpl)
        rows = _rows_from_cells(cells, self.tpl.data_start)
        out = os.path.join(TMP, "transposed.xlsx")
        template.inject(tpl_path, out, tuple(self.tpl.data_start), rows)
        g = template.read_grid(out)
        for c in pv["cells"]:
            t = c["target"]
            got = g["cells"].get((t["row"] - 1, t["col"] - 1))
            self.assertEqual(got, c["value"], f"预演/输出不一致：{t['a1']} → {got!r} != {c['value']!r}")

    def test_cell_per_record_and_transforms(self):
        src = _records_source()
        rules = [
            MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "项目名称")},
                        target={"row": Sel("record", 0), "col": Sel("header", "项目名称")}, mode="cell"),
            MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "合同编码")},
                        target={"row": Sel("record", 0), "col": Sel("fixed", 2)}, mode="cell"),
            MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "金额")},
                        target={"row": Sel("record", 0), "col": Sel("fixed", 3)}, mode="cell", transform="number"),
            MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "联系电话")},
                        target={"row": Sel("record", 0), "col": Sel("fixed", 4)}, mode="cell", transform="date"),
        ]
        prof = MappingProfile(id="p2", name="字段映射", template_id=self.tpl.id, rules=rules)
        cells = mapping.apply_rules(prof, src, self.tpl)
        # 3 条记录 → 目标行 2..4（1 基）
        self.assertEqual(cells[(1, 0)], "示例项目")
        self.assertEqual(cells[(1, 1)], "HT-001")
        self.assertEqual(cells[(1, 2)], 1000)
        self.assertEqual(cells[(2, 0)], "")  # filldown 未开：第二条记录项目名为空 → blank
        self.assertEqual(cells[(2, 1)], "HT-002")
        self.assertEqual(cells[(3, 2)], 3000)
        # 日期：11 位手机号不是合法日期 → 保留原值并告警
        pv = mapping.preview(prof, src, self.tpl, limit=10)
        warnings = " ".join(pv["warnings"])
        self.assertIn("无法把", warnings)
        self.assertEqual(cells[(1, 3)], "13800000001")

    def test_filldown(self):
        src = _records_source()
        rule = MappingRule(
            source={"row": Sel("record", 0), "col": Sel("header", "项目名称")},
            target={"row": Sel("record", 0), "col": Sel("fixed", 1)},
            mode="filldown",
        )
        prof = MappingProfile(id="p3", name="填充", template_id=self.tpl.id, rules=[rule])
        cells = mapping.apply_rules(prof, src, self.tpl)
        self.assertEqual([cells[(r, 0)] for r in (1, 2, 3)], ["示例项目", "示例项目", "示例项目"])

    def test_repeat_constant_and_on_missing(self):
        src = _records_source()
        rules = [
            MappingRule(source={"row": Sel("record", 0), "col": Sel("fixed", 1)},
                        target={"row": Sel("record", 0), "col": Sel("header", "备注")},
                        mode="repeat", transform="constant", constant="固定备注"),
            MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "联系电话")},
                        target={"row": Sel("record", 0), "col": Sel("fixed", 5)},
                        mode="cell", on_missing="skip"),
        ]
        prof = MappingProfile(id="p4", name="常量", template_id=self.tpl.id, rules=rules)
        cells = mapping.apply_rules(prof, src, self.tpl)
        self.assertEqual([cells[(r, 3)] for r in (1, 2, 3)], ["固定备注"] * 3)

        # on_missing=error 时 apply_rules 必须抛错，preview 只告警
        bad = MappingProfile(
            id="p5", name="缺字段", template_id=self.tpl.id,
            rules=[MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "不存在的字段")},
                               target={"row": Sel("record", 0), "col": Sel("fixed", 1)},
                               mode="cell", on_missing="error")],
        )
        with self.assertRaises(ValueError):
            mapping.apply_rules(bad, src, self.tpl)
        pv = mapping.preview(bad, src, self.tpl, limit=10)
        self.assertTrue(any("找不到字段" in w for w in pv["warnings"]))

    def test_block_mode(self):
        src = _records_source()
        # 源区域 B2:C3（合同编码/金额两列的首两行）整块搬到目标 A2:B3
        rule = MappingRule(
            source={"row": Sel("fixed", 2), "col": Sel("fixed", 2), "rows": 2, "cols": 2},
            target={"row": Sel("fixed", 2), "col": Sel("fixed", 1), "rows": 2, "cols": 2},
            mode="block",
        )
        prof = MappingProfile(id="p6", name="区域", template_id=self.tpl.id, rules=[rule])
        pv = mapping.preview(prof, src, self.tpl, limit=10)
        self.assertEqual(
            [(c["target"]["a1"], c["value"]) for c in pv["cells"]],
            [("A2", "HT-001"), ("B2", 1000), ("A3", "HT-002"), ("B3", "2000")],
        )

    def test_sel_kinds_and_sheet_mismatch_warning(self):
        src = _records_source()
        # header/fixed/record/span 均已覆盖；这里验证跨表与越界告警
        prof = MappingProfile(
            id="p7", name="告警", template_id=self.tpl.id,
            rules=[
                MappingRule(source={"sheet": "别的表", "row": Sel("record", 0), "col": Sel("fixed", 1)},
                            target={"row": Sel("record", 0), "col": Sel("fixed", 1)}, mode="cell"),
                MappingRule(source={"row": Sel("fixed", 1), "col": Sel("fixed", 1)},
                            target={"row": Sel("fixed", 1), "col": Sel("fixed", 9)}, mode="cell"),
            ],
        )
        pv = mapping.preview(prof, src, self.tpl, limit=10)
        text = " ".join(pv["warnings"])
        self.assertIn("来源工作表", text)
        self.assertIn("早于模板数据起始行", text)

    def test_preview_limit_and_truncated(self):
        src = _records_source()
        prof = MappingProfile(
            id="p8", name="分页", template_id=self.tpl.id,
            rules=[MappingRule(source={"row": Sel("record", 0), "col": Sel("header", "合同编码")},
                               target={"row": Sel("record", 0), "col": Sel("fixed", 2)}, mode="cell")],
        )
        pv = mapping.preview(prof, src, self.tpl, limit=2)
        self.assertTrue(pv["truncated"])
        self.assertEqual(len(pv["cells"]), 2)
        self.assertEqual(pv["recordCount"], 3)
        self.assertTrue(any("预演只展示前 2 行" in w for w in pv["warnings"]))

    def test_transpose_requires_region_anchor(self):
        """source.col=span 未给 start 时，从来源首列开始，仍能工作。"""
        src = _horizontal_source()
        rule = MappingRule(
            source={"row": Sel("fixed", 3), "col": Sel("span", 2)},
            target={"row": Sel("span", {"start": 2, "count": 2}), "col": Sel("fixed", 6)},
            mode="transpose",
        )
        prof = MappingProfile(id="p9", name="转置2", template_id=self.tpl.id, rules=[rule])
        cells = mapping.apply_rules(prof, src, self.tpl)
        # 来源默认首列 = 第一个字段列（B），即 B3/C3 = 姓名/张三
        self.assertEqual(cells[(1, 5)], "姓名")
        self.assertEqual(cells[(2, 5)], "张三")


# ==========================================================================
# 4. 自动推荐
# ==========================================================================
class TestSuggest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _ensure_tmp()
        cls.tpl = _simple_template()

    def _source(self, names):
        return TableModel(
            source=FileRef(id="s", path="", name=""),
            sheet="S",
            header_rows=[0],
            data_start=1,
            fields=[Field(index=i, name=n, kind_hint="text") for i, n in enumerate(names)],
        )

    def test_alias_and_similarity(self):
        src = self._source(["项目名", "合同编号", "金额（元）", "备注说明", "联系人"])
        sugg = mapping.suggest(src, self.tpl)
        mapping_by_target = {s["targetField"]: s for s in sugg}
        self.assertIn("项目名称", mapping_by_target)
        self.assertEqual(mapping_by_target["项目名称"]["sourceField"], "项目名")
        self.assertGreaterEqual(mapping_by_target["项目名称"]["confidence"], 0.8)
        self.assertIn("合同编码", mapping_by_target)
        self.assertEqual(mapping_by_target["合同编码"]["sourceField"], "合同编号")
        self.assertIn("金额", mapping_by_target)
        self.assertIn("备注", mapping_by_target)
        # 推荐结果带可直接落库的 rule
        self.assertIn("rule", mapping_by_target["合同编码"])
        self.assertEqual(mapping_by_target["合同编码"]["rule"]["mode"], "cell")

    def test_no_duplicate_assignment(self):
        src = self._source(["合同名称", "合同名称"])
        sugg = mapping.suggest(src, self.tpl)
        sources = [s["sourceIndex"] for s in sugg]
        self.assertEqual(len(sources), len(set(sources)))

    def test_empty(self):
        empty = TableModel(source=None, sheet="S")
        self.assertEqual(mapping.suggest(empty, self.tpl), [])


# ==========================================================================
# 5. 服务层：模板登记 / 方案 CRUD / 推荐 / 预演 / 转换
# ==========================================================================
class TestMappingService(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _ensure_tmp()
        cls.sample_copy = os.path.join(TMP, "sample_template_copy.xlsx")
        _copy_file(SAMPLE, cls.sample_copy)
        cls.simple_path = os.path.join(TMP, "service_template.xlsx")
        _write_simple_template(cls.simple_path)
        cls.horizontal_path = os.path.join(TMP, "service_horizontal.xlsx")
        _write_horizontal_source(cls.horizontal_path)
        cls.records_path = os.path.join(TMP, "service_records.xlsx")
        _write_records_source(cls.records_path, 20)
        cls.template_ids: list[str] = []
        cls.profile_ids: list[str] = []
        cls.outputs: list[str] = []

    @classmethod
    def tearDownClass(cls):
        for tid in cls.template_ids:
            try:
                svc.delete_template(tid)
            except Exception:  # noqa: BLE001
                pass
        for pid in cls.profile_ids:
            try:
                svc.delete_profile(pid)
            except Exception:  # noqa: BLE001
                pass
        for p in cls.outputs:
            try:
                os.remove(p)
            except OSError:
                pass

    def _add(self, path, name=None, **extra):
        body = {"filePath": path}
        if name:
            body["name"] = name
        body.update(extra)
        r = svc.add_template(body)
        tid = r["template"]["id"]
        self.template_ids.append(tid)
        return r["template"]

    def test_templates_crud_and_anchor(self):
        tpl = self._add(self.sample_copy, "样本台账模板")
        self.assertEqual(tpl["name"], "样本台账模板")
        self.assertEqual(tuple(tpl["headerRange"]), (2, 3, 1, 58))
        self.assertEqual(tuple(tpl["dataStart"]), (4, 1))
        self.assertEqual(len(tpl["fields"]), 58)
        self.assertIn(tpl["id"], [t["id"] for t in svc.templates()])

        # 重新框选锚点：单行表头 + 数据起始行 3
        simple = self._add(self.simple_path, "简单模板")
        r = svc.set_anchor(simple["id"], {"headerRange": "A1:F1", "dataStart": "A2"})
        self.assertEqual(tuple(r["template"]["headerRange"]), (1, 1, 1, 6))
        self.assertEqual(tuple(r["template"]["dataStart"]), (2, 1))
        # 用锚点对象形式再改一次
        r2 = svc.set_anchor(simple["id"], {"headerRange": [1, 1, 1, 6], "dataStart": [2, 1]})
        self.assertEqual(tuple(r2["template"]["dataStart"]), (2, 1))

        del_r = svc.delete_template(simple["id"])
        self.assertEqual(del_r["deleted"], simple["id"])
        self.assertNotIn(simple["id"], [t["id"] for t in svc.templates()])
        self.template_ids.remove(simple["id"])
        with self.assertRaises(ValueError):
            svc.set_anchor("不存在的模板", {"dataStart": [2, 1]})

    def test_add_template_uploads_bytes(self):
        with open(self.simple_path, "rb") as f:
            raw = f.read()
        r = svc.add_template({"__filename__": "上传模板.xlsx", "__raw__": raw, "name": "上传模板"})
        tid = r["template"]["id"]
        self.template_ids.append(tid)
        self.assertTrue(os.path.isfile(r["template"]["filePath"]))
        self.assertEqual(r["template"]["name"], "上传模板")

    def test_profiles_crud(self):
        tpl = self._add(self.simple_path, "方案模板")
        rules = [
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "姓名"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 1}},
             "mode": "cell", "transform": "identity", "onMissing": "blank", "constant": None, "note": ""},
        ]
        r = svc.save_profile({"name": "方案A", "templateId": tpl["id"], "rules": rules, "sourceSignature": "sig-1"})
        pid = r["profile"]["id"]
        self.profile_ids.append(pid)
        self.assertEqual(r["profile"]["sourceSignature"], "sig-1")
        self.assertIn(pid, [p["id"] for p in svc.profiles()])
        # 更新
        r2 = svc.save_profile({"id": pid, "name": "方案A2", "templateId": tpl["id"], "rules": rules})
        self.assertEqual(r2["profile"]["name"], "方案A2")
        # 非法规则
        with self.assertRaises(ValueError):
            svc.save_profile({"templateId": tpl["id"], "rules": [{"mode": "nope"}]})
        # 非法模板
        with self.assertRaises(ValueError):
            svc.save_profile({"templateId": "tpl-nope", "rules": []})
        # 删除
        svc.delete_profile(pid)
        self.profile_ids.remove(pid)
        with self.assertRaises(ValueError):
            svc.delete_profile(pid)

    def test_suggest_by_ids(self):
        tpl = self._add(self.sample_copy, "推荐模板")
        res = svc.suggest_by_ids({"filePath": self.sample_copy, "sheetIndex": 0, "templateId": tpl["id"]})
        self.assertEqual(len(res["suggestions"]), 58)
        self.assertTrue(res["sourceSignature"])
        by_target = {s["targetField"]: s for s in res["suggestions"]}
        self.assertEqual(by_target["序号"]["sourceField"], "序号")
        self.assertGreaterEqual(by_target["序号"]["confidence"], 0.9)
        self.assertEqual(len(res["templateFields"]), 58)
        self.assertGreaterEqual(res["recordCount"], 1)
        # 再次推荐命中同一表头指纹 → matchedProfile
        svc.save_profile({"name": "指纹方案", "templateId": tpl["id"],
                          "rules": [], "sourceSignature": res["sourceSignature"]})
        res2 = svc.suggest_by_ids({"filePath": self.sample_copy, "sheetIndex": 0, "templateId": tpl["id"]})
        self.assertIsNotNone(res2["matchedProfile"])
        svc.delete_profile(res2["matchedProfile"]["id"]) if res2["matchedProfile"]["id"] in [
            p["id"] for p in svc.profiles()] else None

    def test_preview_location_of_contract_code_by_header_name(self):
        """回归（Lead 集成用例）：header=合同编码 必须解析到 AF 列并取到源值。"""
        tpl = self._add(self.sample_copy, "叶子名回归模板")
        rules = [{"source": {"row": {"kind": "record", "value": 0},
                             "col": {"kind": "header", "value": "合同编码"}},
                  "target": {"row": {"kind": "record", "value": 0},
                             "col": {"kind": "header", "value": "合同编码"}},
                  "mode": "cell"}]
        res = svc.preview_by_ids({"filePath": self.sample_copy, "sheetIndex": 0,
                                  "templateId": tpl["id"], "limit": 3, "rules": rules})
        first = res["cells"][0]
        self.assertEqual(first["sourceRef"]["a1"], "AF4")
        self.assertEqual(first["target"]["a1"], "AF4")
        self.assertEqual(first["value"], "DEMO-ZS-24001")

    def test_preview_temporary_rules(self):
        """Lead 补充要求：body.rules 未保存方案也能预演。"""
        tpl = self._add(self.simple_path, "临时规则模板")
        rules = [
            {"source": {"row": {"kind": "fixed", "value": 3},
                        "col": {"kind": "span", "value": {"start": 2, "count": 6}}},
             "target": {"row": {"kind": "span", "value": {"start": 2, "count": 6}},
                        "col": {"kind": "fixed", "value": 6}},
             "mode": "transpose"},
        ]
        res = svc.preview_by_ids(
            {"filePath": self.horizontal_path, "sheetIndex": 0, "templateId": tpl["id"], "limit": 50, "rules": rules}
        )
        self.assertEqual(res["profileId"], "")
        values = [c["value"] for c in res["cells"]]
        self.assertEqual(values, ["姓名", "张三", "性别", "男", "年龄", 30])
        self.assertEqual([c["target"]["a1"] for c in res["cells"]], ["F2", "F3", "F4", "F5", "F6", "F7"])
        # 来源坐标是 A 表第 3 行 B..G 列
        self.assertEqual([c["sourceRef"]["a1"] for c in res["cells"]], ["B3", "C3", "D3", "E3", "F3", "G3"])

    def test_preview_saved_profile(self):
        tpl = self._add(self.simple_path, "保存方案模板")
        rules = [{"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "姓名"}},
                  "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 1}},
                  "mode": "cell"}]
        pid = svc.save_profile({"name": "预演方案", "templateId": tpl["id"], "rules": rules})["profile"]["id"]
        self.profile_ids.append(pid)
        # 只传 profileId：templateId 由方案反查
        res = svc.preview_by_ids({"filePath": self.horizontal_path, "sheetIndex": 0, "profileId": pid, "limit": 5})
        self.assertEqual(res["profileId"], pid)
        self.assertEqual(res["cells"][0]["target"]["a1"], "A2")
        self.assertEqual(res["cells"][0]["value"], 1)
        with self.assertRaises(ValueError):
            svc.preview_by_ids({"filePath": self.horizontal_path, "sheetIndex": 0, "templateId": tpl["id"]})

    def test_convert_end_to_end_sample_template(self):
        """用样本台账模板生成 20 行台账：保样式 + 数据正确 + 可重新解析。"""
        tpl = self._add(self.sample_copy, "转换模板")
        rules = [
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "序号"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 1}}},
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "项目名称"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 2}}},
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "合同编码"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 32}}},
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "金额"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 29}},
             "transform": "number"},
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "实施部门"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 57}}},
        ]
        out_name = f"台账_测试_{uuid.uuid4().hex[:6]}.xlsx"
        task = _FakeTask()
        res = svc.start_convert(
            task,
            {"filePath": self.records_path, "sheetIndex": 0, "templateId": tpl["id"],
             "rules": rules, "outName": out_name},
        )
        self.outputs.append(res["path"])
        self.assertTrue(os.path.isfile(res["path"]))
        self.assertEqual(res["count"], 20)
        self.assertEqual(task.progress, 100)
        self.assertEqual(task.result["path"], res["path"])

        # 重新解析：数据行 4..23，未映射的列留空，标题/表头原样
        g = template.read_grid(res["path"])
        self.assertEqual(g["dimension"], "A1:BF23")
        self.assertEqual(g["cells"].get((3, 0)), 1)
        self.assertEqual(g["cells"].get((3, 1)), "测试项目01")
        self.assertEqual(g["cells"].get((3, 28)), 10000)
        self.assertEqual(g["cells"].get((3, 31)), "HT-2024-001")
        self.assertEqual(g["cells"].get((3, 56)), "工程部")
        self.assertEqual(g["cells"].get((22, 1)), "测试项目20")
        self.assertIsNone(g["cells"].get((3, 0 + 2)))  # C4 未映射 → 不含模板旧数据
        # 标题/表头与模板字节一致
        with zipfile.ZipFile(self.sample_copy) as a, zipfile.ZipFile(res["path"]) as b:
            old = a.read("xl/worksheets/sheet1.xml").decode("utf-8")
            new = b.read("xl/worksheets/sheet1.xml").decode("utf-8")
            for r in (1, 2, 3):
                o = re.search(r'<row r="%d"[^>]*>.*?</row>' % r, old, re.S).group(0)
                n = re.search(r'<row r="%d"[^>]*>.*?</row>' % r, new, re.S).group(0)
                self.assertEqual(o, n)

    def test_convert_by_saved_profile(self):
        tpl = self._add(self.simple_path, "方案转换模板")
        rules = [{"source": {"row": {"kind": "fixed", "value": 3},
                             "col": {"kind": "span", "value": {"start": 2, "count": 6}}},
                  "target": {"row": {"kind": "span", "value": {"start": 2, "count": 6}},
                             "col": {"kind": "fixed", "value": 6}},
                  "mode": "transpose"}]
        pid = svc.save_profile({"name": "转置方案", "templateId": tpl["id"], "rules": rules})["profile"]["id"]
        self.profile_ids.append(pid)
        out_name = f"转置_测试_{uuid.uuid4().hex[:6]}.xlsx"
        res = svc.start_convert(None, {"filePath": self.horizontal_path, "sheetIndex": 0,
                                       "profileId": pid, "outName": out_name})
        self.outputs.append(res["path"])
        self.assertEqual(res["count"], 6)
        g = template.read_grid(res["path"])
        self.assertEqual([g["cells"].get((r, 5)) for r in range(1, 7)],
                         ["姓名", "张三", "性别", "男", "年龄", 30])

    def test_convert_multiple_sources_append(self):
        """api 契约 body.sources[]：多个来源按追加语义写入同一模板。"""
        tpl = self._add(self.sample_copy, "多来源模板")
        rules = [
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "序号"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 1}}},
            {"source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "合同编码"}},
             "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 32}}},
        ]
        out_name = f"批量_测试_{uuid.uuid4().hex[:6]}.xlsx"
        res = svc.start_convert(None, {
            "sources": [
                {"filePath": self.records_path, "sheetIndex": 0},
                {"filePath": self.records_path, "sheetIndex": 0},
            ],
            "templateId": tpl["id"], "rules": rules, "outName": out_name,
        })
        self.outputs.append(res["path"])
        self.assertEqual(res["count"], 40)
        self.assertEqual(res["records"], 40)
        self.assertEqual(res["sourceCount"], 2)
        g = template.read_grid(res["path"])
        self.assertEqual(g["cells"].get((3, 0)), 1)
        self.assertEqual(g["cells"].get((22, 0)), 20)
        self.assertEqual(g["cells"].get((23, 0)), 1)  # 第二个来源从第 24 行继续
        self.assertEqual(g["cells"].get((23, 31)), "HT-2024-001")

    def test_convert_bad_requests(self):
        with self.assertRaises(ValueError):
            svc.start_convert(None, {"filePath": self.records_path})
        tpl = self._add(self.simple_path, "坏参数模板")
        with self.assertRaises(ValueError):
            svc.start_convert(None, {"filePath": os.path.join(TMP, "不存在.xlsx"), "templateId": tpl["id"],
                                     "rules": []})


class _FakeTask:
    def __init__(self):
        self.progress = 0
        self.message = ""
        self.result = None


if __name__ == "__main__":
    unittest.main()
