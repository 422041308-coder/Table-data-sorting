# -*- coding: utf-8 -*-
"""S2 结构识别单测：标题行 / 多级合并表头 / 数据区 / 小计·占位·空行排除。

fixture 全部自造 ``SheetModel``（不依赖 S1 的 ``read_workbook``）；
台账 fixture 按 ``samples/structure.txt`` 的 58 列 × 2 行表头布局构造。
"""
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
    FileRef,
    SheetModel,
    WorkbookModel,
)
from app.core.normalize import format_number, to_date_text  # noqa: E402
from app.core.detect import detect_table, detect_workbook  # noqa: E402


# --------------------------------------------------------------------------
# fixture 工具
# --------------------------------------------------------------------------
def _cell(value):
    """python 值 → Cell。"""
    if isinstance(value, Cell):
        return value
    if value is None:
        return Cell(kind=KIND_BLANK)
    if isinstance(value, bool):
        return Cell(raw=value, kind=KIND_BOOL, text="TRUE" if value else "FALSE")
    if isinstance(value, (dt.date, dt.datetime)):
        return Cell(raw=value, kind=KIND_DATE, text=to_date_text(value))
    if isinstance(value, (int, float)):
        return Cell(raw=value, kind=KIND_NUMBER, text=format_number(value), num=float(value))
    return Cell(raw=value, kind=KIND_TEXT, text=str(value))


def _formula(expr, cached=None):
    if cached is None:
        return Cell(raw=expr, kind=KIND_FORMULA, text=expr, formula=expr)
    return Cell(raw=cached, kind=KIND_FORMULA, text=format_number(cached), num=float(cached), formula=expr)


def _sheet(name, cells, merges=(), n_rows=None, n_cols=None):
    built = {}
    for key, value in cells.items():
        cell = _cell(value)
        if cell is not None:
            built[(int(key[0]), int(key[1]))] = cell
    max_r = max((r for r, _ in built), default=-1) + 1
    max_c = max((c for _, c in built), default=-1) + 1
    return SheetModel(
        name=name,
        n_rows=n_rows if n_rows is not None else max_r,
        n_cols=n_cols if n_cols is not None else max_c,
        cells=built,
        merges=[tuple(int(x) for x in m) for m in merges],
    )


# --------------------------------------------------------------------------
# 台账 fixture：58 列，row2 父级 + row3 子级（与 samples/structure.txt 一致）
# --------------------------------------------------------------------------
# (父级文本, 子级文本)；None 表示被左侧/上方的合并覆盖
LEDGER_COLS = [
    ("序号", None),
    ("项目名称", None),
    ("资金来源（自筹/代建/施工/日常经营）", None),
    ("合同名称", None),
    ("委托单位/甲方（插入评论合同编码）", None),
    ("政府投资总额/代建管理费/施工总承包费（元）", None),
    ("代建/甲方合同上传", None),
    ("计划部前期相关资料", "项目决策部门"),
    (None, "项目决策时间"),
    (None, "决策文件上传（会议纪要、决策部门文件、国资委等文件）"),
    (None, "可研时间"),
    (None, "可研（初设）上传资料"),
    (None, "项目立项文件上传（备案）"),
    (None, "代理公司"),
    (None, "代理合同时间"),
    (None, "代理协议资料上传"),
    (None, "招标控制价编制单位"),
    (None, "招标控制价编制时间"),
    (None, "招标控制价编制单位合同上传"),
    (None, "招标控制价（元）"),
    (None, "招标控制价文件上传（工程量清单+控制价）"),
    (None, "招标方式"),
    (None, "招标类型/标段（设计/施工/监理/审计）"),
    ("乙方中标单位", "中标单位名称"),
    (None, "中标时间"),
    (None, "中标通知书资料上传"),
    (None, "负责人"),
    (None, "联系电话"),
    ("乙方中标金额", "含税金额（元）"),
    (None, "税率\n（%）"),
    ("合同签订情况", "签订状态"),
    (None, "合同编码"),
    (None, "签订时间"),
    (None, "计划开工时间"),
    (None, "计划竣工时间"),
    (None, "合同扫描件上传（合同）无资料插入说明"),
    (None, "经办人"),
    (None, "存放位置"),
    ("审计报告", "进度小计"),
    (None, "最终"),
    (None, "审计报告文件上传（只上传前几页）"),
    ("乙方合同结算单", "中期小计"),
    (None, "最终"),
    (None, "质保金、农民工保证民返还"),
    (None, "结算资料上传"),
    ("乙方已付款金额", "金额小计"),
    (None, "资料转交财务时间"),
    ("未付款金额（元）", "（有审计小计未付款=审计小计-已付款)"),
    ("未付款金额（元）", "（有审计小计未付款=合同金额-已付款)"),
    ("乙方发票情况", "发票金额"),
    (None, "发票上传"),
    ("项目资料情况", "实际开工时间"),
    (None, "实际竣工时间"),
    (None, "开竣工报告上传"),
    (None, "项目资料是否齐全及资料存放位置）"),
    ("公司项目负责人", None),
    ("实施部门", None),
    ("备注", None),
]

_LEDGER_DATA = [
    {
        0: 1, 1: "示例项目甲—园区改造工程", 2: "企业自筹", 3: "设计合同",
        19: 55000, 21: "询价", 22: "设计类",
        23: "示例设计院", 24: "2024.3.9", 26: "孙八", 27: "13800000001",
        28: 52800, 29: 0.06, 30: "已签订", 31: "DEMO-ZS-24001", 32: "2024.3.15",
        36: "张三", 37: "2024自筹甲代", 45: 52800, 46: "2024.7.26李四",
        47: _formula("=IFS(...)", 0), 56: "工程部",
    },
    {
        3: "施工合同", 22: "施工类", 23: "示例建筑公司", 26: "周九",
        27: "13800000002", 28: 2788335, 29: 0.09, 30: "已签订",
        31: "DEMO-JS-24002", 32: "2024.3.26", 36: "张三",
        39: 2495529, 42: _formula("=2495529*0.9", 2245976.1), 45: 0,
        47: _formula("=IFS(...)", 2495529), 56: "计划统筹部",
    },
    {
        3: "监理合同", 22: "监理类", 23: "示例咨询公司", 26: "吴十",
        27: 13800000003, 28: 27883, 29: 0.06, 30: "已签订",
        31: "DEMO-JL-24003", 32: "2024.3.26", 36: "王五", 45: 27883,
        47: _formula("=IFS(...)", 0), 56: "工程部",
    },
    {
        3: "造价咨询合同", 22: "审计类", 23: "示例审计公司", 26: "钱七",
        27: 13800000004, 28: 16730, 29: 0.01, 30: "已签订",
        31: "DEMO-ZX-24004", 32: "2024.3.27", 36: "赵六", 45: 0,
        47: _formula("=IFS(...)", 16730), 56: "工程部",
    },
]


def _ledger_sheet():
    cells = {(0, 0): "2024年示例区示例公司项目台账（单位：元）"}
    merges = [(0, 0, 0, 57)]

    for c, (parent, _child) in enumerate(LEDGER_COLS):
        if parent is not None:
            cells[(1, c)] = parent

    # 跨列父级 → 水平合并
    c = 0
    while c < len(LEDGER_COLS):
        if LEDGER_COLS[c][0] is not None:
            end = c
            while end + 1 < len(LEDGER_COLS) and LEDGER_COLS[end + 1][0] is None:
                end += 1
            if end > c:
                merges.append((1, 1, c, end))
            c = end + 1
        else:
            c += 1

    for c, (_parent, child) in enumerate(LEDGER_COLS):
        if child is not None:
            cells[(2, c)] = child
    for c, (parent, child) in enumerate(LEDGER_COLS):
        if parent is not None and child is None:
            merges.append((1, 2, c, c))  # 父级纵向合并到子行

    # 数据行 A/B/C/AL 纵向合并（Excel 4-7 行）
    merges += [(3, 6, 0, 0), (3, 6, 1, 1), (3, 6, 2, 2), (3, 6, 37, 37)]
    for i, row in enumerate(_LEDGER_DATA):
        for c, value in row.items():
            cells[(3 + i, c)] = value

    # 小计行（Excel 第 8 行，0 基 7）
    subtotal = {1: "小计"}
    for col in (28, 38, 39, 41, 42, 43, 45, 47):
        subtotal[col] = _formula("=SUM(...)", 0)
    for c, value in subtotal.items():
        cells[(7, c)] = value

    # 分类占位行（Excel 第 9-11 行，0 基 8-10）
    cells[(8, 1)] = "乙南侧环境整治工程项目"
    cells[(8, 2)] = "甲方代建"
    cells[(9, 1)] = "示例项目丙—土方清运"
    cells[(9, 2)] = "施工类"
    cells[(10, 2)] = "日常经营"

    # 小计之后继续有数据（Excel 12-13 行）
    cells[(11, 0)] = 2
    cells[(11, 1)] = "项目B"
    cells[(11, 2)] = "施工类"
    cells[(11, 28)] = 1000
    cells[(11, 31)] = "DEMO-JS-24009"
    cells[(11, 45)] = 1000
    cells[(11, 56)] = "工程部"
    cells[(12, 0)] = 3
    cells[(12, 1)] = "项目C"
    cells[(12, 2)] = "日常经营"
    cells[(12, 28)] = 2000
    cells[(12, 31)] = "DEMO-ZS-24010"
    cells[(12, 45)] = 2000
    cells[(12, 56)] = "计划统筹部"

    return _sheet("2024年项目台账（自筹-甲代-施工-日常经营）", cells, merges, n_rows=63, n_cols=58)


# --------------------------------------------------------------------------
# 台账：结构验收
# --------------------------------------------------------------------------
class TestLedgerStructure(unittest.TestCase):
    def setUp(self):
        self.sheet = _ledger_sheet()
        self.wb = WorkbookModel(path=r"D:\样本\项目台账统计表.xlsx", fmt="xlsx", sheets=[self.sheet])
        self.table = detect_table(self.wb, self.sheet)

    def test_field_count_58(self):
        self.assertEqual(len(LEDGER_COLS), 58)
        self.assertEqual(len(self.table.fields), 58)
        self.assertEqual([f.index for f in self.table.fields], list(range(58)))

    def test_header_rows_zero_based(self):
        self.assertEqual(self.table.header_rows, [1, 2])

    def test_data_range(self):
        self.assertEqual(self.table.data_start, 3)  # 0 基 = Excel 第 4 行
        self.assertEqual(self.table.data_end, 12)
        self.assertEqual(self.table.row_numbers, [3, 4, 5, 6, 11, 12])
        self.assertEqual(len(self.table.rows), 6)

    def test_subtotal_and_placeholders_excluded(self):
        rows = [r for r, _ in self.table.excluded_rows]
        self.assertEqual(sorted(rows), [7, 8, 9, 10])
        reasons = dict(self.table.excluded_rows)
        self.assertIn("小计", reasons[7])
        for r in (8, 9, 10):
            self.assertIn("占位", reasons[r])

    def test_rows_align_with_row_numbers(self):
        self.assertEqual(len(self.table.rows), len(self.table.row_numbers))
        self.assertTrue(any(c.kind == KIND_NUMBER for c in self.table.rows[0].values()))
        # 第 4-7 行 A/B/C/AL 是纵向合并，后续行只有锚点单元格
        self.assertIn(0, self.table.rows[0])
        self.assertNotIn(0, self.table.rows[1])
        self.assertNotIn(1, self.table.rows[1])

    def test_field_names_multilevel(self):
        fields = {f.index: f for f in self.table.fields}
        self.assertEqual(fields[0].name, "序号")
        self.assertEqual(fields[1].name, "项目名称")
        self.assertEqual(fields[23].name, "乙方中标单位 / 中标单位名称")
        self.assertEqual(fields[28].name, "乙方中标金额 / 含税金额（元）")
        self.assertEqual(fields[31].name, "合同签订情况 / 合同编码")
        self.assertEqual(fields[45].name, "乙方已付款金额 / 金额小计")
        self.assertEqual(fields[9].name, "计划部前期相关资料 / 决策文件上传（会议纪要、决策部门文件、国资委等文件）")
        self.assertEqual(fields[56].name, "实施部门")
        self.assertEqual(fields[57].name, "备注")

    def test_field_levels(self):
        fields = {f.index: f for f in self.table.fields}
        self.assertEqual(fields[0].level, 1)
        self.assertEqual(fields[23].level, 2)
        self.assertEqual(fields[28].level, 2)
        self.assertEqual(sum(1 for f in self.table.fields if f.level > 1), 48)
        self.assertEqual(fields[45].col_span, (45, 45))

    def test_kind_hints(self):
        fields = {f.index: f for f in self.table.fields}
        self.assertEqual(fields[0].kind_hint, "number")   # 序号
        self.assertEqual(fields[27].kind_hint, "phone")   # 联系电话（文本+数字混存）
        self.assertEqual(fields[31].kind_hint, "code")    # 合同编码
        self.assertEqual(fields[28].kind_hint, "money")   # 价格 / 金额（元）
        self.assertEqual(fields[45].kind_hint, "money")   # 已付款金额
        self.assertEqual(fields[29].kind_hint, "number")  # 税率（%）不是金额
        self.assertEqual(fields[24].kind_hint, "date")    # 中标时间

    def test_notes_document_decisions(self):
        text = "\n".join(self.table.parse_notes)
        self.assertIn("大标题行", text)
        self.assertIn("表头行", text)
        self.assertIn("数据区", text)
        self.assertIn("字段：58 列", text)

    def test_source_from_workbook(self):
        self.assertIsNotNone(self.table.source)
        self.assertEqual(self.table.source.path, self.wb.path)
        self.assertEqual(self.table.path, self.wb.path)
        self.assertEqual(self.table.sheet, self.sheet.name)


class TestFileRef(unittest.TestCase):
    def test_explicit_file_ref_used(self):
        sheet = _ledger_sheet()
        wb = WorkbookModel(path=r"D:\a\b.xlsx", sheets=[sheet])
        ref = FileRef(id="f1", path=r"D:\a\b.xlsx", name="b.xlsx")
        table = detect_table(wb, sheet, ref)
        self.assertIs(table.source, ref)
        self.assertEqual(table.path, r"D:\a\b.xlsx")


# --------------------------------------------------------------------------
# 工单类表：单行表头、无小计（真实语料示例预警.xls 的形状）
# --------------------------------------------------------------------------
ORDER_HEADER = ["工单编号", "工单标题", "所属部门", "状态", "截止办理时间", "类型", "联系电话", "问题描述"]


def _order_sheet():
    cells = {}
    for c, text in enumerate(ORDER_HEADER):
        cells[(0, c)] = text
    data = [
        ["DEMO-GD-24001", "施工噪音扰民问题（示例区住建局）", "示例区住建局", "事件待处理",
         "2026-11-08 13:01:00", "噪音", "15047870340", "新华大街与东影南路交汇处工地夜间施工"],
        ["DEMO-GD-24002", "墙皮脱落问题（示例区兴业路）", "示例区住建局", "事件待处理",
         "2026-11-08 14:31:00", "墙皮", "13754010002", "兴业路小学小区4号楼1单元"],
        ["WX9915010026110103878", "垃圾清运问题（示例区）", "示例区城管局", "投拆",
         "2026-11-09 09:00:00", "垃圾", "13674811944", "泰安小区垃圾堆放"],
        ["DEMO-GD-24004", "欠薪问题（示例区）", "示例区人社局", "事件待处理",
         "2026-11-09 09:00:00", "欠薪", "15754831562", "某公司拖欠工资"],
    ]
    for i, row in enumerate(data):
        for c, value in enumerate(row):
            cells[(1 + i, c)] = value
    return _sheet("Sheet0", cells, [])


class TestOrderTable(unittest.TestCase):
    def setUp(self):
        self.sheet = _order_sheet()
        self.wb = WorkbookModel(path=r"D:\语料\示例预警.xls", fmt="xls", sheets=[self.sheet])
        self.table = detect_table(self.wb, self.sheet)

    def test_single_header_row_no_subtotal(self):
        self.assertEqual(self.table.header_rows, [0])
        self.assertEqual(self.table.data_start, 1)
        self.assertEqual(self.table.excluded_rows, [])
        self.assertEqual(len(self.table.rows), 4)
        self.assertEqual(self.table.row_numbers, [1, 2, 3, 4])

    def test_fields(self):
        self.assertEqual([f.name for f in self.table.fields], ORDER_HEADER)
        self.assertTrue(all(f.level == 1 for f in self.table.fields))

    def test_kind_hints(self):
        fields = {f.index: f for f in self.table.fields}
        self.assertEqual(fields[0].kind_hint, "code")     # 工单编号
        self.assertEqual(fields[4].kind_hint, "date")     # 截止办理时间（文本日期）
        self.assertEqual(fields[6].kind_hint, "phone")    # 联系电话
        self.assertEqual(fields[1].kind_hint, "text")


# --------------------------------------------------------------------------
# 带大标题 + 汇总小表（真实语料 "按部门_扣分部门" 的形状）
# --------------------------------------------------------------------------
class TestTitleAndSmallPivot(unittest.TestCase):
    def setUp(self):
        cells = {
            (0, 0): "截至11月9日各单位扣分工作量",
            (1, 0): "扣分部门",
            (1, 1): "扣分工单",
            (1, 2): "逾期工单",
            (2, 0): "示例区住建局", (2, 1): 9, (2, 2): "-",
            (3, 0): "乌兰察布东路街道", (3, 1): 7, (3, 2): 1,
            (4, 0): "城市投资公司", (4, 1): 2, (4, 2): "-",
        }
        self.sheet = _sheet("按部门_扣分部门", cells, [(0, 0, 0, 2)])
        self.wb = WorkbookModel(path=r"D:\语料\x.xls", fmt="xls", sheets=[self.sheet])
        self.table = detect_table(self.wb, self.sheet)

    def test_title_not_header(self):
        text = "\n".join(self.table.parse_notes)
        self.assertIn("大标题行：第 1 行", text)
        self.assertEqual(self.table.header_rows, [1])
        self.assertEqual(self.table.data_start, 2)
        self.assertEqual(len(self.table.fields), 3)
        self.assertEqual(len(self.table.rows), 3)
        self.assertEqual(self.table.excluded_rows, [])


# --------------------------------------------------------------------------
# 空行 / 重复表头 / 三级表头 / 无表头纯数据 / 空表
# --------------------------------------------------------------------------
class TestBlankRowInsideData(unittest.TestCase):
    def test_blank_row_excluded(self):
        cells = {
            (0, 0): "编号", (0, 1): "名称", (0, 2): "数量", (0, 3): "备注",
            (1, 0): "A-1", (1, 1): "甲", (1, 2): 1, (1, 3): "ok",
            (3, 0): "A-2", (3, 1): "乙", (3, 2): 2, (3, 3): "ok",
        }
        sheet = _sheet("S", cells)
        table = detect_table(WorkbookModel(path="x"), sheet)
        self.assertEqual(table.row_numbers, [1, 3])
        self.assertEqual([r for r, _ in table.excluded_rows], [2])
        self.assertIn("空行", table.excluded_rows[0][1])


class TestRepeatHeaderRow(unittest.TestCase):
    def test_repeated_header_excluded(self):
        header = {"编号": 0, "名称": 1, "数量": 2, "备注": 3}
        cells = {}
        for c, text in enumerate(header):
            cells[(0, c)] = text
        cells.update({(1, 0): "A-1", (1, 1): "甲", (1, 2): 1, (1, 3): "ok"})
        cells.update({(2, 0): "A-2", (2, 1): "乙", (2, 2): 2, (2, 3): "ok"})
        for c, text in enumerate(header):  # 打印分页重复的表头
            cells[(3, c)] = text
        cells.update({(4, 0): "A-3", (4, 1): "丙", (4, 2): 3, (4, 3): "ok"})
        cells.update({(5, 0): "A-4", (5, 1): "丁", (5, 2): 4, (5, 3): "ok"})
        table = detect_table(WorkbookModel(path="x"), _sheet("S", cells))
        self.assertEqual([r for r, _ in table.excluded_rows], [3])
        self.assertIn("重复表头", table.excluded_rows[0][1])
        self.assertEqual(table.row_numbers, [1, 2, 4, 5])


class TestThreeLevelHeader(unittest.TestCase):
    def setUp(self):
        cells = {
            (0, 0): "序号", (0, 1): "基本信息", (0, 4): "备注",
            (1, 1): "人员", (1, 3): "金额",
            (2, 1): "姓名", (2, 2): "电话", (2, 3): "金额小计",
            (3, 0): 1, (3, 1): "张三", (3, 2): "13800000000", (3, 3): 100, (3, 4): "备注1",
            (4, 0): 2, (4, 1): "李四", (4, 2): "13900000000", (4, 3): 200, (4, 4): "备注2",
        }
        merges = [(0, 2, 0, 0), (0, 0, 1, 3), (0, 2, 4, 4), (1, 1, 1, 2)]
        self.table = detect_table(WorkbookModel(path="x"), _sheet("S", cells, merges))

    def test_header_rows(self):
        self.assertEqual(self.table.header_rows, [0, 1, 2])
        self.assertEqual(self.table.data_start, 3)

    def test_names(self):
        names = [f.name for f in self.table.fields]
        self.assertEqual(names, ["序号", "基本信息 / 人员 / 姓名", "基本信息 / 人员 / 电话", "基本信息 / 金额 / 金额小计", "备注"])
        levels = [f.level for f in self.table.fields]
        self.assertEqual(levels, [1, 3, 3, 3, 1])


class TestNoHeaderPureData(unittest.TestCase):
    def test_numeric_matrix_without_header(self):
        cells = {
            (0, 0): 1, (0, 1): 2, (0, 2): 3,
            (1, 0): 4, (1, 1): 5, (1, 2): 6,
        }
        table = detect_table(WorkbookModel(path="x"), _sheet("S", cells))
        self.assertEqual(table.header_rows, [])
        self.assertEqual(table.data_start, 0)
        self.assertEqual([f.name for f in table.fields], ["列A", "列B", "列C"])
        self.assertEqual(len(table.rows), 2)

    def test_text_dump_without_header(self):
        """系统导出的明细页常常没有表头：首行就是长文本数据，不能当表头吃掉。"""
        long_title = "市民反映施工噪音扰民问题希望相关部门尽快核实处理并在规定时限内反馈办理结果，市民表示如果不处理将继续向上级部门反映情况"
        long_body = "市民反映：2026年11月8日晚间位于新华大街与东影南路交汇处的施工工地持续夜间施工，大型机械噪音严重影响周边居民正常休息，诉求请相关职能部门尽快到现场核实处理并对施工时间进行规范"
        cells = {}
        for r in range(3):
            cells[(r, 0)] = f"DEMO-GD-240{r}"
            cells[(r, 1)] = long_title
            cells[(r, 2)] = long_body
        table = detect_table(WorkbookModel(path="x"), _sheet("近期办结", cells))
        self.assertEqual(table.header_rows, [])
        self.assertEqual(table.data_start, 0)
        self.assertEqual(len(table.rows), 3)
        self.assertEqual(table.row_numbers, [0, 1, 2])
        self.assertTrue(table.fields[0].name.startswith("列"))


class TestEmptySheets(unittest.TestCase):
    def test_empty_sheet_yields_empty_table(self):
        sheet = SheetModel(name="Sheet1")
        table = detect_table(WorkbookModel(path="x"), sheet)
        self.assertEqual(table.fields, [])
        self.assertEqual(table.rows, [])
        self.assertTrue(table.parse_notes)

    def test_detect_workbook_skips_empty(self):
        ledger = _ledger_sheet()
        empty = SheetModel(name="Sheet1")
        wb = WorkbookModel(path=r"D:\样本\项目台账统计表.xlsx", sheets=[ledger, empty])
        tables = detect_workbook(wb)
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0].sheet, ledger.name)
        self.assertEqual(len(tables[0].fields), 58)

    def test_single_cell_sheet_skipped(self):
        sheet = _sheet("S", {(0, 0): "备注：无"})
        wb = WorkbookModel(path="x", sheets=[sheet])
        self.assertEqual(detect_workbook(wb), [])


if __name__ == "__main__":
    unittest.main()
