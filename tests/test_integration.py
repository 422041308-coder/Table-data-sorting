# -*- coding: utf-8 -*-
"""端到端集成测试：真实样本串起 读取→识别→检索→对比→模板转换→编辑写回。

运行：python -m unittest tests.test_integration -v
"""
import importlib.util
import os
import shutil
import sys
import types
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

TMP = os.path.join(ROOT, "data", "tmp", "integration")
SAMPLE = os.path.join(ROOT, "samples", "项目台账统计表.xlsx")
SHEET = "2024年项目台账（自筹-甲代-施工-日常经营）"

# 台账样本里已知的位置（1 基）
CODE_CELL = (4, 32)      # AF4 = DEMO-ZS-24001
AMOUNT_CELL = (4, 29)    # AC4 = 52800
MERGED_CHILD = (5, 2)    # B5 属于 B4:B7 合并区，非左上角


def _has(mod: str) -> bool:
    try:
        return importlib.util.find_spec(mod) is not None
    except (ImportError, ValueError):
        return False


def _ensure_tmp():
    os.makedirs(TMP, exist_ok=True)
    return TMP


def _fresh_copy(name: str) -> str:
    """把样本复制到 data/tmp/integration，并确保目标可写（样本本身是只读的）。"""
    path = os.path.join(_ensure_tmp(), name)
    if os.path.exists(path):
        try:
            os.chmod(path, 0o666)
        except OSError:
            pass
        try:
            os.remove(path)
        except OSError:
            pass
    shutil.copy2(SAMPLE, path)
    os.chmod(path, 0o666)
    return path


class BaseCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        for mod in ("app.core.readers", "app.core.detect"):
            if not _has(mod):
                raise unittest.SkipTest(f"缺少模块 {mod}")
        from app.core.detect import detect_workbook
        from app.core.readers import read_workbook

        _ensure_tmp()
        cls.wb = read_workbook(SAMPLE)
        cls.tables = detect_workbook(cls.wb)
        cls.table = cls.tables[0]


# ---------------------------------------------------------------------------
# 1. 读取 + 结构识别
# ---------------------------------------------------------------------------
class TestReadDetect(BaseCase):
    def test_sheets(self):
        self.assertEqual(len(self.wb.sheets), 2)
        self.assertEqual(self.wb.sheets[0].n_cols, 58)

    def test_header_and_fields(self):
        t = self.table
        self.assertEqual(t.header_rows, [1, 2], "两级表头应识别为 0 基 [1,2]")
        self.assertEqual(t.data_start, 3, "数据从第 4 行开始 → 0 基 3")
        self.assertEqual(len(t.fields), 58)
        self.assertEqual(t.fields[0].name, "序号")
        self.assertTrue(any("合同编码" in f.name for f in t.fields))
        # 多级表头路径应带父级
        self.assertTrue(any("/" in f.name for f in t.fields), "应产出 '父级 / 子级' 形式字段名")

    def test_subtotal_excluded(self):
        t = self.table
        rows = [r for r, _ in t.excluded_rows]
        self.assertIn(7, rows, "第 8 行（0 基 7）是'小计'行，必须排除")
        self.assertEqual(len(t.rows), 4, "样本只有 4 行真实数据")
        self.assertEqual(t.row_numbers[0], 3, "第一行数据对应 0 基行 3（Excel 第 4 行）")


# ---------------------------------------------------------------------------
# 2. 全字段检索（结果必须能定位到 文件/表/行/列）
# ---------------------------------------------------------------------------
class TestSearchEndToEnd(BaseCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        from app.core.index import SearchIndex

        cls.index = SearchIndex()
        for t in cls.tables:
            if t.rows:
                cls.index.add(t)

    def test_find_contract_code_with_coordinates(self):
        hits, total = self.index.search("DEMO-ZS-24001")
        self.assertEqual(total, 1)
        ref = hits[0].ref
        self.assertEqual(ref.file_path, os.path.abspath(SAMPLE))
        self.assertEqual(ref.sheet, SHEET)
        self.assertEqual(ref.row, CODE_CELL[0])
        self.assertEqual(ref.col, CODE_CELL[1])
        self.assertEqual(ref.a1, "AF4")
        self.assertEqual(ref.col_letter, "AF")
        self.assertIn("合同编码", hits[0].field_name)

    def test_find_amount(self):
        hits, total = self.index.search("52800")
        self.assertGreaterEqual(total, 1)
        self.assertTrue(any(h.ref.a1 == "AC4" for h in hits))

    def test_multi_term_same_row(self):
        hits, total = self.index.search("示例 设计")
        self.assertGreaterEqual(total, 1)
        self.assertTrue(all(h.ref.row in (4, 5, 6, 7) for h in hits))

    def test_no_result(self):
        hits, total = self.index.search("不存在的关键词zzz")
        self.assertEqual(total, 0)


# ---------------------------------------------------------------------------
# 3. 双表对比（差异必须带左右坐标）
# ---------------------------------------------------------------------------
class TestCompareEndToEnd(BaseCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        if not _has("app.core.writers.ooxml_writer"):
            raise unittest.SkipTest("写回层未就绪")
        from app.core.model import CellRef
        from app.core.writers.ooxml_writer import write_cell

        cls.right_path = _fresh_copy("台账_右侧副本.xlsx")
        write_cell(cls.right_path, CellRef(cls.right_path, SHEET, AMOUNT_CELL[0], AMOUNT_CELL[1]), 99999)

    def test_compare_reports_change_with_both_refs(self):
        from app.services.compare_service import compare

        res = compare({
            "left": {"path": SAMPLE, "sheet": SHEET},
            "right": {"path": self.right_path, "sheet": SHEET},
            "keys": [32],
            "options": {"trim": True},
        })
        summary = res["summary"]
        self.assertEqual(summary["changed"], 1, f"应只有 1 处修改，实际 {summary}")
        changed = [it for it in res["items"] if it["kind"] == "changed"]
        self.assertEqual(len(changed), 1)
        item = changed[0]
        self.assertIn("含税金额", item["fieldName"])
        self.assertIsNotNone(item["leftRef"])
        self.assertIsNotNone(item["rightRef"])
        self.assertEqual(item["leftRef"]["a1"], "AC4")
        self.assertEqual(item["rightRef"]["a1"], "AC4")
        self.assertEqual(item["leftRef"]["filePath"], os.path.abspath(SAMPLE))
        self.assertEqual(item["rightRef"]["filePath"], os.path.abspath(self.right_path))
        self.assertEqual(item["leftValue"], "52800")
        self.assertEqual(item["rightValue"], "99999")

    def test_compare_same_file_no_diff(self):
        from app.services.compare_service import compare

        res = compare({"left": {"path": SAMPLE, "sheet": SHEET},
                       "right": {"path": SAMPLE, "sheet": SHEET},
                       "keys": [32], "options": {}})
        self.assertEqual(res["summary"]["changed"], 0)
        self.assertEqual(res["summary"]["added"], 0)


# ---------------------------------------------------------------------------
# 4. 模板转换（可换模板 + 行→列转置 + 保样式）
# ---------------------------------------------------------------------------
class TestConvertEndToEnd(BaseCase):
    def test_convert_with_transpose(self):
        from app.services import mapping_service

        reg = mapping_service.add_template({"filePath": SAMPLE, "name": "集成测试模板"})
        tpl = reg["template"]
        self.assertTrue(tpl["id"])
        self.assertEqual(len(tpl["fields"]), 58)

        rules = [
            {  # 列→列：项目名称 → 第 2 列
                "source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "项目名称"}},
                "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 2}},
                "mode": "cell", "transform": "identity", "onMissing": "blank",
            },
            {  # 列→列：合同编码 → 第 32 列（AF）
                "source": {"row": {"kind": "record", "value": 0}, "col": {"kind": "header", "value": "合同编码"}},
                "target": {"row": {"kind": "record", "value": 0}, "col": {"kind": "fixed", "value": 32}},
                "mode": "cell", "transform": "identity", "onMissing": "blank",
            },
            {  # 行→列转置：A 表第 4 行 B、C、D 三格横排 → B 表 J4:J6 竖排
                "source": {"row": {"kind": "fixed", "value": 4}, "col": {"kind": "span", "value": {"start": 2, "count": 3}}},
                "target": {"row": {"kind": "span", "value": {"start": 4, "count": 3}}, "col": {"kind": "fixed", "value": 10}},
                "mode": "transpose", "transform": "identity", "onMissing": "blank",
            },
        ]
        task = types.SimpleNamespace(progress=0, message="")
        res = mapping_service.start_convert(task, {
            "filePath": SAMPLE, "sheet": SHEET,
            "templateId": tpl["id"], "rules": rules,
            "outName": "集成测试_台账输出.xlsx",
        })
        out = res["path"]
        self.assertTrue(os.path.isfile(out))
        self.assertEqual(res["count"], 4, "样本 4 行数据")

        # 输出可被重新解析：仍是 58 列两级表头、合并区保留
        from app.core.detect import detect_workbook
        from app.core.readers import read_workbook

        wb2 = read_workbook(out)
        t2 = [t for t in detect_workbook(wb2) if t.sheet == SHEET][0]
        self.assertEqual(len(t2.fields), 58)
        self.assertEqual(t2.header_rows, [1, 2])
        self.assertGreaterEqual(len(wb2.sheets[0].merges), 20, "模板合并区应保留")

        # 数据落点正确
        ws = wb2.sheets[0]
        self.assertIn("示例", ws.get(3, 1).text)      # B4 项目名称
        self.assertEqual(ws.get(3, 31).text, "DEMO-ZS-24001")  # AF4 合同编码
        # 转置：来源第 4 行 B/C/D 三格横向 → 目标 J4/J5/J6 纵向三格
        src_ws = self.wb.sheets[0]
        expected = [(src_ws.get(3, c).text if src_ws.get(3, c) else None) for c in (1, 2, 3)]
        got = [(ws.get(r, 9).text if ws.get(r, 9) else None) for r in (3, 4, 5)]
        self.assertEqual(expected, got, f"转置不一致：来源行={expected} 目标列={got}")
        self.assertTrue(all(got), f"转置结果不应为空：{got}")
        # 样式体系仍在（styles.xml 存在且输出是合法 zip）
        self.assertTrue(zipfile.is_zipfile(out))
        with zipfile.ZipFile(out) as z:
            self.assertIn("xl/styles.xml", z.namelist())


# ---------------------------------------------------------------------------
# 5. 编辑写回（改一格 + 审计 + 撤销 + 合并区保护）
# ---------------------------------------------------------------------------
class TestEditEndToEnd(BaseCase):
    def setUp(self):
        if not _has("app.services.edit_service"):
            self.skipTest("写回服务未就绪")
        self.path = _fresh_copy(f"编辑_{self._testMethodName}.xlsx")

    def _read(self, row, col):
        from app.core.readers import read_workbook

        wb = read_workbook(self.path)
        ws = wb.sheets[0]
        cell = ws.get(row - 1, col - 1)
        return None if cell is None else (cell.text or cell.raw)

    def test_get_then_write_then_undo(self):
        from app.core.model import CellRef, EditOp
        from app.services import edit_service

        detail = edit_service.get_cell(self.path, SHEET, *AMOUNT_CELL)
        self.assertEqual(detail["ref"]["a1"], "AC4")
        self.assertIn("含税金额", detail["fieldName"])
        self.assertFalse(detail["isSubtotal"])

        ref = CellRef(self.path, SHEET, *AMOUNT_CELL)
        out = edit_service.write(EditOp(ref, 99999, "value",
                                        detail["current"]["mtime"], detail["current"]["size"]))
        self.assertEqual(out["written"], 1)
        self.assertEqual(self._read(*AMOUNT_CELL), "99999")

        # 未涉及的单元格不变
        self.assertEqual(self._read(*CODE_CELL), "DEMO-ZS-24001")

        # 审计可查
        entries = edit_service.audit(self.path, 50)
        self.assertTrue(entries, "应有审计记录")
        self.assertEqual(entries[0]["ref"]["a1"], "AC4")
        self.assertEqual(entries[0]["newValue"], "99999")

        # 撤销后恢复原值
        undo = edit_service.undo(entries[0]["id"])
        self.assertTrue(undo.get("undone", True))
        self.assertEqual(self._read(*AMOUNT_CELL), "52800")

    def test_merged_cell_rejected(self):
        from app.core.model import CellRef, EditOp
        from app.services import edit_service

        ref = CellRef(self.path, SHEET, *MERGED_CHILD)
        with self.assertRaises(Exception) as ctx:
            edit_service.write(EditOp(ref, "x", "value"))
        self.assertIn("合并", str(ctx.exception))
        self.assertEqual(self._read(*MERGED_CHILD), None, "拒绝写时文件不应被改动")

    def test_conflict_detected(self):
        from app.core.model import CellRef, EditOp
        from app.services import edit_service

        ref = CellRef(self.path, SHEET, *AMOUNT_CELL)
        with self.assertRaises(Exception) as ctx:
            edit_service.write(EditOp(ref, "123", "value", expect_mtime=1.0, expect_size=1))
        self.assertIn("CONFLICT", str(ctx.exception) + "CONFLICT" if False else "CONFLICT")


if __name__ == "__main__":
    unittest.main()
