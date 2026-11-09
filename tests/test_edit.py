# -*- coding: utf-8 -*-
"""S8 编辑写回单测：最小改动 / 字节不变性 / 合并区 / 冲突 / 回滚 / 撤销 / 批量 / 查找替换。

运行：
    python -m unittest tests.test_edit -v
或整套：
    python -m unittest discover -s tests -v

测试文件全部放在 ``data/tmp/``（工作区内），不使用系统 Temp。
"""
import hashlib
import os
import re
import shutil
import sys
import unittest
import uuid
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from app.core import audit as audit_mod  # noqa: E402
from app.core import xlsxwriter  # noqa: E402
from app.core.model import CellRef, EditOp  # noqa: E402
from app.core.writers import ooxml_writer as writer  # noqa: E402
from app.server import ApiError  # noqa: E402
from app.services import edit_service  # noqa: E402

TMP_BASE = os.path.join(ROOT, "data", "tmp", "edit_test")


def _sha256(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def _entries(path: str) -> dict:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _raw_cell(path: str, part: str, a1: str):
    with zipfile.ZipFile(path) as z:
        xml = z.read(part).decode("utf-8")
    m = re.search(r'<c\b[^>]*\br="' + re.escape(a1) + r'"[^>]*?(?:/>|>.*?</c>)', xml, re.DOTALL)
    return m.group(0) if m else None


def _backup_dirs() -> set:
    if not os.path.isdir(audit_mod.BACKUP_DIR):
        return set()
    return {
        os.path.join(audit_mod.BACKUP_DIR, n)
        for n in os.listdir(audit_mod.BACKUP_DIR)
        if os.path.isdir(os.path.join(audit_mod.BACKUP_DIR, n))
    }


def _write_entries(path: str, entries: dict) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)


def _convert_to_shared_strings(path: str) -> None:
    """把 xlsxwriter 生成的 inlineStr 单元格改成 sharedStrings 引用（模拟 Excel/WPS 真实文件）。"""
    from xml.sax.saxutils import escape as xesc
    from xml.sax.saxutils import unescape as xunesc

    with zipfile.ZipFile(path) as z:
        entries = {n: z.read(n) for n in z.namelist()}

    strings: list[str] = []
    index: dict[str, int] = {}
    count = 0

    def repl(m):
        nonlocal count
        attrs = (m.group(1) + m.group(2)).strip()
        text = xunesc(m.group(3))
        if text not in index:
            index[text] = len(strings)
            strings.append(text)
        count += 1
        return f'<c {attrs} t="s"><v>{index[text]}</v></c>'

    pattern = re.compile(
        r'<c\b([^>]*?)\bt="inlineStr"([^>]*)><is><t[^>]*>(.*?)</t></is></c>',
        re.DOTALL,
    )
    for name in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
        xml = entries[name].decode("utf-8")
        entries[name] = pattern.sub(repl, xml).encode("utf-8")

    si = "".join(f'<si><t xml:space="preserve">{xesc(s)}</t></si>' for s in strings)
    entries["xl/sharedStrings.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{count}" uniqueCount="{len(strings)}">{si}</sst>'
    ).encode("utf-8")

    ct = entries["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(
        "</Types>",
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
    )
    entries["[Content_Types].xml"] = ct.encode("utf-8")

    rels = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdSst" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/></Relationships>',
    )
    entries["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    _write_entries(path, entries)


def _build_fixture(path: str) -> None:
    """造含中文 / 数值 / 合并区 / 两个工作表的测试文件，并给工作表补 <dimension>。"""
    xlsxwriter.write_xlsx(path, [
        {
            "name": "台账",
            "header": ["项目", "金额", "说明", "备注"],
            "rows": [
                ["甲项目", 1234.5, "中文备注", "备注一"],
                ["乙项目", 2345, "第二行", "备注二"],
                ["丙项目", 3456.25, "第三行", "备注三"],
                ["合计", 7035.75, "", "备注四"],
            ],
            "merges": [(4, 4, 3, 4)],  # C4:D4 合并（1 基）
            "widths": [12, 10, 20, 10],
        },
        {
            "name": "第二表",
            "header": ["A", "B"],
            "rows": [["保持", 1], ["不变", 2]],
        },
    ])
    # xlsxwriter 不写 <dimension>；真实 Excel 会写。这里补上以覆盖 clear 时更新 dimension 的逻辑。
    with zipfile.ZipFile(path) as z:
        entries = {n: z.read(n) for n in z.namelist()}
    sheet1 = entries["xl/worksheets/sheet1.xml"].decode("utf-8")
    sheet1 = re.sub(r"(<worksheet\b[^>]*>)", r'\1<dimension ref="A1:D5"/>', sheet1, count=1)
    entries["xl/worksheets/sheet1.xml"] = sheet1.encode("utf-8")
    sheet2 = entries["xl/worksheets/sheet2.xml"].decode("utf-8")
    sheet2 = re.sub(r"(<worksheet\b[^>]*>)", r'\1<dimension ref="A1:B3"/>', sheet2, count=1)
    entries["xl/worksheets/sheet2.xml"] = sheet2.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                  "xl/_rels/workbook.xml.rels", "xl/styles.xml",
                  "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
            z.writestr(n, entries[n])


class EditTestBase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.makedirs(TMP_BASE, exist_ok=True)
        cls._backup_before = _backup_dirs()

    @classmethod
    def tearDownClass(cls):
        # 只清理本测试类新建的备份目录，不动已有备份
        for d in _backup_dirs() - cls._backup_before:
            shutil.rmtree(d, ignore_errors=True)

    def setUp(self):
        # 注意：不用 tempfile.mkdtemp（当前 DSH Windows 沙箱会给 mkdtemp 目录加
        # 拒绝访问 ACL，导致后续无法写入）；改为在 data/tmp 下自建唯一目录。
        self.work = os.path.join(TMP_BASE, "w_" + uuid.uuid4().hex[:10])
        os.makedirs(self.work, exist_ok=True)
        self.path = os.path.join(self.work, "台账.xlsx")
        _build_fixture(self.path)
        self.sheet1 = "xl/worksheets/sheet1.xml"
        self.sheet2 = "xl/worksheets/sheet2.xml"

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    # ---------------- helpers ----------------
    def ref(self, sheet: str, row: int, col: int) -> CellRef:
        return CellRef(self.path, sheet, row, col)

    def assert_only_part_changed(self, before: dict, after: dict, changed_parts: set):
        self.assertEqual(list(before.keys()), list(after.keys()), "ZIP 条目清单不应变化")
        for name in before:
            if name in changed_parts:
                self.assertNotEqual(before[name], after[name], f"目标条目应发生变化：{name}")
            else:
                self.assertEqual(before[name], after[name], f"非目标条目字节被改动：{name}")

    def cell_values(self) -> dict:
        """全部单元格语义值快照（用于验证其他单元格不变）。"""
        snap = {}
        for info in writer.scan_cells(self.path):
            snap[(info["sheet"], info["a1"])] = (info["kind"], info["value"])
        return snap


class TestGetCell(EditTestBase):
    def test_shape(self):
        d = edit_service.get_cell(self.path, "台账", 2, 3)
        for key in ("ref", "value", "kind", "formula", "display", "fieldName",
                    "isSubtotal", "mergeOwner", "editable", "current"):
            self.assertIn(key, d)
        self.assertEqual(d["value"], "中文备注")
        self.assertEqual(d["kind"], "text")
        self.assertIsNone(d["formula"])
        self.assertEqual(d["display"], "中文备注")
        self.assertTrue(d["editable"])
        self.assertIsInstance(d["fieldName"], str)
        self.assertFalse(d["isSubtotal"])
        st = os.stat(self.path)
        self.assertEqual(d["current"]["size"], st.st_size)
        self.assertAlmostEqual(d["current"]["mtime"], st.st_mtime, places=4)
        ref = d["ref"]
        for key in ("filePath", "sheet", "row", "col", "colLetter", "a1", "display"):
            self.assertIn(key, ref)
        self.assertEqual(ref["a1"], "C2")
        self.assertEqual(ref["sheet"], "台账")

    def test_number_and_blank(self):
        self.assertEqual(edit_service.get_cell(self.path, "台账", 2, 2)["value"], 1234.5)
        blank = edit_service.get_cell(self.path, "第二表", 9, 9)
        self.assertEqual(blank["kind"], "blank")
        self.assertIsNone(blank["value"])

    def test_missing_sheet(self):
        with self.assertRaises(ApiError) as cm:
            edit_service.get_cell(self.path, "不存在", 1, 1)
        self.assertEqual(cm.exception.code, "NOT_FOUND")


class TestMinimalWrite(EditTestBase):
    def test_write_text_and_byte_invariance(self):
        before = _entries(self.path)
        values_before = self.cell_values()

        res = writer.write_cell(self.path, self.ref("台账", 2, 3), "新的中文说明")
        self.assertEqual(res["result"], "ok")
        self.assertEqual(res["oldValue"], "中文备注")
        self.assertEqual(res["newValue"], "新的中文说明")
        self.assertTrue(os.path.isfile(res["backupPath"]), "应先备份再写回")
        self.assertTrue(os.path.isfile(self.path))

        after = _entries(self.path)
        # (c) 其他条目（styles.xml / 其他 sheet / workbook.xml）字节完全一致
        self.assert_only_part_changed(before, after, {self.sheet1})

        # (a) 目标格已是新值
        info = writer.read_cell(self.path, "台账", 2, 3)
        self.assertEqual(info["value"], "新的中文说明")
        self.assertEqual(info["kind"], "text")

        # (b) 其他单元格值不变
        values_after = self.cell_values()
        changed = {k for k in values_before if values_before[k] != values_after.get(k)}
        self.assertEqual(changed, {("台账", "C2")})
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 2)["value"], 1234.5)
        self.assertEqual(writer.read_cell(self.path, "第二表", 2, 2)["value"], 1.0)

        # XML 必须仍然可解析
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(self.path) as z:
            for n in z.namelist():
                if n.endswith((".xml", ".rels")):
                    ET.fromstring(z.read(n))

    def test_write_number_keeps_style(self):
        raw_before = _raw_cell(self.path, self.sheet1, "B2")
        self.assertIn('s="0"', raw_before)
        writer.write_cell(self.path, self.ref("台账", 2, 2), 999.25)
        raw_after = _raw_cell(self.path, self.sheet1, "B2")
        self.assertIn('s="0"', raw_after, "写回必须保留原单元格样式索引")
        self.assertIn("<v>999.25</v>", raw_after)
        self.assertNotIn('t="', raw_after)
        info = writer.read_cell(self.path, "台账", 2, 2)
        self.assertEqual(info["kind"], "number")
        self.assertEqual(info["value"], 999.25)

    def test_write_outside_dimension_expands(self):
        with zipfile.ZipFile(self.path) as z:
            xml = z.read(self.sheet1).decode("utf-8")
        self.assertIn('ref="A1:D5"', xml)
        writer.write_cell(self.path, self.ref("台账", 7, 6), "新区域")
        with zipfile.ZipFile(self.path) as z:
            xml = z.read(self.sheet1).decode("utf-8")
        self.assertIn('ref="A1:F7"', xml, "写入超出 dimension 时应扩展")
        self.assertEqual(writer.read_cell(self.path, "台账", 7, 6)["value"], "新区域")

    def test_write_cells_single_pass(self):
        before = _entries(self.path)
        n = writer.write_cells(self.path, [
            (self.ref("台账", 2, 1), "甲改", "value"),
            (self.ref("台账", 2, 2), 888.5, "value"),
        ])
        self.assertEqual(n, 2)
        after = _entries(self.path)
        self.assert_only_part_changed(before, after, {self.sheet1})
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 1)["value"], "甲改")
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 2)["value"], 888.5)


class TestFormulaAndClear(EditTestBase):
    def test_formula_then_cached_value(self):
        # kind="formula" → 写 <f>
        writer.write_cell(self.path, self.ref("台账", 6, 3), "=B2*2", kind="formula")
        raw = _raw_cell(self.path, self.sheet1, "C6")
        self.assertIn("<f>B2*2</f>", raw)
        info = writer.read_cell(self.path, "台账", 6, 3)
        self.assertEqual(info["kind"], "formula")
        self.assertEqual(info["formula"], "=B2*2")

        # kind="value" → 只改缓存 <v>，保留 <f>
        writer.write_cell(self.path, self.ref("台账", 6, 3), 1998.5, kind="value")
        raw = _raw_cell(self.path, self.sheet1, "C6")
        self.assertIn("<f>B2*2</f>", raw)
        self.assertIn("<v>1998.5</v>", raw)
        info = writer.read_cell(self.path, "台账", 6, 3)
        self.assertEqual(info["formula"], "=B2*2")
        self.assertEqual(info["value"], 1998.5)

    def test_formula_empty_rejected(self):
        with self.assertRaises(ValueError):
            writer.write_cell(self.path, self.ref("台账", 6, 3), "=", kind="formula")

    def test_clear_removes_cell_and_shrinks_dimension(self):
        before = _entries(self.path)
        # 清空第 5 行（A5/B5/C5(空串单元格)/D5），dimension 应收缩到 A1:D4
        n = writer.write_cells(self.path, [
            (self.ref("台账", 5, 1), None, "clear"),
            (self.ref("台账", 5, 2), None, "clear"),
            (self.ref("台账", 5, 3), None, "clear"),
            (self.ref("台账", 5, 4), None, "clear"),
        ])
        self.assertEqual(n, 4)
        after = _entries(self.path)
        self.assert_only_part_changed(before, after, {self.sheet1})
        for a1 in ("A5", "B5", "C5", "D5"):
            self.assertIsNone(_raw_cell(self.path, self.sheet1, a1), f"{a1} 节点应被删除")
            self.assertEqual(writer.read_cell(self.path, "台账",
                                              int(a1[1:]), ord(a1[0]) - 64)["kind"], "blank")
        with zipfile.ZipFile(self.path) as z:
            xml = z.read(self.sheet1).decode("utf-8")
        self.assertIn('<dimension ref="A1:D4"/>', xml, "clear 后应同步收缩 dimension")


class TestMerges(EditTestBase):
    def test_non_topleft_rejected(self):
        # C4:D4 合并，D4 不是左上角
        h = _sha256(self.path)
        with self.assertRaises(ValueError) as cm:
            writer.write_cell(self.path, self.ref("台账", 4, 4), "不允许")
        self.assertIn("合并区", str(cm.exception))
        self.assertIn("左上角", str(cm.exception))
        self.assertEqual(_sha256(self.path), h, "被拒绝时不得改动文件")

        with self.assertRaises(ValueError):
            edit_service.write(EditOp(self.ref("台账", 4, 4), "不允许", "value"))
        self.assertEqual(_sha256(self.path), h)

    def test_topleft_allowed_and_get_cell_meta(self):
        writer.write_cell(self.path, self.ref("台账", 4, 3), "合并左上角")
        self.assertEqual(writer.read_cell(self.path, "台账", 4, 3)["value"], "合并左上角")
        info = writer.read_cell(self.path, "台账", 4, 4)
        self.assertFalse(info["editable"])
        self.assertEqual(info["mergeOwner"], "C4")
        self.assertEqual(info["merge"], "C4:D4")
        # 合并定义必须保留
        with zipfile.ZipFile(self.path) as z:
            self.assertIn(b"mergeCell", z.read(self.sheet1))


class TestConflict(EditTestBase):
    def test_conflict_mtime(self):
        st = os.stat(self.path)
        h = _sha256(self.path)
        op = EditOp(self.ref("台账", 2, 2), 1, "value", st.st_mtime - 60, st.st_size)
        with self.assertRaises(ApiError) as cm:
            edit_service.write(op)
        self.assertEqual(cm.exception.code, "CONFLICT")
        self.assertEqual(cm.exception.status, 409)
        self.assertEqual(_sha256(self.path), h)

    def test_conflict_size(self):
        st = os.stat(self.path)
        h = _sha256(self.path)
        with self.assertRaises(ApiError) as cm:
            edit_service.write(EditOp(self.ref("台账", 2, 2), 1, "value",
                                      st.st_mtime, st.st_size + 1))
        self.assertEqual(cm.exception.code, "CONFLICT")
        self.assertEqual(_sha256(self.path), h)

    def test_correct_expects_ok(self):
        st = os.stat(self.path)
        res = edit_service.write(EditOp(self.ref("台账", 2, 2), 777, "value",
                                        st.st_mtime, st.st_size))
        self.assertEqual(res["result"], "ok")
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 2)["value"], 777.0)


class TestUndoAndAudit(EditTestBase):
    def test_undo_restores_file(self):
        h0 = _sha256(self.path)
        st = os.stat(self.path)
        res = edit_service.write(EditOp(self.ref("台账", 2, 3), "审计写回", "value",
                                        st.st_mtime, st.st_size))
        self.assertTrue(res["opId"])
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 3)["value"], "审计写回")
        self.assertNotEqual(_sha256(self.path), h0)

        undo_res = edit_service.undo(res["opId"])
        self.assertEqual(undo_res["result"], "ok")
        self.assertEqual(undo_res["undoOf"], res["opId"])
        self.assertEqual(_sha256(self.path), h0, "undo 必须按备份整文件还原")
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 3)["value"], "中文备注")

        # 审计里应有 undo 记录
        entries = edit_service.audit(self.path, 20)
        self.assertTrue(any(e.get("action") == "undo" for e in entries))
        # 撤销未知 op → ApiError
        with self.assertRaises(ApiError) as cm:
            edit_service.undo("no-such-op")
        self.assertEqual(cm.exception.code, "UNDO_FAILED")

    def test_audit_entries_shape(self):
        st = os.stat(self.path)
        res = edit_service.write(EditOp(self.ref("台账", 3, 1), "乙改", "value",
                                        st.st_mtime, st.st_size))
        entries = edit_service.audit(self.path, 50)
        rec = next(e for e in entries if e.get("id") == res["opId"])
        for key in ("ts", "ref", "oldValue", "newValue", "backupPath", "result", "filePath", "id"):
            self.assertIn(key, rec)
        self.assertEqual(rec["oldValue"], "乙项目")
        self.assertEqual(rec["newValue"], "乙改")
        self.assertEqual(rec["result"], "ok")
        self.assertTrue(os.path.isfile(rec["backupPath"]))
        for key in ("filePath", "sheet", "row", "col", "colLetter", "a1", "display"):
            self.assertIn(key, rec["ref"])


class TestBatch(EditTestBase):
    def test_batch_two_sheets_one_backup(self):
        before = _entries(self.path)
        backup_before = _backup_dirs()
        st = os.stat(self.path)
        res = edit_service.write_batch([
            EditOp(self.ref("台账", 2, 1), "甲改", "value", st.st_mtime, st.st_size),
            EditOp(self.ref("第二表", 2, 2), 99, "value", st.st_mtime, st.st_size),
        ])
        self.assertEqual(res["written"], 2)
        self.assertEqual(len(res["opIds"]), 2)
        self.assertEqual(len(res["results"]), 2)
        new_backups = _backup_dirs() - backup_before
        self.assertEqual(len(new_backups), 1, "同文件批量写回只应备份一次")
        backup_dir = os.path.abspath(next(iter(new_backups)))
        self.assertTrue(
            os.path.abspath(res["backupPath"]).startswith(backup_dir + os.sep),
            f"备份应落在本次新建的备份目录下：{res['backupPath']}",
        )

        after = _entries(self.path)
        self.assert_only_part_changed(before, after, {self.sheet1, self.sheet2})
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 1)["value"], "甲改")
        self.assertEqual(writer.read_cell(self.path, "第二表", 2, 2)["value"], 99.0)

        # 两条审计记录共用同一次备份
        entries = edit_service.audit(self.path, 50)
        recs = [e for e in entries if e.get("id") in set(res["opIds"])]
        self.assertEqual(len(recs), 2)
        self.assertEqual(len({r["backupPath"] for r in recs}), 1)

    def test_batch_empty_rejected(self):
        with self.assertRaises(ApiError):
            edit_service.write_batch([])


class TestRollback(EditTestBase):
    def test_selfcheck_failure_rolls_back(self):
        h = _sha256(self.path)
        original = writer._self_check

        def boom(*args, **kwargs):
            raise RuntimeError("injected self-check failure")

        writer._self_check = boom
        try:
            with self.assertRaises(writer.WriteError) as cm:
                writer.write_cells(self.path, [(self.ref("台账", 2, 2), 123, "value")])
        finally:
            writer._self_check = original
        self.assertEqual(cm.exception.result, "rolled_back")
        self.assertTrue(os.path.isfile(cm.exception.backup_path), "失败也要保留备份")
        self.assertEqual(_sha256(self.path), h, "自检失败必须保持原文件不变")
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 2)["value"], 1234.5)


class TestFormats(EditTestBase):
    def test_ole_rejected(self):
        bad = os.path.join(self.work, "旧文件.xls")
        with open(bad, "wb") as f:
            f.write(writer.OLE_MAGIC + b"0" * 128)
        with self.assertRaises(ApiError) as cm:
            writer.write_cell(bad, CellRef(bad, "Sheet1", 1, 1), "x")
        self.assertEqual(cm.exception.code, "UNSUPPORTED_FORMAT")
        self.assertIn("xlsx", cm.exception.message)
        with self.assertRaises(ApiError) as cm2:
            edit_service.get_cell(bad, "Sheet1", 1, 1)
        self.assertEqual(cm2.exception.code, "UNSUPPORTED_FORMAT")

        bad_et = os.path.join(self.work, "旧文件.et")
        with open(bad_et, "wb") as f:
            f.write(writer.OLE_MAGIC + b"0" * 128)
        with self.assertRaises(ApiError) as cm3:
            writer.write_cell(bad_et, CellRef(bad_et, "Sheet1", 1, 1), "x")
        self.assertEqual(cm3.exception.code, "UNSUPPORTED_FORMAT")

    def test_xlsm_keeps_vba_bytes(self):
        xlsm = os.path.join(self.work, "宏文件.xlsm")
        shutil.copyfile(self.path, xlsm)
        vba = b"FAKE-VBA-PROJECT-BYTES"
        with zipfile.ZipFile(xlsm, "a", zipfile.ZIP_DEFLATED) as z:
            z.writestr("xl/vbaProject.bin", vba)
        before = _entries(xlsm)
        res = writer.write_cell(xlsm, CellRef(xlsm, "台账", 2, 3), "宏文件内改动")
        self.assertEqual(res["result"], "ok")
        after = _entries(xlsm)
        self.assertEqual(after["xl/vbaProject.bin"], vba, "VBA 宏必须原样保留")
        for name in before:
            if name == "xl/worksheets/sheet1.xml":
                continue
            self.assertEqual(before[name], after[name], f"xlsm 非目标条目被改动：{name}")

    def test_et_zip_supported(self):
        # WPS ET 的 zip 变体：与 xlsx 同构，应可写回
        et = os.path.join(self.work, "表格.et")
        shutil.copyfile(self.path, et)
        res = writer.write_cell(et, CellRef(et, "台账", 2, 3), "ET写回")
        self.assertEqual(res["result"], "ok")
        self.assertEqual(writer.read_cell(et, "台账", 2, 3)["value"], "ET写回")


class TestReplace(EditTestBase):
    def test_preview_shape_and_truncate(self):
        pv = edit_service.replace_preview({
            "find": "备注", "replace": "说明", "filePath": self.path,
            "matchCase": False, "limit": 200,
        })
        for key in ("matches", "count", "truncated"):
            self.assertIn(key, pv)
        self.assertGreaterEqual(pv["count"], 3)
        self.assertFalse(pv["truncated"])
        self.assertEqual(pv["count"], len(pv["matches"]))
        for m in pv["matches"]:
            self.assertIn("ref", m)
            self.assertIn("oldValue", m)
            self.assertIn("newValue", m)
            self.assertIn("备注", m["oldValue"])
            self.assertIn("说明", m["newValue"])
            for key in ("filePath", "sheet", "row", "col", "colLetter", "a1"):
                self.assertIn(key, m["ref"])

        only1 = edit_service.replace_preview({
            "find": "备注", "replace": "说明", "filePath": self.path, "limit": 1,
        })
        self.assertEqual(only1["count"], 1)
        self.assertTrue(only1["truncated"])

        none = edit_service.replace_preview({
            "find": "不存在的词", "replace": "x", "filePath": self.path,
        })
        self.assertEqual(none["count"], 0)
        self.assertFalse(none["truncated"])

    def test_apply_contract_body(self):
        pv = edit_service.replace_preview({
            "find": "备注", "replace": "说明", "filePath": self.path,
        })
        items = [{"ref": m["ref"], "newValue": m["newValue"]} for m in pv["matches"]]
        before = _entries(self.path)
        st = os.stat(self.path)
        res = edit_service.replace_apply({"items": items,
                                          "expectMtime": st.st_mtime,
                                          "expectSize": st.st_size})
        self.assertEqual(res["written"], pv["count"])
        self.assertEqual(len(res["results"]), pv["count"])
        for item in res["results"]:
            self.assertEqual(item["result"], "ok")
            self.assertIn("ref", item)
            self.assertIn("opId", item)
        after = _entries(self.path)
        self.assertTrue(res["backupPath"] and os.path.isfile(res["backupPath"]))
        # 目标格已替换
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 3)["value"], "中文说明")
        self.assertNotEqual(before[self.sheet1], after[self.sheet1])

    def test_apply_requires_file_for_empty(self):
        res = edit_service.replace_apply({"items": []})
        self.assertEqual(res["written"], 0)


class TestEditServiceWrite(EditTestBase):
    def test_write_returns_contract_fields(self):
        st = os.stat(self.path)
        res = edit_service.write(EditOp(self.ref("台账", 2, 1), "服务写回", "value",
                                        st.st_mtime, st.st_size))
        for key in ("ref", "oldValue", "newValue", "result", "opId", "backupPath", "current"):
            self.assertIn(key, res)
        self.assertEqual(res["oldValue"], "甲项目")
        self.assertEqual(res["newValue"], "服务写回")
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 1)["value"], "服务写回")
        # 写后 mtime/size 变化，可用于下次冲突检测
        st2 = os.stat(self.path)
        self.assertEqual(res["current"]["size"], st2.st_size)

    def test_write_numeric_string_from_frontend(self):
        """前端 JSON 传字符串 "456.5" 到数值列：应写成数值而不是文本。"""
        st = os.stat(self.path)
        edit_service.write(EditOp(self.ref("台账", 2, 2), "456.5", "value",
                                  st.st_mtime, st.st_size))
        info = writer.read_cell(self.path, "台账", 2, 2)
        self.assertEqual(info["kind"], "number")
        self.assertEqual(info["value"], 456.5)
        raw = _raw_cell(self.path, self.sheet1, "B2")
        self.assertIn("<v>456.5</v>", raw)
        self.assertNotIn('t="', raw)

    def test_write_text_into_text_column_stays_text(self):
        """前端传 "00123" 到文本列：必须保持文本（不能吃掉前导 0）。"""
        st = os.stat(self.path)
        edit_service.write(EditOp(self.ref("台账", 2, 3), "00123", "value",
                                  st.st_mtime, st.st_size))
        info = writer.read_cell(self.path, "台账", 2, 3)
        self.assertEqual(info["kind"], "text")
        self.assertEqual(info["value"], "00123")


class TestReadOnlyAndMisc(EditTestBase):
    def test_readonly_file_can_be_written_mode_restored(self):
        os.chmod(self.path, 0o444)  # Windows 下即“只读”属性
        try:
            res = writer.write_cell(self.path, self.ref("台账", 2, 3), "只读文件写回")
            self.assertEqual(res["result"], "ok")
            self.assertEqual(writer.read_cell(self.path, "台账", 2, 3)["value"], "只读文件写回")
            self.assertEqual(os.stat(self.path).st_mode & 0o200, 0, "写回后应恢复只读属性")
        finally:
            os.chmod(self.path, 0o666)

    def test_workspace_scope_preview(self):
        """未传 filePath/fileId 时按工作区范围预览（前端“范围：全部文件”）。"""
        from app.services import workspace_service

        original = workspace_service.files
        try:
            workspace_service.files = lambda: [
                {"id": "f1", "path": self.path, "status": "ok"},
                {"id": "f2", "path": os.path.join(self.work, "不存在.xlsx"), "status": "ok"},
            ]
            pv = edit_service.replace_preview({"find": "备注", "replace": "说明"})
            self.assertGreaterEqual(pv["count"], 3)
            self.assertIsNone(pv["filePath"])
            self.assertGreaterEqual(pv["files"], 1)
        finally:
            workspace_service.files = original


class TestRealSample(EditTestBase):
    """用真实样本（samples/项目台账统计表.xlsx）的副本跑一遍写回 + 撤销。"""

    def test_real_sample_roundtrip(self):
        src = os.path.join(ROOT, "samples", "项目台账统计表.xlsx")
        if not os.path.isfile(src):
            self.skipTest("样本文件不存在")
        p = os.path.join(self.work, "样本.xlsx")
        shutil.copyfile(src, p)
        cells = writer.scan_cells(p)
        self.assertGreater(len(cells), 0, "真实样本应能读到单元格")
        target = next(
            (c for c in cells if c.get("editable", True) and c["kind"] in ("text", "number")),
            None,
        )
        self.assertIsNotNone(target, "样本里应至少有一个可编辑的文本/数值单元格")

        before = _entries(p)
        h0 = _sha256(p)
        st = os.stat(p)
        op = EditOp(
            CellRef(p, target["sheet"], target["row"], target["col"]),
            "S8写回自检",
            "value",
            st.st_mtime,
            st.st_size,
        )
        res = edit_service.write(op)
        self.assertEqual(res["result"], "ok")
        info = writer.read_cell(p, target["sheet"], target["row"], target["col"])
        self.assertEqual(info["value"], "S8写回自检")
        after = _entries(p)
        changed = [n for n in before if before[n] != after[n]]
        self.assertEqual(len(changed), 1, f"真实样本也只应改动目标 sheet：{changed}")
        # 撤销（整文件按备份还原）
        edit_service.undo(res["opId"])
        self.assertEqual(_sha256(p), h0, "真实样本撤销后应逐字节还原")


class TestSharedStrings(EditTestBase):
    """sharedStrings 形式（Excel/WPS 真实文件常见）：只改单元格，不改共享串表。"""

    def test_shared_string_cell_edit(self):
        _convert_to_shared_strings(self.path)
        before = _entries(self.path)
        self.assertIn("xl/sharedStrings.xml", before)
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 3)["value"], "中文备注")

        res = writer.write_cell(self.path, self.ref("台账", 2, 3), "共享串改造")
        self.assertEqual(res["result"], "ok")
        after = _entries(self.path)
        self.assertEqual(after["xl/sharedStrings.xml"], before["xl/sharedStrings.xml"],
                         "sharedStrings.xml 必须按原始字节保留")
        self.assert_only_part_changed(before, after, {self.sheet1})
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 3)["value"], "共享串改造")
        # 仍然引用 sharedStrings 的其他单元格不受影响
        self.assertEqual(writer.read_cell(self.path, "台账", 2, 1)["value"], "甲项目")
        self.assertEqual(writer.read_cell(self.path, "第二表", 1, 1)["value"], "A")


class TestStructuralEdges(EditTestBase):
    def test_write_into_selfclosing_row(self):
        with zipfile.ZipFile(self.path) as z:
            entries = {n: z.read(n) for n in z.namelist()}
        xml = entries[self.sheet1].decode("utf-8")
        xml = xml.replace("</sheetData>", '<row r="10" spans="1:4"/></sheetData>', 1)
        entries[self.sheet1] = xml.encode("utf-8")
        _write_entries(self.path, entries)

        writer.write_cell(self.path, self.ref("台账", 10, 3), "空行写入")
        self.assertEqual(writer.read_cell(self.path, "台账", 10, 3)["value"], "空行写入")
        self.assertIsNotNone(_raw_cell(self.path, self.sheet1, "C10"))
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(self.path) as z:
            ET.fromstring(z.read(self.sheet1))
            xml2 = z.read(self.sheet1).decode("utf-8")
        self.assertIn('ref="A1:D10"', xml2, "写入新行后 dimension 应扩展")

    def test_invalid_coordinates_rejected(self):
        h = _sha256(self.path)
        with self.assertRaises(ValueError):
            writer.write_cell(self.path, CellRef(self.path, "台账", 0, 1), "x")
        with self.assertRaises(ValueError):
            writer.write_cell(self.path, CellRef(self.path, "台账", 1, 0), "x")
        self.assertEqual(_sha256(self.path), h)


class TestNoLeftoverTempFiles(EditTestBase):
    def test_no_tmp_files_after_write(self):
        writer.write_cell(self.path, self.ref("台账", 2, 3), "清理")
        leftovers = [n for n in os.listdir(self.work) if ".tmp-" in n]
        self.assertEqual(leftovers, [], f"写回后不应残留临时文件：{leftovers}")


if __name__ == "__main__":
    unittest.main()
