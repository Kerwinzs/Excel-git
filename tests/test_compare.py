import csv
import tempfile
import unittest
from datetime import datetime
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl.styles import Font

from excelgit.common import ExcelGitError
from excelgit.compare import compare_workbooks, open_workbook, write_report
from .fixtures import workbook_bytes


class ComparisonTests(unittest.TestCase):
    def test_identical_workbooks_have_no_changes(self):
        data = workbook_bytes({"数据": {"A1": "青岛", "B2": 12}})
        self.assertEqual(compare_workbooks(data, data), [])

    def test_add_change_remove_in_row_order(self):
        old = workbook_bytes({"数据": {"A1": 1, "B2": "移除", "A10": 10}})
        new = workbook_bytes({"数据": {"A1": 2, "C2": "新增", "A10": 11}})
        changes = compare_workbooks(old, new)
        self.assertEqual([(c.cell, c.change) for c in changes], [
            ("A1", "cell_changed"), ("B2", "cell_removed"),
            ("C2", "cell_added"), ("A10", "cell_changed")])
        self.assertEqual(changes[0].before.value, 1)
        self.assertEqual(changes[0].after.value, 2)

    def test_formula_text_is_compared_without_calculation(self):
        old = workbook_bytes({"数据": {"A1": "=1+1"}})
        new = workbook_bytes({"数据": {"A1": "=2"}})
        change, = compare_workbooks(old, new)
        self.assertEqual(change.before.kind, "f")
        self.assertEqual(change.after.value, "=2")

    def test_types_distinguish_boolean_number_and_text(self):
        for before, after in [(True, 1), (1, "1"), (False, 0)]:
            with self.subTest(before=before, after=after):
                old = workbook_bytes({"数据": {"A1": before}})
                new = workbook_bytes({"数据": {"A1": after}})
                self.assertEqual(len(compare_workbooks(old, new)), 1)

    def test_zero_and_false_are_not_treated_as_blank(self):
        old = workbook_bytes()
        new = workbook_bytes({"数据": {"A1": 0, "A2": False}})
        self.assertEqual(len(compare_workbooks(old, new)), 2)

    def test_added_and_deleted_empty_sheets_are_reported(self):
        old = workbook_bytes({"保留": {}, "旧表": {}})
        new = workbook_bytes({"保留": {}, "新表": {}})
        self.assertEqual({(c.sheet, c.change) for c in compare_workbooks(old, new)},
                         {("旧表", "sheet_removed"), ("新表", "sheet_added")})

    def test_removed_sheet_includes_its_old_cells(self):
        old = workbook_bytes({"保留": {}, "旧表": {"A1": "值"}})
        new = workbook_bytes({"保留": {}})
        changes = compare_workbooks(old, new)
        self.assertEqual([c.change for c in changes], ["sheet_removed", "cell_removed"])

    def test_style_only_changes_are_ignored(self):
        from openpyxl import load_workbook

        old = workbook_bytes({"数据": {"A1": 1}})
        book = load_workbook(BytesIO(old))
        book.active["A1"].font = Font(bold=True)
        book.active["D5"].font = Font(bold=True)  # A formatted empty cell.
        stream = BytesIO()
        book.save(stream)
        book.close()
        self.assertEqual(compare_workbooks(old, stream.getvalue()), [])

    def test_dates_are_readable_and_inputs_unchanged(self):
        old = workbook_bytes({"数据": {"A1": datetime(2026, 1, 1)}})
        new = workbook_bytes({"数据": {"A1": datetime(2026, 1, 2)}})
        original = (old, new)
        with tempfile.TemporaryDirectory() as folder:
            report = Path(folder) / "日期.csv"
            write_report(compare_workbooks(old, new), report)
            self.assertIn("2026-01-01T00:00:00", report.read_text(encoding="utf-8-sig"))
        self.assertEqual((old, new), original)

    def test_csv_handles_chinese_commas_newlines_and_literal_formulas(self):
        old = workbook_bytes({"=危险表名": {"A1": "旧值"}})
        new = workbook_bytes({"=危险表名": {"A1": "中文,引号\"\n下一行", "B1": "=1+1"}})
        with tempfile.TemporaryDirectory() as folder:
            output = Path(folder) / "diff.csv"
            write_report(compare_workbooks(old, new), output)
            self.assertTrue(output.read_bytes().startswith(b"\xef\xbb\xbf"))
            with output.open(encoding="utf-8-sig", newline="") as stream:
                rows = list(csv.reader(stream))
            self.assertEqual(rows[1][0], "'=危险表名")
            self.assertIn("中文,引号", rows[1][6])
            self.assertEqual(rows[2][6], '"=1+1"')
            original = output.read_bytes()
            with self.assertRaises(ExcelGitError):
                write_report([], output)
            self.assertEqual(output.read_bytes(), original)

    def test_invalid_workbook_is_a_friendly_error(self):
        with self.assertRaises(ExcelGitError):
            open_workbook(b"not a workbook")

    def test_scan_limit_is_explicit(self):
        data = workbook_bytes({"数据": {"B2": 1}})
        with patch("excelgit.compare.MAX_CELLS", 3):
            with self.assertRaises(ExcelGitError):
                compare_workbooks(data, data)
