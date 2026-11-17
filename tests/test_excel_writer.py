from pathlib import Path
import tempfile
import unittest

from openpyxl import load_workbook

from models import (
    ColumnInfo,
    ExcelRecord,
    WorkbookSnapshot,
    ChangeRow,
    ChangeReport,
    ReportStats,
    MatchCandidate,
    ReviewCase,
)
from excel_writer import save_report


class ExcelWriterTests(unittest.TestCase):
    def test_report_has_dynamic_headers_styles_summary_and_numeric_zero_placeholders(self):
        before = WorkbookSnapshot("Old.xlsx", "Sheet1", (
            ColumnInfo(0, "A", "ID"), ColumnInfo(1, "B", "ObjectText"),
        ), (), 0)
        after = WorkbookSnapshot("New.xlsx", "Sheet1", (
            ColumnInfo(0, "A", "ID"), ColumnInfo(1, "B", "ObjectText"), ColumnInfo(2, "C", "Owner"),
        ), (), 0)
        row = ChangeRow("new", 0, "New.xlsx!B4", (0, 0), ("R2", "新需求", "Lee"), ())
        out_row = ChangeRow("out", "Old.xlsx!B7", 0, ("R3", "已删除"), (0, 0, 0), ())
        report = ChangeReport([row, out_row], ReportStats(0, 1, 1, 0, 1, 1, 0, 0))
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "result.xlsx"
            save_report(output, report, before, after, before_content_column=1, after_content_column=1)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook["Change_Tracking"]
            self.assertEqual("Change_Type", sheet["A1"].value)
            self.assertEqual("A_ID_before", sheet["D1"].value)
            self.assertEqual("C_Owner_after", sheet["I1"].value)
            self.assertEqual(["new", 0, "New.xlsx!B4", 0, "R2", 0, "新需求", 0, "Lee"],
                             [sheet.cell(2, column).value for column in range(1, 10)])
            self.assertEqual(["out", "Old.xlsx!B7", 0, "R3", 0, "已删除", 0, 0, 0],
                             [sheet.cell(3, column).value for column in range(1, 10)])
            self.assertEqual("D2", sheet.freeze_panes)
            self.assertTrue(sheet.auto_filter.ref)
            self.assertTrue(sheet["A1"].font.bold)
            self.assertIn("Summary", workbook.sheetnames)
            self.assertIsNotNone(workbook["Summary"]["A1"].value)

    def test_pending_candidates_are_exported_as_traceable_text_not_formulas(self):
        columns = (
            ColumnInfo(0, "A", "ID"),
            ColumnInfo(1, "B", "ObjectText"),
        )
        before_row = ExcelRecord(4, ("R1", "=1+1"))
        after_row = ExcelRecord(9, ("R1", "中文候选"))
        before = WorkbookSnapshot("Old.xlsx", "Sheet1", columns, (before_row,), 1)
        after = WorkbookSnapshot("New.xlsx", "Sheet1", columns, (after_row,), 1)
        candidate = MatchCandidate(after_row, 0.68, "semantic")
        case = ReviewCase(before_row, (candidate,))
        report = ChangeReport(
            [],
            ReportStats(0, 0, 0, 0, 1, 1, 0, 0, 1),
            (case,),
        )
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "review.xlsx"
            save_report(output, report, before, after, 1, 1)
            workbook = load_workbook(output, data_only=True)
            self.assertIn("Review_Candidates", workbook.sheetnames)
            review = workbook["Review_Candidates"]
            self.assertEqual("=1+1", review["B2"].value)
            self.assertEqual("s", review["B2"].data_type)
            self.assertEqual("中文候选", review["D2"].value)


if __name__ == "__main__":
    unittest.main()
