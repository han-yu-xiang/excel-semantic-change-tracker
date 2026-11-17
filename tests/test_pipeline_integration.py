from pathlib import Path
import tempfile
import unittest

import numpy as np
from openpyxl import Workbook, load_workbook

from change_detector import build_change_report
from excel_reader import read_filtered_sheet
from excel_writer import DIFF_FILL, save_report
from semantic_matcher import match_records


class FixtureEmbedder:
    vectors = {
        "The vehicle speed signal is sent every 100 ms.": [1.0, 0.0, 0.0],
        "Send vehicle speed at 50 ms intervals.": [0.99, 0.1, 0.0],
        "The system shall detect an overvoltage condition.": [0.0, 1.0, 0.0],
        "The ECU monitors supply voltage to identify overvoltage.": [0.05, 0.99, 0.0],
        "The system shall retain a service log.": [0.0, 0.0, 1.0],
    }

    def encode(self, texts, **_kwargs):
        return np.asarray([self.vectors[text] for text in texts], dtype=float)


class PipelineIntegrationTests(unittest.TestCase):
    def test_filtered_workbooks_produce_semantic_changes_new_row_and_report_locations(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            before_path = directory / "Before.xlsx"
            after_path = directory / "After.xlsx"
            for path, rows in (
                (before_path, [
                    ["R1", "Requirement", "Open", "The vehicle speed signal is sent every 100 ms.", 100],
                    ["R2", "Requirement", "Open", "The system shall detect an overvoltage condition.", None],
                    ["H1", "Heading", "Open", "Chapter title", 0],
                ]),
                (after_path, [
                    ["R2", "Requirement", "Done", "The ECU monitors supply voltage to identify overvoltage.", None],
                    ["R1", "Requirement", "Open", "Send vehicle speed at 50 ms intervals.", 50],
                    ["R3", "Requirement", "Open", "The system shall retain a service log.", 0],
                ]),
            ):
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Requirements"
                sheet.append(["ID", "Item Type", "Status", "ObjectText", "Value"])
                for row in rows:
                    sheet.append(row)
                workbook.save(path)

            before = read_filtered_sheet(before_path, "Requirements", 1, "Requirement")
            after = read_filtered_sheet(after_path, "Requirements", 1, "Requirement")
            headers = [column.header for column in before.columns]
            plan = match_records(
                before.records,
                after.records,
                3,
                before_content_column=3,
                after_content_column=3,
                threshold=0.77,
                embedder=FixtureEmbedder(),
                before_headers=headers,
                after_headers=headers,
            )
            self.assertEqual({(2, 3), (3, 2)}, {(m.before.row_number, m.after.row_number) for m in plan.matches})
            report = build_change_report(before, after, plan, 3, 3)
            self.assertEqual(["changed", "changed", "new"], [row.change_type for row in report.rows])
            self.assertEqual("Before.xlsx!D2", report.rows[0].before_location)
            self.assertEqual("After.xlsx!D3", report.rows[0].after_location)

            output = directory / "Change_Tracking_Result.xlsx"
            save_report(output, report, before, after, 3, 3)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook["Change_Tracking"]
            self.assertEqual(4, sheet.max_row)
            self.assertEqual("changed", sheet["A2"].value)
            self.assertEqual("new", sheet["A4"].value)
            self.assertEqual(0, sheet["D4"].value)
            self.assertEqual("R3", sheet["E4"].value)
            self.assertEqual(2, workbook["Summary"]["B2"].value)
            self.assertEqual(1, workbook["Summary"]["B3"].value)

    def test_ignored_change_is_preserved_in_excel_but_not_highlighted(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            before_path = directory / "Before.xlsx"
            after_path = directory / "After.xlsx"
            for path, status, interval in (
                (before_path, "Open", 100),
                (after_path, "Done", 50),
            ):
                workbook = Workbook()
                sheet = workbook.active
                sheet.append(["ID", "Item Type", "ObjectText", "Status", "Interval"])
                sheet.append(["R1", "Requirement", "Same wording", status, interval])
                workbook.save(path)

            before = read_filtered_sheet(before_path, "Sheet", 1, "Requirement")
            after = read_filtered_sheet(after_path, "Sheet", 1, "Requirement")
            plan = match_records(before.records, after.records, 2, embedder=None)
            report = build_change_report(before, after, plan, 2, 2, ignored_column_indices={4})
            self.assertEqual("changed", report.rows[0].change_type)
            self.assertEqual((3,), report.rows[0].changed_column_indices)

            output = directory / "ignored-column-result.xlsx"
            save_report(output, report, before, after, 2, 2)
            workbook = load_workbook(output, data_only=True)
            sheet = workbook["Change_Tracking"]
            self.assertEqual("Open", sheet.cell(2, 10).value)
            self.assertEqual("Done", sheet.cell(2, 11).value)
            self.assertEqual(100, sheet.cell(2, 12).value)
            self.assertEqual(50, sheet.cell(2, 13).value)
            self.assertEqual(DIFF_FILL.fgColor.rgb, sheet.cell(2, 10).fill.fgColor.rgb)
            self.assertEqual(DIFF_FILL.fgColor.rgb, sheet.cell(2, 11).fill.fgColor.rgb)
            self.assertNotEqual(DIFF_FILL.fgColor.rgb, sheet.cell(2, 12).fill.fgColor.rgb)
            self.assertNotEqual(DIFF_FILL.fgColor.rgb, sheet.cell(2, 13).fill.fgColor.rgb)

    def test_item_type_filter_still_works_when_type_column_is_ignored(self):
        with tempfile.TemporaryDirectory() as directory:
            directory = Path(directory)
            snapshots = []
            for name, type_value in (("Before", " Requirement "), ("After", "requirement")):
                path = directory / f"{name}.xlsx"
                workbook = Workbook()
                sheet = workbook.active
                sheet.append(["ID", "Item Type", "ObjectText", "Status"])
                sheet.append(["R1", type_value, "Same wording", "Open"])
                sheet.append(["H1", "Heading", "Chapter title", "Open"])
                workbook.save(path)
                snapshots.append(read_filtered_sheet(path, "Sheet", 1, "Requirement"))

            before, after = snapshots
            self.assertEqual([2], [record.row_number for record in before.records])
            self.assertEqual([2], [record.row_number for record in after.records])
            plan = match_records(before.records, after.records, 2, embedder=None)
            report = build_change_report(before, after, plan, 2, 2, ignored_column_indices={1})

            self.assertEqual(["still"], [row.change_type for row in report.rows])
            self.assertEqual(" Requirement ", report.rows[0].before_values[1])
            self.assertEqual("requirement", report.rows[0].after_values[1])


if __name__ == "__main__":
    unittest.main()
