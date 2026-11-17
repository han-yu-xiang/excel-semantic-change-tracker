import unittest

from models import ColumnInfo, ExcelRecord, WorkbookSnapshot, Match, MatchPlan
from change_detector import build_change_report, values_equal


def snapshot(name, rows):
    columns = tuple(ColumnInfo(i, chr(65 + i), header) for i, header in enumerate(("ID", "Item Type", "ObjectText", "Status", "Interval")))
    records = tuple(ExcelRecord(row, tuple(values)) for row, values in rows)
    return WorkbookSnapshot(name, "Sheet1", columns, records, len(records))


class ChangeDetectorTests(unittest.TestCase):
    def test_classifies_field_changes_new_out_and_still_in_required_order(self):
        before = snapshot("Old.xlsx", [
            (2, ["R1", "Requirement", "same", "Open", 100]),
            (3, ["R2", "Requirement", "old wording", "Open", 100]),
            (4, ["R3", "Requirement", "deleted", "Open", 1]),
            (5, ["R4", "Requirement", "same", "Open", 10]),
        ])
        after = snapshot("New.xlsx", [
            (8, ["R1", "Requirement", "same", "Done", 100]),
            (9, ["R2", "Requirement", "new wording", "Open", 50]),
            (10, ["R5", "Requirement", "added", "Open", 1]),
            (11, ["R4", "Requirement", "same", "Open", 10]),
        ])
        matches = MatchPlan(
            matches=[
                Match(before.records[0], after.records[0], 1.0, "exact"),
                Match(before.records[1], after.records[1], 0.9, "semantic"),
                Match(before.records[3], after.records[3], 1.0, "exact"),
            ],
            unmatched_before=[before.records[2]],
            unmatched_after=[after.records[2]],
        )
        report = build_change_report(before, after, matches, 2, 2)
        self.assertEqual(["changed", "changed", "new", "out", "still"], [row.change_type for row in report.rows])
        self.assertEqual(["Old.xlsx!C2", "New.xlsx!C8"], [report.rows[0].before_location, report.rows[0].after_location])
        self.assertEqual(0, report.rows[2].before_values[0])
        self.assertEqual(0, report.rows[3].after_values[0])
        self.assertEqual((3,), report.rows[0].changed_column_indices)
        self.assertEqual("changed", report.rows[1].change_type)
        self.assertIn(4, report.rows[1].changed_column_indices)
        self.assertEqual(3, report.stats.matched_count)

    def test_row_number_and_line_ending_only_do_not_create_false_changes(self):
        before = snapshot("Old.xlsx", [(2, ["R1", "Requirement", "line one\r\nline two", "Open", 5])])
        after = snapshot("New.xlsx", [(55, ["R1", "Requirement", "line one\nline two", "Open", 5])])
        plan = MatchPlan(matches=[Match(before.records[0], after.records[0], 1.0, "exact")])
        report = build_change_report(before, after, plan, 2, 2)
        self.assertEqual("still", report.rows[0].change_type)

    def test_numeric_cell_types_do_not_create_false_changes(self):
        self.assertTrue(values_equal(100, 100.0))
        self.assertTrue(values_equal(None, ""))
        self.assertFalse(values_equal(0, None))
        self.assertFalse(values_equal("1", 1))

    def test_report_rejects_a_match_plan_that_loses_a_filtered_record(self):
        before = snapshot("Old.xlsx", [
            (2, ["R1", "Requirement", "one", "Open", 1]),
            (3, ["R2", "Requirement", "two", "Open", 2]),
        ])
        after = snapshot("New.xlsx", [(8, ["R1", "Requirement", "one", "Open", 1])])
        plan = MatchPlan(matches=[Match(before.records[0], after.records[0], 1.0, "exact")])
        with self.assertRaises(AssertionError):
            build_change_report(before, after, plan, 2, 2)


if __name__ == "__main__":
    unittest.main()
