import unittest

import numpy as np

from models import ColumnInfo, ExcelRecord, WorkbookSnapshot
from semantic_matcher import match_records, apply_review_decisions


def snapshot(name, rows, headers=("ID", "Item Type", "ObjectText", "Status")):
    columns = tuple(ColumnInfo(i, chr(65 + i), header) for i, header in enumerate(headers))
    records = tuple(ExcelRecord(row_number, tuple(values)) for row_number, values in rows)
    return WorkbookSnapshot(name, "Sheet1", columns, records, len(records))


class FakeEmbedder:
    """A tiny deterministic encoder for exercising the matching pipeline."""
    vectors = {
        "old speed 100": [1.0, 0.0, 0.0],
        "new speed 50": [0.99, 0.1, 0.0],
        "old voltage": [0.0, 1.0, 0.0],
        "new voltage": [0.05, 0.99, 0.0],
        "new brake": [0.0, 0.0, 1.0],
        "old action a": [0.8, 0.6, 0.0],
        "old action b": [0.8, 0.6, 0.0],
        "new action": [0.8, 0.6, 0.0],
        "new action 2": [0.8, 0.6, 0.0],
        "transmit speed": [0.8, 0.6, 0.0],
        "receive speed": [0.8, 0.6, 0.0],
        "monitor overvoltage": [0.0, 1.0, 0.0],
        "monitor undervoltage": [0.0, 1.0, 0.0],
        "not similar": [-1.0, 0.0, 0.0],
    }

    def encode(self, texts, **kwargs):
        return np.asarray([self.vectors[text] for text in texts], dtype=float)


class IdenticalEmbedder:
    def encode(self, texts, **_kwargs):
        return np.ones((len(texts), 8), dtype=float)


class SemanticMatcherTests(unittest.TestCase):
    def test_identical_content_matches_without_loading_a_semantic_model(self):
        before = snapshot("before.xlsx", [(2, ["R1", "Requirement", "same exact requirement", "Open"])])
        after = snapshot("after.xlsx", [(7, ["R1", "Requirement", "same exact requirement", "Open"])])
        plan = match_records(
            before.records,
            after.records,
            2,
            embedder=None,
            before_headers=[column.header for column in before.columns],
            after_headers=[column.header for column in after.columns],
        )
        self.assertEqual([(2, 7, "exact")], [(m.before.row_number, m.after.row_number, m.method) for m in plan.matches])

    def test_reworded_numeric_requirement_uses_semantic_embedding(self):
        before = snapshot("before.xlsx", [(2, ["R1", "Requirement", "old speed 100", "Open"])])
        after = snapshot("after.xlsx", [(7, ["R1", "Requirement", "new speed 50", "Open"])])
        plan = match_records(before.records, after.records, 2, threshold=0.77, embedder=FakeEmbedder())
        self.assertEqual(1, len(plan.matches))
        self.assertEqual("semantic", plan.matches[0].method)

    def test_content_column_can_move_between_workbook_versions(self):
        before = [ExcelRecord(2, ("R1", "Requirement", "Open", "same wording"))]
        after = [ExcelRecord(7, ("R1", "same wording", "Requirement", "Open"))]
        plan = match_records(before, after, 3, before_content_column=3, after_content_column=1, embedder=None)
        self.assertEqual([(2, 7)], [(m.before.row_number, m.after.row_number) for m in plan.matches])

    def test_different_actions_are_left_for_review_even_when_vectors_are_close(self):
        before = snapshot("before.xlsx", [(2, ["R1", "Requirement", "transmit speed", "Open"])])
        after = snapshot("after.xlsx", [(8, ["R2", "Requirement", "receive speed", "Open"])])
        plan = match_records(before.records, after.records, 2, threshold=0.77, embedder=FakeEmbedder())
        self.assertEqual([], plan.matches)
        self.assertTrue(plan.review_cases)

    def test_opposite_voltage_conditions_are_not_automatically_matched(self):
        before = snapshot("before.xlsx", [(2, ["R1", "Requirement", "monitor overvoltage", "Open"])])
        after = snapshot("after.xlsx", [(8, ["R2", "Requirement", "monitor undervoltage", "Open"])])
        plan = match_records(before.records, after.records, 2, threshold=0.77, embedder=FakeEmbedder())
        self.assertEqual([], plan.matches)

    def test_global_assignment_never_reuses_a_new_record(self):
        before = snapshot("before.xlsx", [
            (2, ["R1", "Requirement", "old action a", "Open"]),
            (3, ["R2", "Requirement", "old action b", "Open"]),
        ])
        after = snapshot("after.xlsx", [(8, ["R3", "Requirement", "new action", "Open"])])
        plan = match_records(before.records, after.records, 2, threshold=0.7, embedder=FakeEmbedder(), ambiguity_margin=0.04)
        paired_after_rows = [match.after.row_number for match in plan.matches]
        self.assertEqual(len(paired_after_rows), len(set(paired_after_rows)))
        self.assertLessEqual(len(paired_after_rows), 1)
        self.assertTrue(plan.review_cases)

    def test_candidate_recall_stays_bounded_across_large_batches(self):
        count = 257
        before = [
            ExcelRecord(row + 2, (f"R{row}", "Requirement", f"old condition {row}"))
            for row in range(count)
        ]
        after = [
            ExcelRecord(row + 1000, (f"R{row}", "Requirement", f"new condition {row}"))
            for row in range(count)
        ]
        plan = match_records(before, after, 2, threshold=0.68, embedder=IdenticalEmbedder())
        self.assertLessEqual(len(plan.candidate_edges), 20 * count)

    def test_ambiguous_pair_is_not_auto_matched_and_can_be_confirmed(self):
        before = snapshot("before.xlsx", [(2, ["R1", "Requirement", "old action a", "Open"])])
        after = snapshot("after.xlsx", [
            (8, ["R2", "Requirement", "new action", "Open"]),
            (9, ["R3", "Requirement", "new action 2", "Open"]),
        ])
        plan = match_records(before.records, after.records, 2, threshold=0.7, embedder=FakeEmbedder(), ambiguity_margin=0.04)
        self.assertEqual([], plan.matches)
        self.assertEqual(1, len(plan.review_cases))
        resolved = apply_review_decisions(plan, {2: 8})
        self.assertEqual([(2, 8, True)], [(m.before.row_number, m.after.row_number, m.manual) for m in resolved.matches])

    def test_confirming_one_pair_recomputes_other_unresolved_records(self):
        before = snapshot("before.xlsx", [
            (2, ["R1", "Requirement", "old action a", "Open"]),
            (3, ["R2", "Requirement", "old action b", "Open"]),
        ])
        after = snapshot("after.xlsx", [
            (8, ["R3", "Requirement", "new action", "Open"]),
            (9, ["R4", "Requirement", "new action 2", "Open"]),
        ])
        plan = match_records(before.records, after.records, 2, threshold=0.7, embedder=FakeEmbedder(), ambiguity_margin=0.04)
        resolved = apply_review_decisions(plan, {2: 8})
        self.assertEqual({(2, 8), (3, 9)}, {(m.before.row_number, m.after.row_number) for m in resolved.matches})
        self.assertEqual(1, sum(m.manual for m in resolved.matches))

    def test_rejecting_candidate_does_not_force_a_different_match(self):
        before = snapshot("before.xlsx", [(2, ["R1", "Requirement", "old action a", "Open"])])
        after = snapshot("after.xlsx", [
            (8, ["R2", "Requirement", "new action", "Open"]),
            (9, ["R3", "Requirement", "new action 2", "Open"]),
        ])
        plan = match_records(before.records, after.records, 2, threshold=0.7, embedder=FakeEmbedder(), ambiguity_margin=0.04)
        resolved = apply_review_decisions(plan, {2: None})
        self.assertEqual([], resolved.matches)
        self.assertEqual([2], [r.row_number for r in resolved.unmatched_before])
        self.assertEqual([8, 9], [r.row_number for r in resolved.unmatched_after])

    def test_duplicate_exact_text_uses_other_fields_and_preserves_one_to_one_pairs(self):
        before = snapshot("before.xlsx", [
            (2, ["R1", "Requirement", "same requirement", "Open"]),
            (3, ["R2", "Requirement", "same requirement", "Done"]),
        ])
        after = snapshot("after.xlsx", [
            (8, ["R2", "Requirement", "same requirement", "Open"]),
            (9, ["R1", "Requirement", "same requirement", "Done"]),
        ])
        plan = match_records(
            before.records,
            after.records,
            2,
            embedder=None,
            before_headers=[column.header for column in before.columns],
            after_headers=[column.header for column in after.columns],
        )
        pairs = {(m.before.row_number, m.after.row_number) for m in plan.matches}
        self.assertEqual({(2, 9), (3, 8)}, pairs)

    def test_duplicate_text_uses_matching_header_names_after_columns_move(self):
        before_headers = ["ID", "Item Type", "ObjectText", "Status"]
        after_headers = ["Item Type", "Status", "ID", "ObjectText"]
        before = [
            ExcelRecord(2, ("R1", "Requirement", "same requirement", "Open")),
            ExcelRecord(3, ("R2", "Requirement", "same requirement", "Done")),
        ]
        after = [
            ExcelRecord(8, ("Requirement", "Open", "R2", "same requirement")),
            ExcelRecord(9, ("Requirement", "Done", "R1", "same requirement")),
        ]
        plan = match_records(
            before,
            after,
            2,
            before_content_column=2,
            after_content_column=3,
            before_headers=before_headers,
            after_headers=after_headers,
        )
        self.assertEqual({(2, 9), (3, 8)}, {(m.before.row_number, m.after.row_number) for m in plan.matches})


if __name__ == "__main__":
    unittest.main()
