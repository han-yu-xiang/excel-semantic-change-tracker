"""Compare original field values after semantic matching."""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import math
import unicodedata

from openpyxl.utils import get_column_letter

from models import ChangeReport, ChangeRow, MatchPlan, ReportStats, WorkbookSnapshot


def _normalized_text(value: str) -> str:
    return unicodedata.normalize("NFKC", value).replace("\r\n", "\n").replace("\r", "\n")


def values_equal(left: object, right: object) -> bool:
    if left is None or right is None:
        return left is right or (left in (None, "") and right in (None, ""))
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float, Decimal)) and isinstance(right, (int, float, Decimal)):
        try:
            return Decimal(str(left)) == Decimal(str(right))
        except (InvalidOperation, ValueError):
            return float(left) == float(right)
    if isinstance(left, (datetime, date)) and isinstance(right, (datetime, date)):
        return left == right
    if isinstance(left, str) and isinstance(right, str):
        return _normalized_text(left) == _normalized_text(right)
    return type(left) is type(right) and left == right


def _location(snapshot: WorkbookSnapshot, column_index: int, row_number: int | None) -> str | int:
    if row_number is None:
        return 0
    if column_index < 0 or column_index >= len(snapshot.columns):
        raise ValueError("内容列选择超出工作表列范围。")
    column_letter = snapshot.columns[column_index].letter or get_column_letter(column_index + 1)
    return f"{snapshot.file_name}!{column_letter}{row_number}"


def build_change_report(
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
    plan: MatchPlan,
    before_content_column: int,
    after_content_column: int,
) -> ChangeReport:
    matched_before_rows = [match.before.row_number for match in plan.matches]
    matched_after_rows = [match.after.row_number for match in plan.matches]
    unmatched_before_rows = [record.row_number for record in plan.unmatched_before]
    unmatched_after_rows = [record.row_number for record in plan.unmatched_after]
    before_seen = matched_before_rows + unmatched_before_rows
    after_seen = matched_after_rows + unmatched_after_rows
    expected_before = [record.row_number for record in before.records]
    expected_after = [record.row_number for record in after.records]
    if (
        len(before_seen) != len(set(before_seen))
        or len(after_seen) != len(set(after_seen))
        or set(before_seen) != set(expected_before)
        or set(after_seen) != set(expected_after)
    ):
        raise AssertionError("匹配结果未能让 Before 和 After 的每条筛选记录恰好出现一次。")

    maximum_columns = max(len(before.columns), len(after.columns))
    categorized: dict[str, list[tuple[int, ChangeRow]]] = {
        "changed": [], "new": [], "out": [], "still": [],
    }

    for match in plan.matches:
        differences = []
        before_values = []
        after_values = []
        for index in range(maximum_columns):
            left_exists = index < len(before.columns) and index < len(match.before.values)
            right_exists = index < len(after.columns) and index < len(match.after.values)
            left = match.before.values[index] if left_exists else 0
            right = match.after.values[index] if right_exists else 0
            before_values.append(left)
            after_values.append(right)
            if left_exists != right_exists or not values_equal(left, right):
                differences.append(index)
        change_type = "changed" if differences else "still"
        row = ChangeRow(
            change_type,
            _location(before, before_content_column, match.before.row_number),
            _location(after, after_content_column, match.after.row_number),
            tuple(before_values),
            tuple(after_values),
            tuple(differences),
        )
        categorized[change_type].append((match.before.row_number, row))

    for record in plan.unmatched_after:
        values = tuple(record.values[index] if index < len(record.values) and index < len(after.columns) else 0
                       for index in range(maximum_columns))
        row = ChangeRow(
            "new", 0, _location(after, after_content_column, record.row_number),
            tuple(0 for _ in range(maximum_columns)), values, (),
        )
        categorized["new"].append((record.row_number, row))

    for record in plan.unmatched_before:
        values = tuple(record.values[index] if index < len(record.values) and index < len(before.columns) else 0
                       for index in range(maximum_columns))
        row = ChangeRow(
            "out", _location(before, before_content_column, record.row_number), 0,
            values, tuple(0 for _ in range(maximum_columns)), (),
        )
        categorized["out"].append((record.row_number, row))

    ordered_rows = []
    for category in ("changed", "new", "out", "still"):
        ordered_rows.extend(row for _, row in sorted(categorized[category], key=lambda item: item[0]))
    stats = ReportStats(
        changed_count=len(categorized["changed"]),
        new_count=len(categorized["new"]),
        out_count=len(categorized["out"]),
        still_count=len(categorized["still"]),
        before_count=len(before.records),
        after_count=len(after.records),
        matched_count=len(plan.matches),
        manual_confirmed_count=sum(1 for match in plan.matches if match.manual),
        pending_review_count=len(plan.review_cases),
    )
    if stats.changed_count + stats.new_count + stats.out_count + stats.still_count != len(before.records) + stats.new_count:
        # A successful match contributes once, an unmatched Before contributes once, and unmatched After contributes once.
        raise AssertionError("变更分类数量不一致，存在遗漏或重复记录。")
    return ChangeReport(ordered_rows, stats, tuple(plan.review_cases))
