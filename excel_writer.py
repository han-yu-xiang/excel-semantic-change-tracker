"""Styled .xlsx output for change tracking reports."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from models import ChangeReport, WorkbookSnapshot


TYPE_FILLS = {
    "changed": PatternFill("solid", fgColor="FCE4D6"),
    "new": PatternFill("solid", fgColor="E2F0D9"),
    "out": PatternFill("solid", fgColor="F4CCCC"),
    "still": PatternFill("solid", fgColor="EAF2F8"),
}
DIFF_FILL = PatternFill("solid", fgColor="FFF2CC")


def _safe_value(cell, value: Any) -> None:
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"


def _side_header(snapshot: WorkbookSnapshot, index: int, suffix: str) -> str:
    if index < len(snapshot.columns):
        column = snapshot.columns[index]
        return f"{column.letter}_{column.header}_{suffix}"
    return f"{get_column_letter(index + 1)}_0_{suffix}"


def save_report(
    output_path: str | Path,
    report: ChangeReport,
    before: WorkbookSnapshot,
    after: WorkbookSnapshot,
    before_content_column: int,
    after_content_column: int,
) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Change_Tracking"
    maximum_columns = max(len(before.columns), len(after.columns))
    headers = ["Change_Type", "Before_Location", "After_Location"]
    for index in range(maximum_columns):
        headers.extend((_side_header(before, index, "before"), _side_header(after, index, "after")))
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for row_index, item in enumerate(report.rows, start=2):
        values = [item.change_type, item.before_location, item.after_location]
        for column_index in range(maximum_columns):
            before_value = item.before_values[column_index] if column_index < len(item.before_values) else 0
            after_value = item.after_values[column_index] if column_index < len(item.after_values) else 0
            values.extend((before_value, after_value))
        for column_index, value in enumerate(values, start=1):
            _safe_value(sheet.cell(row=row_index, column=column_index), value)
        type_fill = TYPE_FILLS[item.change_type]
        for cell in sheet[row_index]:
            cell.fill = type_fill
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for physical_index in item.changed_column_indices:
            before_col = 4 + physical_index * 2
            after_col = before_col + 1
            sheet.cell(row=row_index, column=before_col).fill = DIFF_FILL
            sheet.cell(row=row_index, column=after_col).fill = DIFF_FILL

    sheet.freeze_panes = "D2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, sheet.max_row)}"
    sheet.row_dimensions[1].height = 32
    for row_index in range(2, sheet.max_row + 1):
        sheet.row_dimensions[row_index].height = 42
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 34
    sheet.column_dimensions["C"].width = 34
    for index in range(maximum_columns):
        source_column = before.columns[index] if index < len(before.columns) else None
        target_column = after.columns[index] if index < len(after.columns) else None
        content_side = (
            (source_column is not None and source_column.index == before_content_column)
            or (target_column is not None and target_column.index == after_content_column)
        )
        width = 58 if content_side else 22
        sheet.column_dimensions[get_column_letter(4 + index * 2)].width = width
        sheet.column_dimensions[get_column_letter(5 + index * 2)].width = width

    summary = workbook.create_sheet("Summary")
    summary.append(["Metric", "Count"])
    summary_rows = [
        ("changed", report.stats.changed_count),
        ("new", report.stats.new_count),
        ("out", report.stats.out_count),
        ("still", report.stats.still_count),
        ("Before filtered records", report.stats.before_count),
        ("After filtered records", report.stats.after_count),
        ("Matched records", report.stats.matched_count),
        ("Manually confirmed matches", report.stats.manual_confirmed_count),
        ("Pending review candidates", report.stats.pending_review_count),
    ]
    for row in summary_rows:
        summary.append(row)
    for cell in summary[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 14
    summary.freeze_panes = "A2"

    if report.pending_candidates:
        review = workbook.create_sheet("Review_Candidates")
        review_headers = ["Before_Row", "Before_Content", "After_Row", "After_Content", "Score", "Method"]
        review.append(review_headers)
        for cell in review[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="7F6000")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        for case in report.pending_candidates:
            before_text = case.before.values[before_content_column] if before_content_column < len(case.before.values) else None
            for candidate in case.candidates:
                after_text = candidate.after.values[after_content_column] if after_content_column < len(candidate.after.values) else None
                values = [
                    case.before.row_number,
                    before_text,
                    candidate.after.row_number,
                    after_text,
                    round(candidate.score, 4),
                    candidate.method,
                ]
                review_row = review.max_row + 1
                for column_index, value in enumerate(values, start=1):
                    _safe_value(review.cell(row=review_row, column=column_index), value)
        review.freeze_panes = "A2"
        review.auto_filter.ref = f"A1:F{max(1, review.max_row)}"
        for column, width in (("A", 13), ("B", 65), ("C", 13), ("D", 65), ("E", 12), ("F", 22)):
            review.column_dimensions[column].width = width
        for row in review.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    try:
        workbook.save(output_path)
    finally:
        workbook.close()
    return output_path
