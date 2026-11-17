"""Excel workbook discovery and Item Type filtering."""

from __future__ import annotations

from pathlib import Path
import unicodedata

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from models import ColumnInfo, ExcelRecord, SheetPreview, WorkbookSnapshot


SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}


def _open_workbook(path: str | Path):
    workbook_path = Path(path)
    if workbook_path.suffix.casefold() not in SUPPORTED_SUFFIXES:
        raise ValueError("只支持 .xlsx 和 .xlsm 文件。")
    return load_workbook(
        workbook_path,
        read_only=True,
        data_only=False,
        keep_vba=workbook_path.suffix.casefold() == ".xlsm",
    )


def list_sheet_names(path: str | Path) -> tuple[str, ...]:
    workbook = _open_workbook(path)
    try:
        return tuple(workbook.sheetnames)
    finally:
        workbook.close()


def read_sheet_preview(path: str | Path, sheet_name: str) -> SheetPreview:
    workbook = _open_workbook(path)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"工作表不存在：{sheet_name}")
        sheet = workbook[sheet_name]
        header_values = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        columns = tuple(
            ColumnInfo(index=i, letter=get_column_letter(i + 1), header="" if value is None else str(value))
            for i, value in enumerate(header_values)
        )
        return SheetPreview(path, sheet_name, columns, max(0, (sheet.max_row or 0) - 1))
    finally:
        workbook.close()


def _normalized_filter(value: object) -> str:
    if value is None:
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip().casefold()


def read_filtered_sheet(
    path: str | Path,
    sheet_name: str,
    item_type_column: int,
    filter_value: str,
) -> WorkbookSnapshot:
    """Read a worksheet and retain matching rows with their original Excel row numbers.

    ``item_type_column`` is zero-based, matching the selection widgets and Python APIs.
    """
    preview = read_sheet_preview(path, sheet_name)
    if item_type_column < 0 or item_type_column >= len(preview.columns):
        raise ValueError("Item Type 列选择超出工作表列范围。")

    workbook = _open_workbook(path)
    expected = _normalized_filter(filter_value)
    records: list[ExcelRecord] = []
    try:
        sheet = workbook[sheet_name]
        for row_number, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if item_type_column >= len(values):
                item_type = None
            else:
                item_type = values[item_type_column]
            if _normalized_filter(item_type) == expected:
                records.append(ExcelRecord(row_number, tuple(values[: len(preview.columns)])))
    finally:
        workbook.close()

    return WorkbookSnapshot(
        path=path,
        sheet_name=sheet_name,
        columns=preview.columns,
        records=tuple(records),
        source_data_row_count=preview.data_row_count,
    )
