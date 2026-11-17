from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook

from excel_reader import read_sheet_preview, read_filtered_sheet


class ExcelReaderTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.path = Path(self.temp_dir.name) / "requirements.xlsx"
        workbook = Workbook()
        first = workbook.active
        first.title = "Overview"
        first.append(["ID", "Item Type", "ObjectText", "Status"])
        first.append(["I-1", " Requirement ", "车速每100毫秒发送", "Open"])
        first.append(["I-2", "Information", "说明行", "Open"])
        first.append(["I-3", "requirement", "电压过压监测", "Done"])
        other = workbook.create_sheet("Archive")
        other.append(["Key", "Category"])
        other.append(["X", "Requirement"])
        workbook.save(self.path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_preview_exposes_physical_column_letters_and_sheet_dimensions(self):
        preview = read_sheet_preview(self.path, "Overview")
        self.assertEqual([("A", 1, "ID"), ("B", 2, "Item Type"), ("C", 3, "ObjectText"), ("D", 4, "Status")],
                         [(column.letter, column.number, column.header) for column in preview.columns])
        self.assertEqual(3, preview.data_row_count)

    def test_filter_is_trimmed_case_insensitive_and_preserves_original_rows(self):
        snapshot = read_filtered_sheet(self.path, "Overview", item_type_column=1, filter_value="Requirement")
        self.assertEqual([2, 4], [record.row_number for record in snapshot.records])
        self.assertEqual("车速每100毫秒发送", snapshot.records[0].values[2])
        self.assertEqual(3, snapshot.source_data_row_count)

    def test_custom_filter_and_selected_sheet_are_supported(self):
        snapshot = read_filtered_sheet(self.path, "Archive", item_type_column=1, filter_value="requirement")
        self.assertEqual("Archive", snapshot.sheet_name)
        self.assertEqual([2], [record.row_number for record in snapshot.records])

    def test_custom_filter_value_selects_a_different_item_type(self):
        snapshot = read_filtered_sheet(self.path, "Overview", item_type_column=1, filter_value="Information")
        self.assertEqual([3], [record.row_number for record in snapshot.records])

    def test_xlsm_extension_is_supported_for_reading(self):
        macro_enabled_path = Path(self.temp_dir.name) / "requirements.xlsm"
        macro_enabled_path.write_bytes(self.path.read_bytes())
        snapshot = read_filtered_sheet(macro_enabled_path, "Overview", item_type_column=1, filter_value="Requirement")
        self.assertEqual([2, 4], [record.row_number for record in snapshot.records])


if __name__ == "__main__":
    unittest.main()
