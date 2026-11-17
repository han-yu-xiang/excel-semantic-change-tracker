import tkinter as tk
import unittest

from app import ChangeTrackerApp, build_ignored_column_options
from models import ColumnInfo, SheetPreview


def preview(path, headers):
    columns = tuple(ColumnInfo(index, chr(65 + index), header) for index, header in enumerate(headers))
    return SheetPreview(path, "Sheet1", columns, 1)


class IgnoredColumnOptionTests(unittest.TestCase):
    def test_options_pair_by_physical_index_and_show_both_original_headers(self):
        before = preview("Old.xlsx", ("ID", "Item Type", "ObjectText", "Status"))
        after = preview("New.xlsx", ("Identifier", "Category", "Description", "Owner", "Extra"))

        options = build_ignored_column_options(before, after, 2, 2)

        self.assertEqual([0, 1, 2, 3, 4], [option.index for option in options])
        self.assertIn("A_before", options[0].label)
        self.assertIn("ID", options[0].label)
        self.assertIn("A_after", options[0].label)
        self.assertIn("Identifier", options[0].label)
        self.assertIn("D_before", options[3].label)
        self.assertIn("Status", options[3].label)
        self.assertIn("D_after", options[3].label)
        self.assertIn("Owner", options[3].label)
        self.assertIn("(无此列)", options[4].label)
        self.assertIn("E_after", options[4].label)
        self.assertTrue(options[2].disabled)

    def test_content_columns_on_either_side_are_disabled(self):
        before = preview("Old.xlsx", ("ID", "Item Type", "ObjectText", "Status"))
        after = preview("New.xlsx", ("ID", "Item Type", "Status", "Description"))

        options = build_ignored_column_options(before, after, 2, 3)

        self.assertEqual({2, 3}, {option.index for option in options if option.disabled})


class IgnoredColumnGuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.app = ChangeTrackerApp()
        except tk.TclError as exc:
            self.skipTest(f"Tk cannot initialize in this environment: {exc}")
        self.app.withdraw()
        self.addCleanup(self._destroy_app)

    def _destroy_app(self):
        for callback_id in self.app.tk.call("after", "info"):
            self.app.after_cancel(callback_id)
        self.app.destroy()

    def test_content_settings_receive_their_requested_height(self):
        self.app.deiconify()
        self.app.update()
        body = self.app.winfo_children()[1]
        settings = next(
            child for child in body.winfo_children()
            if child.winfo_class() == "TLabelframe" and child.cget("text").startswith("2.")
        )

        self.assertGreaterEqual(settings.winfo_height(), settings.winfo_reqheight())

    def test_file_path_row_keeps_the_full_side_panel_width(self):
        self.app.deiconify()
        self.app.update()
        body = self.app.winfo_children()[1]
        files_box = body.winfo_children()[0]
        before_panel = files_box.winfo_children()[0]
        path_row = before_panel.winfo_children()[1]

        self.assertGreaterEqual(path_row.winfo_width(), before_panel.winfo_width() - 40)

    def test_checklist_defaults_clear_all_and_refreshes_content_column_state(self):
        before = preview("Old.xlsx", ("ID", "Item Type", "ObjectText", "Status", "Interval"))
        after = preview("New.xlsx", ("ID", "Item Type", "Description", "Status", "Interval"))
        self.app._set_preview("before", before)
        self.app._set_preview("after", after)

        self.assertEqual(set(), self.app._selected_ignored_column_indices())
        self.assertTrue(self.app._ignored_column_disabled[2])
        self.assertEqual("grid", self.app._ignored_scrollbar.winfo_manager())
        content_option = next(
            child for child in self.app._ignored_items_frame.winfo_children()
            if child.winfo_class() == "TCheckbutton" and "C_before" in child.cget("text")
        )
        self.assertTrue(content_option.instate(["disabled"]))

        self.app._select_all_ignored_columns()
        self.assertEqual({0, 1, 3, 4}, self.app._selected_ignored_column_indices())
        self.app._clear_ignored_columns()
        self.assertEqual(set(), self.app._selected_ignored_column_indices())

        self.app._ignored_column_vars[1].set(True)
        selected_content = next(
            label for label, index in self.app._column_options["before"].items() if index == 3
        )
        self.app.content_vars["before"].set(selected_content)
        self.app._refresh_ignored_columns(preserve=True)
        self.assertEqual({1}, self.app._selected_ignored_column_indices())
        self.assertEqual({2, 3}, {index for index, disabled in self.app._ignored_column_disabled.items() if disabled})

    def test_new_sheet_drops_stale_indices_and_resets_selection(self):
        before = preview("Old.xlsx", ("ID", "Item Type", "ObjectText", "Status", "Interval"))
        after = preview("New.xlsx", ("ID", "Item Type", "ObjectText", "Status", "Interval"))
        self.app._set_preview("before", before)
        self.app._set_preview("after", after)
        self.app._ignored_column_vars[4].set(True)

        self.app._set_preview("before", preview("Old.xlsx", ("ID", "Item Type", "ObjectText")))
        self.app._set_preview("after", preview("New.xlsx", ("ID", "Item Type", "ObjectText")))

        self.assertNotIn(4, self.app._ignored_column_vars)
        self.assertEqual(set(), self.app._selected_ignored_column_indices())
        self.assertEqual([0, 1, 2], [option.index for option in self.app._ignored_column_options])


if __name__ == "__main__":
    unittest.main()
