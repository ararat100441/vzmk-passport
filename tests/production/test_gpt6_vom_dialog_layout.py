"""CODEX51 regression: shipment VOM action bar must remain visible.

Requires an interactive Windows Tk desktop; opens no production workbook,
does not paste data, generate documents, or touch the VZMK registry.
"""
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import tkinter as tk
from tkinter import ttk

from production_app.shipment_dialog import open_shipment


@unittest.skipUnless(os.name == "nt", "Native Windows Tk geometry test")
class ShipmentDialogFooter(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.geometry("320x180+0+0")
        self.addCleanup(self.root.destroy)
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

        class FakeApp:
            active = None

            def __init__(self, root):
                self.root = root
                self.shipment_rows = None
                self.vom = tk.StringVar(master=root, value="")
                self.vom_summary = tk.StringVar(master=root, value="")
                self.messages = []
                self.on_change = None

            def fit(self, window, width, height):
                window.geometry("920x420+20+20")

            def editable(self, grid, **kwargs):
                # Capture the callback wired to the existing cell editor.
                self.on_change = kwargs.get("on_change")

            def error(self, exc):
                raise AssertionError(str(exc))

            def write(self, message):
                self.messages.append(message)

        self.app = FakeApp(self.root)
        self.window = open_shipment(self.app)
        self.addCleanup(self.safe_destroy)

    def safe_destroy(self):
        if self.window.winfo_exists():
            self.window.grab_release()
            self.window.destroy()

    def find_controls(self):
        found = {}

        def visit(widget):
            for child in widget.winfo_children():
                if isinstance(child, ttk.Button):
                    found[child.cget("text")] = child
                visit(child)

        visit(self.window)
        return found

    def tree(self):
        trees = []

        def visit(widget):
            for child in widget.winfo_children():
                if isinstance(child, ttk.Treeview):
                    trees.append(child)
                visit(child)

        visit(self.window)
        self.assertEqual(len(trees), 1)
        return trees[0]

    def status(self):
        labels = [w for w in self.window.winfo_children()
                  if isinstance(w, ttk.Label)
                  and str(w.cget("text")).startswith(("ВОМ", "3 строк", "2 строк"))]
        self.assertEqual(len(labels), 1)
        return str(labels[0].cget("text"))

    def add_three_rows(self):
        grid = self.tree()
        entries = [
            ("К1-3", "Колонна", "1", "300", "300", "10", "10"),
            ("Б1-1", "Балка", "1", "250", "250", "8", "8"),
            ("С1-1", "Связь", "1", "292.92", "292.92", "8.877", "8.877"),
        ]
        for row in entries:
            grid.insert("", "end", values=row)
        self.assertTrue(callable(self.app.on_change))
        self.app.on_change()
        return grid

    def test_removing_row_updates_totals_exactly(self):
        grid = self.add_three_rows()
        self.assertIn("3 изделий / 842.92 кг / 26.877 м²", self.status())
        grid.selection_set(grid.get_children()[0])
        self.find_controls()["Удалить строку"].invoke()
        self.assertEqual(len(grid.get_children()), 2)
        self.assertIn("2 изделий / 542.92 кг / 16.877 м²", self.status())
        self.assertNotIn("842.92", self.status())

    def test_deleting_every_row_clears_the_summary(self):
        grid = self.add_three_rows()
        grid.selection_set(grid.get_children())
        self.find_controls()["Удалить строку"].invoke()
        self.assertEqual(grid.get_children(), ())
        self.assertEqual(self.status(), "ВОМ пуста — вставьте строки из Excel.")

    def test_editing_a_cell_invalidates_old_totals(self):
        grid = self.add_three_rows()
        item = grid.get_children()[0]
        values = list(grid.item(item, "values"))
        values[2] = "2"  # Quantity changed; total mass is no longer valid.
        grid.item(item, values=values)
        self.app.on_change()  # Same callback used by App.editable on commit.
        self.assertIn("ВОМ требует исправления:", self.status())
        self.assertNotIn("842.92 кг", self.status())

    def test_accept_after_deletion_writes_only_updated_private_vom(self):
        from openpyxl import load_workbook
        grid = self.add_three_rows()
        grid.selection_set(grid.get_children()[0])
        self.find_controls()["Удалить строку"].invoke()
        with patch("production_app.shipment_dialog.local_root",
                   return_value=Path(self.temp.name)):
            self.find_controls()["ОК"].invoke()
        self.assertFalse(self.window.winfo_exists())
        self.assertEqual(len(self.app.shipment_rows), 2)
        self.assertIn("542.92 кг", self.app.vom_summary.get())
        output = Path(self.app.vom.get())
        self.assertTrue(output.is_relative_to(self.temp.name))
        self.assertTrue(output.is_file())
        book = load_workbook(output, read_only=True)
        try:
            rows = list(book.active.values)
            self.assertEqual(len(rows), 4)  # 2 header + 2 shipment lines
        finally:
            book.close()

    def test_cancel_discards_unaccepted_changes(self):
        self.add_three_rows()
        self.find_controls()["Отмена"].invoke()
        self.assertFalse(self.window.winfo_exists())
        self.assertFalse(self.app.vom.get())
        self.assertIsNone(self.app.shipment_rows)

    def test_footer_buttons_remain_visible_when_treeview_is_clipped(self):
        buttons = self.find_controls()
        self.assertEqual(set(buttons), {"ОК", "Отмена", "Удалить строку"})
        footer = buttons["ОК"].master
        self.assertIs(footer, buttons["Отмена"].master)
        self.assertIs(footer, buttons["Удалить строку"].master)

        # In Tk pack(), earlier managed bottom widgets keep their space before
        # expandable siblings. This guards against the original defect.
        self.assertIs(self.window.pack_slaves()[0], footer)
        self.assertEqual(footer.pack_info()["side"], "bottom")

        for height in (360, 420):
            with self.subTest(height=height):
                self.window.geometry(f"920x{height}+20+20")
                self.window.deiconify()
                self.root.update()
                lower_edge = self.window.winfo_rooty() + self.window.winfo_height()
                for name, button in buttons.items():
                    self.assertTrue(
                        button.winfo_ismapped(), f"{name}: not mapped"
                    )
                    self.assertGreater(
                        button.winfo_height(), 0, f"{name}: zero height"
                    )
                    self.assertLessEqual(
                        button.winfo_rooty() + button.winfo_height(),
                        lower_edge,
                        f"{name}: outside visible window",
                    )

    def test_treeview_does_not_require_twenty_rows_at_startup(self):
        trees = []

        def visit(widget):
            for child in widget.winfo_children():
                if isinstance(child, ttk.Treeview):
                    trees.append(child)
                visit(child)

        visit(self.window)
        self.assertEqual(len(trees), 1)
        self.assertLessEqual(int(trees[0].cget("height")), 8)
        self.assertIsNotNone(trees[0].cget("yscrollcommand"))


if __name__ == "__main__":
    unittest.main()
