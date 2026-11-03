"""CODEX51 regression: shipment VOM action bar must remain visible.

Requires an interactive Windows Tk desktop; opens no production workbook,
does not paste data, generate documents, or touch the VZMK registry.
"""
import os
import unittest
import tkinter as tk
from tkinter import ttk
from types import SimpleNamespace

from production_app.shipment_dialog import open_shipment


@unittest.skipUnless(os.name == "nt", "Native Windows Tk geometry test")
class ShipmentDialogFooter(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.geometry("320x180+0+0")
        self.addCleanup(self.root.destroy)

        class FakeApp:
            active = None
            shipment_rows = None
            vom = SimpleNamespace(get=lambda: "")

            def __init__(self, root):
                self.root = root

            def fit(self, window, width, height):
                window.geometry("920x420+20+20")

            def editable(self, grid):
                # Only testing geometry; cell editing belongs to existing tests.
                pass

            def error(self, exc):
                raise AssertionError(str(exc))

            def write(self, message):
                pass

        self.window = open_shipment(FakeApp(self.root))
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
