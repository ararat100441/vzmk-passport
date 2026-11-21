import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook

from production_app.shared_uzk import HEADERS, decrement_candidate, read_shared_uzk, refund_candidate
from production_app.storage import Stop, sha

X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
BOOK = "xl/workbook.xml"
SHEET = "xl/worksheets/sheet2.xml"


class FilteredRefund(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        root = Path(folder.name)
        self.start = root / "start.xlsm"
        self.spent = root / "spent.xlsm"
        self.result = root / "result.xlsm"
        workbook = Workbook()
        first = workbook.active
        first.title = "Иной объект"
        first.append(["Keep", "All"])
        first.append(["unchanged", 7])
        first.auto_filter.ref = "A1:B2"
        sheet = workbook.create_sheet("УЗК")
        sheet.append(HEADERS)
        for row in [("P", "A", 1), ("P", "M", 2), ("P", "B", 7), ("Q", "N", 3)]:
            sheet.append(row)
        sheet.auto_filter.ref = "A1:C5"
        workbook.save(self.start)
        workbook.close()
        decrement_candidate(self.start, self.spent, "P", {"A": 1, "M": 2})
        self.originals = {mark: {"project": "P", "mark": mark} for mark in ("A", "M", "B")}

    def refund(self, quantities):
        refund_candidate(self.spent, self.result, "P", quantities, self.originals)

    def change_part(self, path, part, edit):
        other = path.with_name("edited.xlsm")
        with ZipFile(path) as before, ZipFile(other, "x") as after:
            for item in before.infolist():
                raw = before.read(item)
                if item.filename == part:
                    tree = etree.fromstring(raw)
                    edit(tree)
                    raw = etree.tostring(tree)
                after.writestr(item, raw)
        other.replace(path)

    def test_restore_two_exhausted_marks_and_extend_only_expected_ranges(self):
        original_hash = sha(self.spent)
        self.refund({"A": 1, "M": 2})
        self.assertEqual(sha(self.spent), original_hash)
        before = read_shared_uzk(self.start)
        after = read_shared_uzk(self.result)
        self.assertEqual({k: v["remaining"] for k, v in before.items()},
                         {k: v["remaining"] for k, v in after.items()})
        with ZipFile(self.spent) as first, ZipFile(self.result) as last:
            self.assertIsNone(last.testzip())
            self.assertEqual(set(first.namelist()), set(last.namelist()))
            self.assertEqual({p for p in first.namelist() if first.read(p) != last.read(p)},
                             {BOOK, SHEET})
            for book, end in [(first, 3), (last, 5)]:
                tree = etree.fromstring(book.read(SHEET))
                self.assertEqual(tree.find(X + "dimension").get("ref"), f"A1:C{end}")
                self.assertEqual(tree.find(X + "autoFilter").get("ref"), f"A1:C{end}")
                names = etree.fromstring(book.read(BOOK)).findall(X + "definedNames/" + X + "definedName")
                self.assertEqual({node.get("localSheetId"): node.text for node in names},
                                 {"0": "'Иной объект'!$A$1:$B$2", "1": f"'УЗК'!$A$1:$C${end}"})
            old_book = etree.fromstring(first.read(BOOK))
            old_book.findall(X + "definedNames/" + X + "definedName")[1].text = "'УЗК'!$A$1:$C$5"
            self.assertEqual(etree.tostring(old_book, method="c14n"),
                             etree.tostring(etree.fromstring(last.read(BOOK)), method="c14n"))

    def test_wider_used_range_and_office_filter_uid_are_preserved(self):
        revision_uid = "{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}uid"

        def widen(tree):
            tree.find(X + "dimension").set("ref", "A1:F3")
            tree.find(X + "autoFilter").set(revision_uid, "{00000000-0000-0000-0000-000000000001}")

        self.change_part(self.spent, SHEET, widen)
        self.refund({"M": 2})
        with ZipFile(self.spent) as old, ZipFile(self.result) as new:
            first = etree.fromstring(old.read(SHEET))
            last = etree.fromstring(new.read(SHEET))
            self.assertEqual(last.find(X + "dimension").get("ref"), "A1:F4")
            self.assertEqual(last.find(X + "autoFilter").get("ref"), "A1:C4")
            self.assertEqual(first.find(X + "autoFilter").get(revision_uid),
                             last.find(X + "autoFilter").get(revision_uid))

    def test_shorter_filter_than_used_rows_stops_before_publication(self):
        def shorten(tree):
            tree.find(X + "autoFilter").set("ref", "A1:C2")

        self.change_part(self.spent, SHEET, shorten)
        with self.assertRaises(Stop):
            self.refund({"M": 2})
        self.assertFalse(self.result.exists())

    def test_active_criteria_stop_without_output(self):
        def add_criterion(tree):
            etree.SubElement(tree.find(X + "autoFilter"), X + "filterColumn", colId="0")

        self.change_part(self.spent, SHEET, add_criterion)
        before = sha(self.spent)
        with self.assertRaises(Stop):
            self.refund({"A": 1})
        self.assertEqual(sha(self.spent), before)
        self.assertFalse(self.result.exists())

    def test_scoped_name_with_wrong_end_stops_without_output(self):
        def change_end(tree):
            name = tree.findall(X + "definedNames/" + X + "definedName")[1]
            name.text = "'УЗК'!$A$1:$C$77"

        self.change_part(self.spent, BOOK, change_end)
        with self.assertRaises(Stop):
            self.refund({"M": 2})
        self.assertFalse(self.result.exists())

    def test_return_to_existing_row_preserves_filter_and_workbook(self):
        self.refund({"B": 2})
        self.assertEqual(read_shared_uzk(self.result)[("p", "b")]["remaining"], 9)
        with ZipFile(self.spent) as old, ZipFile(self.result) as new:
            self.assertEqual(old.read(BOOK), new.read(BOOK))


if __name__ == "__main__":
    unittest.main()
