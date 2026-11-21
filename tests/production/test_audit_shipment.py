import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from passport_app.vom import read_vom
from production_app import catalog_ledger
from production_app.shipment_input import normalize, parse_tsv, write_internal
from production_app.storage import Stop, sha


class ShipmentLedger(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        technical = self.root / "source.xlsx"
        self.ledger = self.root / "ledger.xlsx"
        book = Workbook()
        sheet = book.active
        for row in (("Марка", "Профиль"), ("M", "P"), ("N", "P")):
            sheet.append(row)
        book.save(technical)
        book.close()
        self.old = {"M": 5}
        self.registered = [{"number": "2/2026", "date": date(2026, 11, 20)}]
        history = [{"number": "2/2026", "date": "2026-11-20",
                    "complete": True, "quantities": self.old}]
        catalog_ledger.create_migrated(technical, self.ledger, "P", history)
        book = load_workbook(self.ledger)
        sheet = book.worksheets[0]
        sheet["B2"], sheet["B3"] = 5, 20
        book.save(self.ledger)
        book.close()

    def amend(self, amount):
        candidate = self.root / "candidate.xlsx"
        old_sha = sha(self.ledger)
        catalog_ledger.shipment_candidate(
            self.ledger, candidate, "P", "2/2026", "2026-11-20", {"M": amount},
            self.registered, original_quantities=self.old, final_quantities={"M": amount})
        self.assertEqual(sha(self.ledger), old_sha)
        return candidate

    def test_replace_five_with_four_when_free_stock_is_zero(self):
        candidate = self.amend(4)
        book = load_workbook(candidate)
        try:
            self.assertEqual(book.worksheets[0]["F2"].value, 4)
            self.assertEqual(book.worksheets[0]["B2"].value, 5)
            self.assertEqual(catalog_ledger.metadata(book)["history"][0]["quantities"], {"M": 4})
        finally:
            book.close()

    def test_replacement_cannot_exceed_total(self):
        with self.assertRaises(Stop):
            self.amend(6)
        self.assertFalse((self.root / "candidate.xlsx").exists())

    def test_missing_previously_shipped_mark_blocks_another_issue(self):
        book = load_workbook(self.ledger)
        book.worksheets[0].delete_rows(2)
        book.save(self.ledger)
        book.close()
        before = sha(self.ledger)
        with self.assertRaisesRegex(Stop, "M"):
            catalog_ledger.validate_stock(self.ledger, "P", "3/2026", {"N": 1}, self.registered)
        self.assertEqual(sha(self.ledger), before)


class ClipboardReview(unittest.TestCase):
    def test_repeated_header_after_a_data_row_cannot_discard_first_shipment(self):
        text = ("М1\tБалка\t2\t10\t20\n"
                "Марка\tНаименование\tКоличество\tМасса\tВсего\n"
                "М2\tБалка\t1\t7\t7")
        with self.assertRaises(Stop):
            parse_tsv(text)

    def test_data_label_and_integer_decimal_survive_second_header_heuristic(self):
        text = ("Марка\tНаименование\tКоличество\tМасса\tВсего\n"
                "М1\tБалка (всего)\t2.0\t10\t20\n"
                "М2\tБалка\t1\t7\t7")
        self.assertEqual(parse_tsv(text).quantity, 3)

    def test_title_and_spanning_header(self):
        text = ("ВОМ\t\t\t\t\t\t\n"
                "Марка\tНаименование\tКоличество\tМасса\t\tПлощадь\t\n"
                "\t\t\t1 шт.\tВсего\t1 шт.\tВсего\n"
                "М1\tБалка\t2\t10\t20\t1\t2")
        self.assertEqual(parse_tsv(text).quantity, 2)


class VomScaling(unittest.TestCase):
    def test_optional_area_lookup_does_not_scan_sheet_for_each_row(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "v.xlsx"
            shipment = normalize([[f"M{i}", "Балка", 1, 2, 2] for i in range(50)])
            write_internal(shipment, source)
            width_reads = []
            original = Worksheet.max_column.fget

            def width(sheet):
                width_reads.append(1)
                return original(sheet)

            with patch.object(Worksheet, "max_column", property(width)):
                parsed = read_vom(source)
            self.assertEqual(parsed.quantity, 50)
            self.assertIsNone(parsed.area)
            self.assertLess(len(width_reads), 15)


if __name__ == "__main__":
    unittest.main()
