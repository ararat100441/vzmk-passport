"""Focused review regressions: streaming hashes, UZK quantities and row shifts.

Use a private temp directory. Never touch owner workbooks. Native Microsoft
Office and complete frozen GUI release remain separate acceptance gates.
"""
from __future__ import annotations

import hashlib
import random
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

from openpyxl import Workbook

from passport_app.excel import sha256 as excel_sha256
from production_app.shared_uzk import (
    decrement_candidate, read_shared_uzk, verify_planned_quantities,
)
from production_app.storage import Stop, sha


class StreamedHash(unittest.TestCase):
    def test_both_hashers_match_sha256_for_empty_and_large_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "source.bin"
            for data in (
                b"",
                b"\0",
                b"ABC" * 43,
                random.Random(73).randbytes(3 * 1024 * 1024 + 17),
            ):
                path.write_bytes(data)
                expected = hashlib.sha256(data).hexdigest()
                self.assertEqual(sha(path), expected)
                self.assertEqual(excel_sha256(path), expected)

    def test_missing_workbook_preserves_original_error_contract(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(OSError):
                sha(Path(tmp) / "missing.xlsx")
            with self.assertRaisesRegex(Exception, "Не удалось прочитать"):
                excel_sha256(Path(tmp) / "missing.xlsx")


class PlannedUzkQuantity(unittest.TestCase):
    def test_missing_or_boolean_quantity_does_not_become_one_or_zero(self):
        for expected, value in ((1, None), (0, False), (1, True), (1, "")):
            with self.subTest(expected=expected, value=value):
                facts = {"marks": [{"mark": "M"}]}
                if value is not None:
                    facts["marks"][0]["uzk_quantity"] = value
                with self.assertRaises(Stop):
                    verify_planned_quantities(
                        facts, {"M": {"uzk_quantity": expected}}
                    )

    def test_actual_zero_and_positive_counts_remain_valid(self):
        for quantity in (0, "0", 1, "1", 3, "3"):
            with self.subTest(quantity=quantity):
                self.assertTrue(
                    verify_planned_quantities(
                        {"marks": [{"mark": "M", "uzk_quantity": quantity}]},
                        {"M": {"uzk_quantity": int(quantity)}},
                    )
                )

    def test_duplicate_mark_is_blocked(self):
        facts = {
            "marks": [
                {"mark": "M", "uzk_quantity": 1},
                {"mark": "M", "uzk_quantity": 1},
            ]
        }
        with self.assertRaisesRegex(Stop, "Повторы"):
            verify_planned_quantities(
                facts, {"M": {"uzk_quantity": 1}}
            )


class UzkDataFile(unittest.TestCase):
    def make_xlsm(self, path: Path, rows):
        book = Workbook()
        sheet = book.active
        sheet.title = "УЗК"
        sheet.append(["Заказ", "Марка элемента", "Кол-во (шт)"])
        for row in rows:
            sheet.append(row)
        book.save(path)
        book.close()

    def test_generator_mark_selection_is_not_exhausted_on_first_row(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "input.xlsm"
            self.make_xlsm(source, [
                ("PROJECT", "M", 1),
                (None, "N", 2),  # Same selected mark, invalid project.
            ])
            with self.assertRaisesRegex(Stop, "заказ/марка"):
                read_shared_uzk(
                    source, project="PROJECT", marks=(m for m in ("M", "N"))
                )

    def test_shift_many_rows_keeps_other_project_and_opaque_zip_bytes(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "source.xlsm"
            dst = Path(tmp) / "candidate.xlsm"
            rows = [
                ("PROJECT", f"M{number:03d}", 1)
                for number in range(1, 121)
            ] + [
                ("OTHER", f"N{number:03d}", 4)
                for number in range(1, 121)
            ]
            self.make_xlsm(src, rows)
            with ZipFile(src, "a") as archive:
                archive.writestr("customXml/sentinel.bin", b"DO_NOT_TOUCH")
            before = sha(src)
            counts = {f"M{number:03d}": 1 for number in range(1, 121)}
            result = decrement_candidate(src, dst, "PROJECT", counts)
            self.assertEqual(len(result["deleted_rows"]), 120)
            self.assertEqual(sha(src), before)
            with ZipFile(src) as original, ZipFile(dst) as updated:
                self.assertEqual(
                    original.read("customXml/sentinel.bin"),
                    updated.read("customXml/sentinel.bin"),
                )
            target = read_shared_uzk(dst)
            self.assertEqual(len(target), 120)
            self.assertNotIn(("project", "m001"), target)
            self.assertEqual(target[("other", "n001")]["remaining"], 4)
            self.assertEqual(target[("other", "n120")]["remaining"], 4)


if __name__ == "__main__":
    unittest.main()
