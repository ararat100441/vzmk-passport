import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

from openpyxl import Workbook
from production_app.amendment import control_plan
from production_app.shared_uzk import HEADERS
from production_app.storage import Stop, sha


class UzkReplacement(unittest.TestCase):
    def calculate(self, added, current):
        with tempfile.TemporaryDirectory() as folder:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "УЗК"
            sheet.append(HEADERS)
            if current:
                sheet.append(("P", "М1", current))
            path = Path(folder) / "uzk.xlsm"
            workbook.save(path)
            workbook.close()
            before = sha(path)
            delta = SimpleNamespace(rows=(("М1", "Балка", Decimal(added)),) if added else ())
            facts = {"marks": [{"mark": "М1", "uzk_quantity": 5}]}
            try:
                return control_plan(path, "P", delta, facts, {"М1": 5},
                                    {"М1": added} if added else {})
            finally:
                self.assertEqual(sha(path), before)

    def test_exhausted_mark_cannot_silently_zero_the_control_on_replacement(self):
        with self.assertRaisesRegex(Stop, "УЗК.*М1"):
            self.calculate(4, 0)

    def test_partial_overlap_cannot_silently_reduce_planned_control(self):
        with self.assertRaisesRegex(Stop, "УЗК.*М1"):
            self.calculate(4, 2)

    def test_already_covered_and_pure_refund_remain_available(self):
        plan, debit, refunds = self.calculate(4, 4)
        self.assertEqual(plan["М1"]["uzk_quantity"], 4)
        self.assertEqual(debit["М1"]["uzk_quantity"], 4)
        self.assertEqual(refunds, {"М1": 5})
        self.assertEqual(self.calculate(0, 0), ({}, {}, {"М1": 5}))


if __name__ == "__main__":
    unittest.main()
