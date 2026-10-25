"""VZMK manufacturing START_DATE: same as first passport, not new control date.

All registries here are synthetic and only temp copies are written.
H353 reproduces the approved Garage Kazakhstan first passport layout.
"""
import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

from production_app.registry_dates import first_passport_start
from production_app.storage import Stop, sha


def synthetic_registry(filename,start):
    book=Workbook()
    sheet=book.active
    sheet.title='Паспорт качества'
    sheet['H353']=start
    sheet['N353']=date(2026,9,9)  # First passport ending date; NOT start.
    sheet['H354']=date(2026,10,25) # A later passport's wrong start ignored.
    book.save(filename)
    book.close()


class FirstStartDate(unittest.TestCase):
    def test_literal_h353_is_the_valid_project_start_for_all_later_passports(self):
        with TemporaryDirectory() as td:
            path=Path(td)/'COPY_ВЗМК.xlsx'
            synthetic_registry(path,date(2026,6,3))
            before=sha(path)
            registry=SimpleNamespace(path=path,sheet='Паспорт качества')
            project=SimpleNamespace(first_row=353,first_number='1/2026')
            with patch('production_app.registry_dates.model.columns',
                       return_value={'START_DATE':8}):
                result=first_passport_start(registry,project)
            self.assertEqual(result,'2026-06-03')
            # Neither first END_DATE nor a later START_DATE was used.
            self.assertNotEqual(result,'2026-09-09')
            self.assertNotEqual(result,'2026-10-25')
            self.assertEqual(sha(path),before)

    def test_missing_first_start_does_not_fallback_to_end_or_today(self):
        with TemporaryDirectory() as td:
            path=Path(td)/'COPY_ВЗМК.xlsx'
            synthetic_registry(path,None)
            before=sha(path)
            registry=SimpleNamespace(path=path,sheet='Паспорт качества')
            project=SimpleNamespace(first_row=353,first_number='1/2026')
            with patch('production_app.registry_dates.model.columns',
                       return_value={'START_DATE':8}):
                with self.assertRaisesRegex(Stop,'начала изготовления'):
                    first_passport_start(registry,project)
            self.assertEqual(sha(path),before)

    def test_no_cross_project_fallback_when_exact_first_row_is_missing(self):
        with TemporaryDirectory() as td:
            path=Path(td)/'COPY_ВЗМК.xlsx'
            synthetic_registry(path,None)
            # Other object's first start must not be borrowed.
            from openpyxl import load_workbook
            book=load_workbook(path)
            book.active['H341']=date(2026,1,20)
            book.save(path);book.close()
            registry=SimpleNamespace(path=path,sheet='Паспорт качества')
            project=SimpleNamespace(first_row=353)
            with patch('production_app.registry_dates.model.columns',
                       return_value={'START_DATE':8}):
                with self.assertRaisesRegex(Stop,'начала изготовления'):
                    first_passport_start(registry,project)


if __name__=='__main__':
    unittest.main(verbosity=2)
