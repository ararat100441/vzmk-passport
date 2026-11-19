"""VZMK manufacturing START_DATE: same as first passport, not new control date.

All registries here are synthetic and only temp copies are written.
H353 reproduces the approved Garage Kazakhstan first passport layout.
"""
import unittest
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook,load_workbook

from production_app.registry_dates import first_passport_start,resolve_registry_dates
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
    def test_gui_preflight_accepts_fixed_start_but_refuses_old_control_date(self):
        from production_app.registry_dates import require_new_date_formulas
        with TemporaryDirectory() as td:
            path=Path(td)/'COPY.xlsx'
            w=Workbook();s=w.active;s.title='Data'
            s['H353']=date(2026,6,3)
            for key in ('J','K','L'):s[key+'353']='=I353'
            w.save(path);w.close()
            registry=SimpleNamespace(path=path,sheet='Data')
            project=SimpleNamespace(first_row=353)
            columns={'START_DATE':8,'VIK_DATE':10,'UZK_DATE':11,'AKZ_DATE':12}
            before=sha(path)
            with patch('production_app.registry_dates.model.columns',return_value=columns):
                require_new_date_formulas(path,registry,project)
            self.assertEqual(sha(path),before)
            w=load_workbook(path);w['Data']['J353']=date(2026,6,3);w.save(path);w.close()
            before=sha(path)
            with patch('production_app.registry_dates.model.columns',return_value=columns):
                with self.assertRaisesRegex(Stop,'J353.*VIK_DATE'):
                    require_new_date_formulas(path,registry,project)
            self.assertEqual(sha(path),before)

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


    def test_new_passport_refuses_mismatched_start_on_disposable_copy(self):
        with TemporaryDirectory() as td:
            path=Path(td)/'COPY_ВЗМК.xlsx'
            book=Workbook()
            sheet=book.active
            sheet.title='Паспорт качества'
            sheet['H353']=date(2026,5,14)
            sheet['I353']=date(2026,6,1)
            for key in ('J','K','L'):
                sheet[key+'353']='=I353'  # true current-date formulas
            book.save(path);book.close()
            original=sha(path)
            fields={
                'START_DATE':8,'END_DATE':9,'VIK_DATE':10,
                'UZK_DATE':11,'AKZ_DATE':12
            }
            project=SimpleNamespace(first_row=353,next_number='11/2026')

            def fake_registry(file):
                w=load_workbook(file,data_only=False)
                try:
                    present=w['Паспорт качества']['I354'].value is not None
                finally:w.close()
                entries=[{'row':353,'number':'1/2026'}]
                if present:entries.append({'row':354,'number':'11/2026'})
                return SimpleNamespace(path=Path(file),
                    sheet='Паспорт качества',digest=sha(file),
                    projects={'Гараж Казахстан':entries})

            @contextmanager
            def no_native_excel():
                yield object()

            def bad_copy(excel,copy,registry,selected,vom,day):
                w=load_workbook(copy)
                try:
                    s=w['Паспорт качества']
                    s['H354']=date(2026,10,1)  # invalid new manufacturing start
                    for key in ('I','J','K','L'):
                        s[key+'354']=day
                    w.save(copy)
                finally:w.close()
                return 354,{}
            with patch('production_app.registry_dates.read_registry',
                       side_effect=fake_registry),\
                 patch('production_app.registry_dates.model.select',
                       return_value=project),\
                 patch('production_app.registry_dates.model.columns',
                       return_value=fields),\
                 patch('production_app.registry_dates.excel_session',
                       side_effect=no_native_excel),\
                 patch('production_app.registry_dates.append_row',
                       side_effect=bad_copy):
                with self.assertRaisesRegex(
                        Stop,'не совпала.*первого паспорта'):
                    resolve_registry_dates(
                        {'registry':str(path)},'Гараж Казахстан',
                        '11/2026',None,True,day=date(2026,10,25))
            self.assertEqual(sha(path),original)
            w=load_workbook(path)
            try:
                self.assertIsNone(w['Паспорт качества']['I354'].value)
            finally:w.close()



if __name__=='__main__':
    unittest.main(verbosity=2)
