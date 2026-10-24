"""Regression: original one-line shipping VOM format with grouped header.

Synthetic workbook objects only: the customer's original workbook is never
committed or modified. Legacy two-tier document templates remain untouched.
"""
import unittest
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from passport_app.config import PassportError
from passport_app.tables import find_layout,locate_layout
from passport_app.vom import read_vom,validate_flat_shipping_amounts


class Cell:
    def __init__(self,row,column,value):
        self.row=row
        self.column=column
        self.value=value
        self.coordinate='R'+str(row)+'C'+str(column)


class Merge:
    def __init__(self,r1,c1,r2,c2):
        self.min_row=r1;self.min_col=c1
        self.max_row=r2;self.max_col=c2


class FakeSheet:
    title='Лист1'
    def __init__(self,header=None,*,legacy=False,totals=False):
        if header is None:
            header=['Лист № ','Отправочная марка ','Наименование марки ',
                    'Количество ','Масса (изделие), кг ',None,
                    'Площадь окраски, м2 ',None]
        self.values=[header]
        if legacy:
            self.values.append([None,None,None,None,'1 шт','Всех','1 шт','Всех'])
        self.values.extend([
            [' 3.33','TEST-M1','Балка',8,50.84,406.72,1.943,15.544],
            [' 3.204','TEST-M2','Настил',45,87.1,3919.4999999999995,6.974,313.83],
        ])
        if totals:
            self.values.append(['Итого',None,None,53,None,4326.22,None,329.374])
        self.max_row=len(self.values)
        self.max_column=8
        self.merged_cells=SimpleNamespace(ranges=[
            Merge(1,5,1,6),Merge(1,7,1,8),
        ] + ([Merge(1,c,2,c) for c in (2,3,4)] if legacy else []))

    def cell(self,row,column):
        return Cell(row,column,self.values[row-1][column-1])

    def __getitem__(self,row):
        return tuple(self.cell(row,c) for c in range(1,9))

    def iter_rows(self,min_row=1,max_row=None):
        stop=min(self.max_row,max_row or self.max_row)
        return (self[r] for r in range(min_row,stop+1))


class FakeBook:
    def __init__(self,sheet):
        self.worksheets=[sheet]
        self._sheet=sheet
    def __getitem__(self,key):
        assert key==self._sheet.title
        return self._sheet
    def close(self):
        pass


class FlatShippingVOM(unittest.TestCase):
    def test_realistic_shipping_header_accepts_all_seven_fields(self):
        parsed=find_layout(FakeSheet())
        self.assertTrue(parsed.flat_shipping_header)
        self.assertEqual(parsed.start,2)
        self.assertEqual(parsed.header_row,1)
        self.assertEqual(parsed.columns,{
            'mark':2,'name':3,'quantity':4,
            'mass_unit':5,'mass_total':6,
            'area_unit':7,'area_total':8,
        })
        self.assertEqual(parsed.total_row,None)

    def test_locate_unique_accepted_sheet(self):
        p=locate_layout(FakeBook(FakeSheet()))
        self.assertEqual(p.sheet,'Лист1')
        self.assertEqual(p.start,2)

    def test_legacy_two_tier_grouped_template_still_works(self):
        legacy=FakeSheet(
            header=['Лист №','Отправочная марка','Наименование марки',
                    'Количество','Масса, кг',None,'Площадь',None],
            legacy=True,totals=True,
        )
        layout=find_layout(legacy,appendix=True)
        self.assertFalse(layout.flat_shipping_header)
        self.assertEqual(layout.start,3)
        self.assertEqual(layout.total_row,5)

    def test_flat_invoice_is_not_accepted_as_appendix_template(self):
        with self.assertRaises(PassportError):
            find_layout(FakeSheet(),appendix=True)

    def test_fallback_requires_two_exact_horizontal_merges(self):
        sample=FakeSheet()
        sample.merged_cells.ranges.pop()
        with self.assertRaisesRegex(PassportError,'объединённая шапка'):
            find_layout(sample)

    def test_fallback_refuses_an_unexpected_mass_or_area_caption(self):
        header=['Лист №','Отправочная марка','Наименование марки',
                'Количество','Нет массы',None,'Площадь окраски, м2',None]
        with self.assertRaises(PassportError):
            find_layout(FakeSheet(header=header))

    def test_shipping_row_reconciles_units_and_totals(self):
        validate_flat_shipping_amounts(
            Decimal(8),Decimal('50.84'),Decimal('406.72'),
            Decimal('1.943'),Decimal('15.544'),2,
        )
        # Floating calculation artifacts must not reject actual Excel totals.
        validate_flat_shipping_amounts(
            Decimal(45),Decimal('87.1'),
            Decimal('3919.4999999999995'),
            Decimal('6.974'),Decimal('313.83'),3,
        )

    def test_flat_shipping_wrong_amounts_rejected(self):
        with self.assertRaisesRegex(PassportError,'Масса'):
            validate_flat_shipping_amounts(
                Decimal(8),Decimal('50.84'),Decimal('407'),
                Decimal('1.943'),Decimal('15.544'),2,
            )
        with self.assertRaisesRegex(PassportError,'Площадь'):
            validate_flat_shipping_amounts(
                Decimal(8),Decimal('50.84'),Decimal('406.72'),
                Decimal('1.943'),Decimal('15.500'),2,
            )

    def test_read_vom_full_pipeline_with_two_rows(self):
        fake=FakeSheet()
        books=[FakeBook(fake),FakeBook(fake)]
        with patch('passport_app.vom.sha256',return_value='UNCHANGED'):
            with patch('passport_app.vom.load_book',side_effect=books):
                data=read_vom(Path('synthetic.xlsx'))
        self.assertEqual(len(data.rows),2)
        self.assertEqual(data.quantity,Decimal('53'))
        self.assertAlmostEqual(float(data.mass),4326.22,places=2)
        self.assertEqual(data.connections,'Балка, Настил')
        self.assertTrue(data.layout.flat_shipping_header)

    def test_read_vom_rejects_wrong_flat_shipping_total(self):
        bad=FakeSheet()
        bad.values[1][5]=999
        with patch('passport_app.vom.sha256',return_value='UNCHANGED'):
            with patch('passport_app.vom.load_book',
                       side_effect=[FakeBook(bad),FakeBook(bad)]):
                with self.assertRaisesRegex(PassportError,'Масса'):
                    read_vom(Path('synthetic.xlsx'))


if __name__=='__main__':
    unittest.main(verbosity=2)
