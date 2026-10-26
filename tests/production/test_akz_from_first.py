"""Pure-Python fake native Excel COM test: NEVER writes a spreadsheet.

The user's original AKZ has E27=E28/0.6 (formula in yellow), E22=H4-1,
E30=E22 and D49=H4; A1 of appendix and ALL data rows change with VOM.
"""
import unittest
from unittest.mock import patch
from production_app.akz_from_first import (
    is_first_akz,is_first_akz_appendix,
    fill_first_akz,fill_first_akz_appendix,
)
from production_app.storage import Stop


class Object:
    pass


class Ref:
    def __init__(self,s,name):
        self.s=s;self.name=name
        self.Interior=Object()
        self.Interior.Pattern=None
        self.Borders=Object()
        self.Font=Object()
        self.Formula=None
        self._number='dd.mm.yyyy'
        self.NumberFormat=self._number
        self.NumberFormatLocal='ДД.ММ.ГГГГ'
    @property
    def Value2(self):
        return self.s.data.get(self.name)
    @Value2.setter
    def Value2(self,value):
        self.s.data[self.name]=value
    def FillDown(self):self.s.operations.append(('FillDown',self.name))
    def Insert(self):self.s.operations.append(('Insert',self.name))
    def Delete(self):self.s.operations.append(('Delete',self.name))
    @property
    def Hyperlinks(self):
        return self
    def __getattr__(self,attr):
        if attr=='RowHeight':return 28
        raise AttributeError(attr)


class Cells:
    def __init__(self,s,r,c):
        self.s=s;self.r=r;self.c=c
        self.key=chr(64+c)+str(r)
        self.Hyperlinks=self
    @property
    def Value2(self):
        if self.key in self.s.data:return self.s.data[self.key]
        if self.c==1 and 4<=self.r<=20:return 'К1-'+str(self.r)
        if self.c==1 and self.r==21:return 'Итого'
        return None
    @Value2.setter
    def Value2(self,value):
        self.s.data[self.key]=value
    def Delete(self):pass


class Sheet:
    def __init__(self):
        self.data={}
        self.refs={}
        self.operations=[]
        self.UsedRange=Object()
        self.UsedRange.Row=1
        self.UsedRange.Rows=Object()
        self.UsedRange.Rows.Count=21
        self.PageSetup=Object()
    def Range(self,key):
        if key not in self.refs:self.refs[key]=Ref(self,key)
        return self.refs[key]
    def Cells(self,r,c):return Cells(self,r,c)
    def Rows(self,s):return self.Range('ROWS:'+str(s))


AKZ_FORMULAS={
    'G8':'=E4','I8':'=H4','E22':'=H4-1',
    'E27':'=E28/0.6','E30':'=E22','D49':'=H4'
}


class FirstAKZContract(unittest.TestCase):
    def test_first_act_shape_requires_all_formulas(self):
        values={'E4':{'value':'1/2026 от '}}
        values.update({k:{'formula':v[1:]} for k,v in AKZ_FORMULAS.items()})
        values.update({k:{'value':'1'} for k in ('H4','E28','E31','E32','E33','E34')})
        with patch('production_app.akz_from_first._cells',return_value=values):
            self.assertTrue(is_first_akz('dummy.xlsx'))
            without=values.copy()
            without['E27']={'value':'154'}
            with patch('production_app.akz_from_first._cells',return_value=without):
                self.assertFalse(is_first_akz('dummy.xlsx'))

    def test_first_appendix_shape_requires_current_table(self):
        data={
            'A1':{'value':'Приложение № 1/2026 к акту АКЗ № 1/2026  от 08.09.2026 г.'},
            'A2':{'value':'Отправочная марка'},
            'B2':{'value':'Наименование марки '},
            'A4':{'value':'К1-5'},
            'A21':{'value':'Итого'},
        }
        with patch('production_app.akz_from_first._cells',return_value=data):
            self.assertTrue(is_first_akz_appendix('dummy.xlsx'))
            broken=dict(data,A21={'value':'Сумма'})
            with patch('production_app.akz_from_first._cells',return_value=broken):
                self.assertFalse(is_first_akz_appendix('dummy.xlsx'))

    def test_climate_updates_and_original_formulas_survive(self):
        s=Sheet()
        s.data['E4']='1/2026 от '
        for pos,formula in AKZ_FORMULAS.items():
            s.Range(pos).Formula=formula
        facts={'variant':0,'thickness':'91.2',
               'air':'-2.5','surface':'-3.5','humidity':'53',
               'dew':'-11.2'}
        result=fill_first_akz(
            s,number='11/2026',akz_date='2026-10-22',climate=facts)
        self.assertEqual(s.Range('E4').Value2,'11/2026 от ')
        self.assertEqual(s.Range('E28').Value2,91.2)
        self.assertEqual(s.Range('E32').Value2,-2.5)
        self.assertEqual(s.Range('E33').Value2,53.)
        self.assertEqual(s.Range('E27').Formula,'=E28/0.6')
        self.assertEqual(s.Range('E22').Formula,'=H4-1')
        self.assertEqual(s.Range('D49').Formula,'=H4')
        self.assertTrue(result['wet_film_formula_preserved'])

    def test_missing_data_does_not_persist_any_excel_file(self):
        s=Sheet()
        s.data['E4']='1/2026 от '
        for pos,formula in AKZ_FORMULAS.items():s.Range(pos).Formula=formula
        with self.assertRaisesRegex(Stop,'измеренного значения'):
            fill_first_akz(
                s,number='11/2026',akz_date='2026-10-22',
                climate={'variant':0,'thickness':'','air':'22',
                         'surface':'20','humidity':'55','dew':'12'})

    def test_appendix_four_current_rows_and_current_totals(self):
        s=Sheet()
        s.data.update({
            'A1':'Приложение № 1/2026 к акту АКЗ № 1/2026  от 08.09.2026 г.',
            'A2':'Отправочная марка','B2':'Наименование марки ',
        })
        rows=[
            ('Бп2-1','Балка',2,10.,20.,3.,6.),
            ('Бп2-5','Балка',3,11.,33.,4.,12.),
            ('Бп2-18','Балка',4,12.,48.,5.,20.),
            ('Бп2-19','Балка',2,13.,26.,6.,12.),
        ]
        result=fill_first_akz_appendix(
            s,number='11/2026',akz_date='2026-10-22',rows=rows)
        self.assertEqual(result['total_row'],8)
        self.assertEqual(s.Range('A4:G4').Value2,(rows[0],))
        self.assertEqual(s.Range('A7:G7').Value2,(rows[-1],))
        self.assertEqual(s.Range('C8').Formula,'=SUM(C3:C7)')
        self.assertEqual(s.Range('E8').Formula,'=SUM(E3:E7)')
        self.assertEqual(s.Range('G8').Formula,'=SUM(G3:G7)')
        self.assertEqual(s.PageSetup.PrintTitleRows,'$2:$3')
        self.assertEqual(s.PageSetup.PrintArea,'$A$1:$G$8')
        self.assertEqual(s.Cells(8,1).Value2,'Итого')
        self.assertEqual(s.Range('A1').Value2,
            'Приложение № 11/2026 к акту АКЗ № 11/2026  от 22.10.2026 г.')

    def test_number_and_date_cannot_be_invented(self):
        s=Sheet()
        with self.assertRaisesRegex(Stop,'дата'):
            fill_first_akz_appendix(
                s,number='11/2026',akz_date=None,rows=[1])
        with self.assertRaisesRegex(Stop,'номер'):
            fill_first_akz(
                s,number='fake',akz_date='2026-10-22',
                climate={'variant':0})


if __name__=='__main__':unittest.main(verbosity=2)
