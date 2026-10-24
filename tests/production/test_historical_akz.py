"""Synthetic three-layer АКЗ acceptance: no historical measurements are reused."""
import unittest
from datetime import date, timedelta
from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import patch

from production_app.facts import validate_layer_details
from production_app.office import akz_values, tokens_excel
from production_app.real_forms import require_akz_form_inputs
from production_app.storage import Stop


def facts():
    return {'start_date':'2026-10-18','date':'2026-10-20',
            'akz_date':'2026-10-19','doc_fields':{'AKZ_HYGROMETER':'NEW VERIFIED DEVICE'}}


def layer_details():
    return [{
        'layer':i,'material':'SYNTHETIC COAT '+str(i),
        'paint_batch':'P-'+str(i),'hardener_batch':'H-'+str(i),'color':'RAL 0000',
        'application_date':'2026-10-19',
        'time_start':'08:00','time_end':'10:30','wet_average':str(110*i),
    } for i in (1,2,3)]


def climate():
    return {
        'variant':0,'air':'24','surface':'27','humidity':'55','dew':'14',
        'layers':[{'layer':i,'min':'80','max':'130'} for i in (1,2,3)],
        'dry_averages':[{'layer':1,'average':'110'},
                        {'layer':2,'average':'220'},
                        {'layer':3,'average':'330'}],
        'layer_details':layer_details(),
    }


class Cell:
    def __init__(self, value, *, native_style=None):
        self.Value2=value
        self.HasFormula=False
        self.NumberFormat=('BuiltinDate14' if native_style=='date14'
                           else 'BuiltinTime20' if native_style=='time20'
                           else 'General')
        self.NumberFormatLocal=self.NumberFormat

    @property
    def Text(self):
        if self.NumberFormat=='BuiltinDate14':
            dt=date(1899,12,30)+timedelta(days=int(self.Value2))
            return dt.strftime('%d.%m.%Y')
        if self.NumberFormat=='BuiltinTime20':
            minutes=round(float(self.Value2)*1440)
            return f'{minutes//60}:{minutes%60:02d}'
        if self.NumberFormatLocal=='дд.мм.гггг':
            dt=date(1899,12,30)+timedelta(days=int(self.Value2))
            return dt.strftime('%d.00.%Y')  # PR10/PR11 actual native failure
        return str(self.Value2)


class Sheet:
    def __init__(self, values, *, native_styles=None):
        native_styles=native_styles or {}
        self._cells={(i+1,1):Cell(value,native_style=native_styles.get(i+1))
                     for i,value in enumerate(values)}
        self.UsedRange=SimpleNamespace(Row=1,Column=1,
                                       Rows=SimpleNamespace(Count=len(values)),
                                       Columns=SimpleNamespace(Count=1))
    def Cells(self,row,col):
        return self._cells[(row,col)]


class HistoricalAKZ(unittest.TestCase):
    def test_all_three_independent_coats(self):
        d=validate_layer_details(layer_details(),facts())
        self.assertEqual([x['layer'] for x in d],[1,2,3])
        self.assertEqual(d[2]['paint_batch'],'P-3')

    def test_cross_day_or_late_day_requires_new_climate(self):
        for changes in ({'application_date':'2026-10-18'},
                        {'application_date':'2026-10-21'},
                        {'time_start':'10:00','time_end':'08:00'},
                        {'wet_average':'0'}):
            with self.subTest(changes=changes):
                rows=layer_details();rows[1].update(changes)
                with self.assertRaises(Stop):
                    validate_layer_details(rows,facts())

    def test_missing_or_duplicate_layer_refused(self):
        with self.assertRaises(Stop):
            validate_layer_details(layer_details()[:2],facts())
        bad=layer_details();bad[-1]['layer']=2
        with self.assertRaises(Stop):
            validate_layer_details(bad,facts())

    def test_spreadsheet_values_use_current_batches_and_measured_means(self):
        values=akz_values(climate(),facts())
        self.assertEqual(values['AKZ_LAYER_1_MATERIAL'],'SYNTHETIC COAT 1')
        self.assertEqual(values['AKZ_LAYER_2_PAINT_BATCH'],'P-2')
        self.assertEqual(values['AKZ_LAYER_3_WET_AVG'],'330')
        self.assertEqual(values['LAYER_3_DRY_AVG'],'330')
        self.assertEqual(values['LAYER_3_TOTAL_RANGE'],'240–390')
        self.assertEqual(values['AIR'],'24')

    def test_excel_native_date_time_serials_preserve_source_formats(self):
        # Original 3-coat Excel uses builtin numFmt 14 for H4/E22 and
        # builtin numFmt 20 for E23. Do not set mask strings from Python.
        cells=['{{AKZ_DATE}}','{{AKZ_LAYER_1_APPLICATION_DATE}}',
               '{{AKZ_LAYER_1_TIME_START}}','{{AKZ_LAYER_1_WET_AVG}}',
               'Акт АКЗ от {{AKZ_DATE}}']
        sheet=Sheet(cells,native_styles={1:'date14',2:'date14',3:'time20'})
        values={
            'AKZ_DATE':'19.10.2026',
            'AKZ_LAYER_1_APPLICATION_DATE':'2026-10-19',
            'AKZ_LAYER_1_TIME_START':'08:00',
            'AKZ_LAYER_1_WET_AVG':'130.4',
        }
        formats=[(sheet.Cells(row,1).NumberFormat,
                  sheet.Cells(row,1).NumberFormatLocal) for row in (1,2,3)]
        tokens_excel(sheet,values)
        serial=(date(2026,10,19)-date(1899,12,30)).days
        self.assertEqual(sheet.Cells(1,1).Value2,serial)
        self.assertEqual(sheet.Cells(2,1).Value2,serial)
        self.assertEqual(sheet.Cells(1,1).Text,'19.10.2026')
        self.assertEqual(sheet.Cells(2,1).Text,'19.10.2026')
        self.assertEqual(sheet.Cells(3,1).Text,'8:00')  # legitimate native h:mm
        self.assertAlmostEqual(sheet.Cells(3,1).Value2,1/3)
        self.assertAlmostEqual(sheet.Cells(4,1).Value2,130.4)
        self.assertEqual(sheet.Cells(5,1).Value2,'Акт АКЗ от 19.10.2026')
        self.assertEqual(formats,[(sheet.Cells(row,1).NumberFormat,
                      sheet.Cells(row,1).NumberFormatLocal) for row in (1,2,3)])

    def test_literal_english_time_mask_refused_without_modification(self):
        class WrongVisibleCell(Cell):
            @property
            def Text(self):
                return 'hh:mm'
        sheet=Sheet(['{{AKZ_LAYER_1_TIME_START}}'])
        invalid=WrongVisibleCell('{{AKZ_LAYER_1_TIME_START}}',native_style='time20')
        sheet._cells[(1,1)]=invalid
        with self.assertRaisesRegex(Stop,'корректного встроенного формата'):
            tokens_excel(sheet, {'AKZ_LAYER_1_TIME_START':'08:00'})
        self.assertEqual(invalid.NumberFormat,'BuiltinTime20')

    def test_regression_month_zero_refused_not_masked(self):
        sheet=Sheet(['{{AKZ_DATE}}'])
        cell=sheet.Cells(1,1)
        cell.NumberFormatLocal='дд.мм.гггг'
        with self.assertRaisesRegex(Stop,'19.00.2026'):
            tokens_excel(sheet,{'AKZ_DATE':'19.10.2026'})
        self.assertEqual(cell.NumberFormat,'General')
        self.assertEqual(cell.NumberFormatLocal,'дд.мм.гггг')
        self.assertEqual(cell.Value2,(date(2026,10,19)-date(1899,12,30)).days)

    def test_builtin_date_format_remains_unchanged(self):
        sheet=Sheet(['{{AKZ_DATE}}'],native_styles={1:'date14'})
        cell=sheet.Cells(1,1)
        tokens_excel(sheet, {'AKZ_DATE':'19.10.2026'})
        self.assertEqual(cell.NumberFormat,'BuiltinDate14')
        self.assertEqual(cell.NumberFormatLocal,'BuiltinDate14')
        self.assertEqual(cell.Text,'19.10.2026')

    def test_excel_date_format_com_read_error_is_not_silenced(self):
        class FailingCell(Cell):
            @property
            def NumberFormatLocal(self):
                raise RuntimeError('Excel 1004')
            @NumberFormatLocal.setter
            def NumberFormatLocal(self, fmt):
                pass
        sheet=Sheet(['{{AKZ_DATE}}'])
        sheet._cells[(1,1)]=FailingCell('{{AKZ_DATE}}')
        with self.assertRaisesRegex(Stop,'не удалось прочитать исходный формат'):
            tokens_excel(sheet, {'AKZ_DATE':'19.10.2026'})

    def test_date_in_unformatted_cell_is_refused(self):
        sheet=Sheet(['{{AKZ_DATE}}'])
        with self.assertRaisesRegex(Stop,'исходная ячейка'):
            tokens_excel(sheet,{'AKZ_DATE':'19.10.2026'})

    def test_excel_date_time_leave_neighboring_formulas_untouched(self):
        sheet=Sheet(['{{AKZ_DATE}}','=E22+1','{{AKZ_LAYER_1_TIME_END}}'],
                    native_styles={1:'date14',3:'time20'})
        sheet.Cells(2,1).HasFormula=True
        tokens_excel(sheet,{'AKZ_DATE':'2026-10-19',
                            'AKZ_LAYER_1_TIME_END':'10:30'})
        self.assertEqual(sheet.Cells(2,1).Value2,'=E22+1')
        self.assertEqual(sheet.Cells(3,1).Text,'10:30')
        self.assertAlmostEqual(sheet.Cells(3,1).Value2,10.5/24)
        self.assertEqual(sheet.Cells(1,1).NumberFormat,'BuiltinDate14')

    def test_three_layer_act_require_input_fields_before_copy(self):
        roles={'akz0':{'path':'test.xlsx'}}
        profile={'templates':roles}
        cells=[SimpleNamespace(value='{{AKZ_LAYER_1_MATERIAL}}'),
               SimpleNamespace(value='{{AKZ_HYGROMETER}}')]
        book=SimpleNamespace(worksheets=[list(cells)],active=[list(cells)],close=lambda:None)
        with patch('passport_app.excel.load_book',return_value=book):
            with self.assertRaisesRegex(Stop,'сведения каждого'):
                require_akz_form_inputs(profile,{**facts(),'climates':[dict(climate(),layer_details=[])]},[0])
            self.assertTrue(require_akz_form_inputs(
                profile,{**facts(),'climates':[climate()]},[0]))

    def test_mocked_excel_loader_does_not_import_templates_module(self):
        # Regression for the 92 PASS / 1 FAIL combined Windows suite:
        # importing production_app.templates under a mocked loader persisted
        # the MagicMock as a module-level dependency across later tests.
        import sys
        profile={'templates':{'akz0':{'path':'test.xlsx'}}}
        cells=[SimpleNamespace(value='{{AKZ_HYGROMETER}}')]
        book=SimpleNamespace(worksheets=[list(cells)],active=[list(cells)],close=lambda:None)
        with patch.dict(sys.modules,{'production_app.templates':None}):
            with patch('passport_app.excel.load_book',return_value=book):
                self.assertTrue(require_akz_form_inputs(
                    profile,{**facts(),'climates':[climate()]},[0]))

    def test_missing_device_approval_refused(self):
        profile={'templates':{'akz0':{'path':'test.xlsx'}}}
        cells=[SimpleNamespace(value='{{AKZ_HYGROMETER}}')]
        book=SimpleNamespace(worksheets=[list(cells)],active=[list(cells)],close=lambda:None)
        with patch('passport_app.excel.load_book',return_value=book):
            with self.assertRaisesRegex(Stop,'термогигрометр'):
                require_akz_form_inputs(profile,{**facts(),'doc_fields':{},'climates':[climate()]},[0])


if __name__=='__main__':
    unittest.main(verbosity=2)
