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
    def __init__(self, value):
        self.Value2=value
        self.HasFormula=False
        self.NumberFormat='General'
        self.NumberFormatLocal='Общий'

    @property
    def Text(self):
        if self.NumberFormatLocal == 'дд.мм.гггг':
            dt = date(1899,12,30) + timedelta(days=int(self.Value2))
            return dt.strftime('%d.%m.%Y')
        if self.NumberFormatLocal == 'чч:мм':
            total = round(float(self.Value2)*1440)
            return f'{total//60:02d}:{total%60:02d}'
        return str(self.Value2)


class Sheet:
    def __init__(self, values):
        self._cells={(i+1,1):Cell(value) for i,value in enumerate(values)}
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

    def test_excel_true_date_time_numeric_and_unformatted_text(self):
        cells=['{{AKZ_DATE}}','{{AKZ_LAYER_1_APPLICATION_DATE}}',
               '{{AKZ_LAYER_1_TIME_START}}','{{AKZ_LAYER_1_WET_AVG}}',
               'Акт АКЗ от {{AKZ_DATE}}']
        sheet=Sheet(cells)
        values={
            'AKZ_DATE':'19.10.2026',
            'AKZ_LAYER_1_APPLICATION_DATE':'2026-10-19',
            'AKZ_LAYER_1_TIME_START':'08:00',
            'AKZ_LAYER_1_WET_AVG':'130.4',
        }
        tokens_excel(sheet,values)
        expected=(date(2026,10,19)-date(1899,12,30)).days
        self.assertEqual(sheet.Cells(1,1).Value2,expected)
        self.assertEqual(sheet.Cells(2,1).Value2,expected)
        self.assertEqual(sheet.Cells(1,1).NumberFormatLocal,'дд.мм.гггг')
        self.assertEqual(sheet.Cells(1,1).NumberFormat,'General')
        self.assertEqual(sheet.Cells(1,1).Text,'19.10.2026')
        self.assertEqual(sheet.Cells(3,1).NumberFormatLocal,'чч:мм')
        self.assertEqual(sheet.Cells(3,1).Text,'08:00')
        self.assertAlmostEqual(sheet.Cells(3,1).Value2,1/3)
        self.assertAlmostEqual(sheet.Cells(4,1).Value2,130.4)
        self.assertEqual(sheet.Cells(5,1).Value2,'Акт АКЗ от 19.10.2026')

    def test_ru_excel_refuses_literal_english_time_after_format(self):
        class WrongVisibleCell(Cell):
            @property
            def Text(self):
                return 'hh:mm'  # reproduced broken ru-RU display

        sheet=Sheet(['{{AKZ_LAYER_1_TIME_START}}'])
        sheet._cells[(1,1)] = WrongVisibleCell('{{AKZ_LAYER_1_TIME_START}}')
        with self.assertRaisesRegex(Stop, 'неправильное отображение'):
            tokens_excel(sheet, {'AKZ_LAYER_1_TIME_START':'08:00'})

    def test_excel_format_com_failure_is_not_silenced(self):
        class FailingCell(Cell):
            @property
            def NumberFormatLocal(self):
                return 'Общий'
            @NumberFormatLocal.setter
            def NumberFormatLocal(self, format_code):
                if format_code!='Общий':
                    raise RuntimeError('Excel 1004')
        sheet=Sheet(['{{AKZ_DATE}}'])
        sheet._cells[(1,1)] = FailingCell('{{AKZ_DATE}}')
        with self.assertRaisesRegex(Stop, 'локальный формат'):
            tokens_excel(sheet, {'AKZ_DATE':'19.10.2026'})

    def test_ru_date_and_time_leave_neighboring_excel_formulas_untouched(self):
        sheet=Sheet(['{{AKZ_DATE}}', '=E22+1', '{{AKZ_LAYER_1_TIME_END}}'])
        sheet.Cells(2,1).HasFormula=True
        tokens_excel(sheet, {'AKZ_DATE':'2026-10-19',
                             'AKZ_LAYER_1_TIME_END':'10:30'})
        self.assertEqual(sheet.Cells(2,1).Value2, '=E22+1')
        self.assertEqual(sheet.Cells(3,1).Text, '10:30')
        self.assertAlmostEqual(sheet.Cells(3,1).Value2, 10.5/24)
        self.assertEqual(sheet.Cells(1,1).NumberFormatLocal, 'дд.мм.гггг')

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
