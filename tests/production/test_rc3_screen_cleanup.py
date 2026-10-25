"""Final RC3 screen cleanup and Magnus safety contracts.

Only synthetic data. No user's Windows registry or evidence is changed.
"""
import inspect
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from production_app.gui import App
from production_app.facts import validate as validate_facts
from production_app.measurement_ranges import (
    magnus_dew_point,measurements_from_ranges,
    verify_midrange_provenance,validate_temperature_order,
)
from production_app.storage import Stop


def intervals():
    return {
        'air':{'min':'23','max':'27'},
        'surface':{'min':'20','max':'22'},
        'humidity':{'min':'50','max':'60'},
        'layer_intervals':[{'layer':1,'min':'90','max':'130'}],
    }


class WeatherMath(unittest.TestCase):
    def test_magnus_is_correct_at_saturation(self):
        self.assertEqual(magnus_dew_point('25','100'),Decimal('25.0'))

    def test_magnus_typical_positive_air_and_humidity(self):
        self.assertEqual(magnus_dew_point('25','55'),Decimal('15.3'))

    def test_magnus_zero_humidity_is_rejected(self):
        with self.assertRaisesRegex(Stop,'больше 0%'):
            magnus_dew_point('20','0')

    def test_approved_ranges_calculate_midpoint_and_dew_automatically(self):
        got=measurements_from_ranges(0,intervals())
        self.assertEqual(got['air'],'25')
        self.assertEqual(got['surface'],'21')
        self.assertEqual(got['humidity'],'55')
        self.assertEqual(got['dew'],'15.3')
        self.assertEqual(got['thickness'],'110')
        self.assertEqual(got['dew_method'],'MAGNUS_17.62_243.12')
        self.assertNotIn('dew',got['measurement_intervals'])
        self.assertTrue(verify_midrange_provenance(got))

    def test_tampered_dew_is_rejected_before_issuance(self):
        got=measurements_from_ranges(0,intervals())
        got['dew']='14.0'
        with self.assertRaisesRegex(Stop,'точка росы'):
            verify_midrange_provenance(got)

    def test_manual_dew_interval_is_rejected(self):
        got=measurements_from_ranges(0,intervals())
        got['measurement_intervals']['dew']={'min':'10','max':'12'}
        with self.assertRaisesRegex(Stop,'по Магнусу'):
            verify_midrange_provenance(got)

    def test_air_equal_or_colder_than_surface_is_rejected(self):
        data=intervals()
        data['surface']={'min':'24','max':'26'}
        with self.assertRaisesRegex(Stop,'воздуха должна быть'):
            measurements_from_ranges(0,data)
        with self.assertRaisesRegex(Stop,'строго выше'):
            validate_temperature_order({'air':'23.5',
                'surface':'23.5','dew':'12'})

    def test_insufficient_dew_point_margin_is_rejected(self):
        data=intervals()
        data['surface']={'min':'16','max':'18'}
        with self.assertRaisesRegex(Stop,'3 °C'):
            measurements_from_ranges(0,data)

    def test_wrong_ordered_humidity_range_is_rejected(self):
        data=intervals()
        data['humidity']={'min':'70','max':'40'}
        with self.assertRaisesRegex(Stop,'превышает конец'):
            measurements_from_ranges(0,data)


class GUIContracts(unittest.TestCase):
    def test_just_two_utility_buttons_on_main_screen(self):
        source=inspect.getsource(App.__init__)
        self.assertIn("('Фактические результаты',self.fact_dialog)",source)
        self.assertIn("('Открыть документы',self.open_documents)",source)
        self.assertNotIn("('Черновик справочника',self.catalog)",source)
        self.assertNotIn("('Восстановление',self.recovery)",source)
        self.assertIn("text='Сформировать пакет'",source)

    def test_recovery_is_in_settings_and_original_function_kept(self):
        source=inspect.getsource(App.settings)
        self.assertIn("Восстановление после сбоя",source)
        self.assertIn("self.recovery()",source)
        self.assertTrue(callable(App.recovery))

    def test_mark_table_only_four_columns_and_hidden_fields_retained(self):
        source=inspect.getsource(App.fact_dialog)
        self.assertIn(
            "cols=('mark','profile','akz','uzk_quantity')",source)
        self.assertIn('hidden_fields',source)
        self.assertIn('hidden_fields.get(mark,{})',source)
        self.assertIn('self.editable(grid,immutable={0,3}',source)
        for title in ("'Краска'", "'Результат УЗК'", "'Протокол ВИК'"):
            self.assertNotIn(title,source)

    def test_akz_has_no_separate_attach_button_but_global_evidence_exists(self):
        source=inspect.getsource(App.fact_dialog)
        self.assertNotIn("text='Приложить протокол АКЗ",source)
        self.assertNotIn("text='Протоколов АКЗ №",source)
        # The entire region below AKZ climate is removed from daily UI.
        self.assertNotIn("Приложить первичные протоколы",source)
        self.assertNotIn("Добавить фактическое соединение",source)
        self.assertNotIn("Дополнительные реквизиты утверждённых форм",source)
        self.assertIn("← Назад",source)
        self.assertIn("akz_proofs",source)

    def test_open_documents_uses_last_completed_folder_for_current_project(self):
        with tempfile.TemporaryDirectory() as directory:
            app=object.__new__(App)
            app.last_output_folder=directory
            app.last_output_project='Гараж Казахстан'
            app.plan=None
            app.project=SimpleNamespace(get=lambda:'Гараж Казахстан')
            with patch('production_app.gui.os.startfile',create=True) as opener:
                app.open_documents()
            opener.assert_called_once_with(str(Path(directory)))

    def test_open_documents_does_not_switch_to_previous_project(self):
        with tempfile.TemporaryDirectory() as directory:
            app=object.__new__(App)
            app.last_output_folder=directory
            app.last_output_project='Гараж Казахстан'
            app.plan=None
            app.project=SimpleNamespace(get=lambda:'ТЭМПО 1')
            with patch('production_app.gui.os.startfile',create=True) as opener,\
                 patch('production_app.gui.messagebox.showinfo') as info:
                app.open_documents()
            opener.assert_not_called()
            info.assert_called_once()

    def test_dew_is_read_only_display_not_manual_interval(self):
        source=inspect.getsource(App.fact_dialog)
        self.assertIn("'humidity':'Влажность, %'",source)
        self.assertNotIn("'dew':'Точка росы, °C'",source)
        self.assertIn("magnus_dew_point(",source)
        self.assertIn("for k in ('air','surface','humidity'):",source)
        self.assertNotIn("for key in ('air','surface','humidity','dew'):",source)


if __name__=='__main__':
    unittest.main(verbosity=2)
