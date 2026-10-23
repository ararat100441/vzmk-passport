"""No-production-data tests for the detailed historical-form contracts."""
from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest import TestCase, main
from unittest.mock import patch

from production_app.real_forms import (
    UZK_COLUMNS, UZK_HEADERS, DOC_FIELD_LABELS, UZK_CSV_HEADER,
    normalize_doc_fields, validate_uzk_details, load_uzk_csv,
    acceptance_acts, values_for_documents, require_form_inputs,
)
from production_app.storage import Stop


def shipment():
    return SimpleNamespace(rows=(
        ('M-1','Балка',Decimal(2),Decimal(1),Decimal(2),Decimal(1),Decimal(2)),
        ('M-2','Колонна',Decimal(1),Decimal(1),Decimal(1),Decimal(1),Decimal(1)),
    ))


def facts():
    return {
        'project':'P-TEST','number':'7/1/2026',
        'start_date':'2026-10-18','date':'2026-10-23',
        'vik_date':'2026-10-20','uzk_date':'2026-10-22','akz_date':'2026-10-21',
    }


def record(mark='M-1', date='2026-10-22'):
    return {
        'mark':mark, 'joint':'L=540x8 мм; L=800x12 мм',
        'welders':'Фиктивный исполнитель (только тест)',
        'date':date,'frequency_mhz':'5; 2,5','angle_deg':'70; 65',
        'defects':'TEST DEFECTS','conclusion':'ГОДЕН',
        'basis':'TEST протокол А',
    }


class RealForms(TestCase):
    def test_uzk_detailed_repeats_are_preserved(self):
        f=facts()
        f['uzk_details']=[record('M-1'),record('M-2'),record('M-1')]
        result=validate_uzk_details(f,shipment(),required=True)
        self.assertEqual([x['mark'] for x in result],['M-1','M-2','M-1'])
        self.assertEqual(result[1]['date'],'22.10.2026')
        self.assertEqual(len(UZK_HEADERS),len(UZK_COLUMNS))

    def test_missing_detailed_rows_stop_only_for_detailed_form(self):
        self.assertEqual(validate_uzk_details(facts(),shipment()),[])
        with self.assertRaisesRegex(Stop,'реальные строки'):
            validate_uzk_details(facts(),shipment(),required=True)

    def test_dates_and_wrong_marks_rejected(self):
        for changes in ({'date':'2026-10-30'},{'date':'2026-09-15'},
                        {'mark':'NOT_IN_VOM'}):
            with self.subTest(changes=changes):
                with self.assertRaises(Stop):
                    validate_uzk_details(facts()|{'uzk_details':[record()|changes]},
                                         shipment(),required=True)

    def test_frequency_angle_nonnumeric_and_mismatch_rejected(self):
        for changes in ({'frequency_mhz':'5;2,5','angle_deg':'70'},
                        {'frequency_mhz':'999','angle_deg':'70'},
                        {'frequency_mhz':'abc','angle_deg':'70'}):
            with self.subTest(changes=changes):
                with self.assertRaises(Stop):
                    validate_uzk_details(facts()|{'uzk_details':[record()|changes]},
                                         shipment(),required=True)

    def test_empty_or_injected_data_rejected(self):
        for changes in ({'basis':''},{'welders':'{{NUMBER}}'},
                        {'joint':''}):
            with self.subTest(changes=changes):
                with self.assertRaises(Stop):
                    validate_uzk_details(facts()|{'uzk_details':[record()|changes]},
                                         shipment(),required=True)

    def test_csv_import_preserves_repeated_marks(self):
        with TemporaryDirectory() as td:
            p=Path(td)/'uzk.csv'
            with p.open('w',encoding='utf-8-sig') as w:
                w.write(UZK_CSV_HEADER+'\n')
                for r in (record('M-1'),record('M-2'),record('M-1')):
                    w.write(';'.join(r[k] for k in UZK_COLUMNS)+'\n')
            # Raw semicolons inside fields must be CSV-quoted. Reject malformed
            # input immediately; a separate test covers well-formed repeats.
            with self.assertRaisesRegex(Stop,'Лишний разделитель'):
                load_uzk_csv(p)

    def test_csv_quoted_values_can_contain_semicolons(self):
        import csv
        with TemporaryDirectory() as td:
            p=Path(td)/'uzk.csv'
            with p.open('w',encoding='utf-8-sig',newline='') as out:
                writer=csv.DictWriter(out,fieldnames=UZK_COLUMNS,delimiter=';')
                writer.writeheader()
                writer.writerow(record('M-1'))
                writer.writerow(record('M-1'))
            loaded=load_uzk_csv(p)
            self.assertEqual([x['mark'] for x in loaded],['M-1','M-1'])
            self.assertEqual(loaded[0]['frequency_mhz'],'5; 2,5')
            self.assertEqual(len(validate_uzk_details(facts()|{'uzk_details':loaded},shipment())),2)

    def test_doc_fields_allow_only_named_confirmed_values(self):
        self.assertEqual(normalize_doc_fields({'VIK_DEFECTS':'  нет  '}),{'VIK_DEFECTS':'нет'})
        for obj in ({'UNLISTED':'x'}, {'VIK_DEFECTS':'{{NUMBER}}'},
                    {'VIK_DEFECTS':'a'*5001}):
            with self.subTest(obj=obj),self.assertRaises(Stop):
                normalize_doc_fields(obj)

    def test_acceptance_acts_are_current_not_stale(self):
        f=facts()
        text=acceptance_acts(f,[0,2])
        self.assertIn('7/1/2026',text)
        self.assertIn('20.10.2026',text)
        self.assertIn('22.10.2026',text)
        self.assertIn('21.10.2026',text)
        self.assertNotIn('1/1/2026 от 26.08.2026',text)
        self.assertIn('вариант 2',text)
        self.assertNotIn('вариант 1',text)
        self.assertEqual(values_for_documents(f,[0])['ATTACHMENTS'],
                         acceptance_acts(f,[0]))

    def test_preflight_requires_fields_of_approved_form(self):
        profile={'templates':{name:{'path':name+'.docx'} for name in ('vik','uzk','quality')}}
        data=facts()|{'doc_fields':{'QUALITY_STANDARDS':'PROJECT APPROVED'}}
        responses=[
            ({'NUMBER','CONTRACT','VIK_DEFECTS'},set()),
            ({'NUMBER','UZK_DETAILS_TABLE'},set()),
            ({'NUMBER','QUALITY_STANDARDS'},set()),
        ]
        with patch('production_app.word_controls.inspect_word_fields',side_effect=responses):
            with self.assertRaisesRegex(Stop,'ВИК — выявленные дефекты'):
                require_form_inputs(profile,data)
        responses=[
            ({'NUMBER','CONTRACT','VIK_DEFECTS'},set()),
            ({'NUMBER','UZK_DETAILS_TABLE'},set()),
            ({'NUMBER','QUALITY_STANDARDS'},set()),
        ]
        data['doc_fields']['VIK_DEFECTS']='описание из акта'
        with patch('production_app.word_controls.inspect_word_fields',side_effect=responses):
            self.assertTrue(require_form_inputs(profile,data))

    def test_preflight_unknown_fields_stop(self):
        profile={'templates':{name:{'path':name+'.docx'} for name in ('vik','uzk','quality')}}
        with patch('production_app.word_controls.inspect_word_fields',return_value=({'UNAUTHORIZED_FIELD'},set())):
            with self.assertRaisesRegex(Stop,'Неизвестные поля'):
                require_form_inputs(profile,facts())


if __name__ == '__main__':
    main(verbosity=2)
