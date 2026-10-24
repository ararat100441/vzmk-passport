"""Native Excel Gate A: preserve ORIGINAL builtin XLSX date/time styles.

Usage on Windows:
    py -3.12 -B tests/production/verify_ru_excel_com.py ^
      --format-source "C:\...\1_1. Акт АКЗ (3 cлоя) 7004.xlsx"

Reads the exact historical VZMK three-coat workbook in READ-ONLY mode as
a *format donor*. Copies ONLY native formatting from H4 (built-in date 14)
and E23 (built-in time 20) into an isolated synthetic workbook. Original
business data, values and formulas are never copied or saved.

This is not permission to reuse earlier inspection results in a new passport.
"""
import argparse
import hashlib
import json
import math
import sys
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as E

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from passport_app.excel import excel_session
from production_app.office import tokens_excel,_native_time_matches

NS='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XL_PASTE_FORMATS=-4122


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def ensure_historical_builtin_styles(path):
    """Read-only OOXML check; prevents using a similarly named wrong file."""
    with zipfile.ZipFile(path) as package:
        styles=E.fromstring(package.read('xl/styles.xml'))
        sheet=E.fromstring(package.read('xl/worksheets/sheet1.xml'))
    xfs=list(styles.find(NS+'cellXfs'))
    actual={}
    for address in ('H4','E23'):
        element=sheet.find('.//'+NS+'c[@r="'+address+'"]')
        if element is None:
            raise ValueError('Исходный акт не содержит ячейку '+address)
        index=int(element.get('s','0'))
        actual[address]=int(xfs[index].get('numFmtId','0'))
    if actual != {'H4':14,'E23':20}:
        raise ValueError('Неверный источник форматов: ожидались H4=14 (дата) и '
                         'E23=20 (время), получено '+repr(actual))
    return actual


def validate_sheet(sheet):
    formats={
        'A1':'19.10.2026','A2':'19.10.2026',
        'A3':'08:00','A4':'10:30',
    }
    display={}
    for address,expected in formats.items():
        cell=sheet.Range(address)
        value=cell.Value2
        shown=str(cell.Text).strip()
        if type(value) not in (int,float) or not math.isfinite(value):
            raise AssertionError((address,'not numeric',value))
        correct=(_native_time_matches(shown,expected)
                 if address in ('A3','A4') else shown==expected)
        if not correct:
            raise AssertionError((address,'invalid date/time display',shown,expected))
        display[address]=shown
    assert math.isclose(float(sheet.Range('A4').Value2)-
                        float(sheet.Range('A3').Value2),
                        2.5/24,rel_tol=0,abs_tol=1e-10)
    assert sheet.Range('B1').HasFormula
    assert math.isclose(float(sheet.Range('B1').Value2),
                        float(sheet.Range('A1').Value2)+1,
                        rel_tol=0,abs_tol=1e-8)
    assert sheet.Range('B2').HasFormula
    assert math.isclose(float(sheet.Range('B2').Value2),
                        float(sheet.Range('A4').Value2)-
                        float(sheet.Range('A3').Value2),
                        rel_tol=0,abs_tol=1e-10)
    return {
        'display':display,
        'date_serial':float(sheet.Range('A1').Value2),
        'time_start_serial':float(sheet.Range('A3').Value2),
        'time_end_serial':float(sheet.Range('A4').Value2),
    }


def copy_only_format(original, target, app):
    original.Copy()
    target.PasteSpecial(Paste=XL_PASTE_FORMATS)
    app.CutCopyMode=False


def main(source_path):
    source=Path(source_path).resolve(strict=True)
    if not source.is_file() or source.suffix.lower()!='.xlsx':
        raise ValueError('Нужен точный исходный трехслойный XLSX.')
    before_hash=sha(source)
    style_ids=ensure_historical_builtin_styles(source)
    with tempfile.TemporaryDirectory(prefix='vzmk-native-format-probe-') as dirname:
        output=Path(dirname)/'probe.xlsx'
        with excel_session() as app:
            donor=app.Workbooks.Open(
                Filename=str(source),ReadOnly=True,UpdateLinks=0,
                AddToMru=False,IgnoreReadOnlyRecommended=True,
            )
            workbook=None
            try:
                if not donor.ReadOnly:
                    raise RuntimeError('Файл-образец неожиданно открыт на запись.')
                donor_sheet=donor.Worksheets(1)
                workbook=app.Workbooks.Add(-4167)
                sheet=workbook.Worksheets(1)
                sheet.Columns('A:B').ColumnWidth=20

                for address in ('A1','A2'):
                    copy_only_format(donor_sheet.Range('H4'),
                                     sheet.Range(address),app)
                for address in ('A3','A4'):
                    copy_only_format(donor_sheet.Range('E23'),
                                     sheet.Range(address),app)
            finally:
                donor.Close(SaveChanges=False)
            if workbook is None:
                raise RuntimeError('Не создан синтетический файл для проверки.')
            try:
                for address,token in {
                    'A1':'{{AKZ_DATE}}',
                    'A2':'{{AKZ_LAYER_1_APPLICATION_DATE}}',
                    'A3':'{{AKZ_LAYER_1_TIME_START}}',
                    'A4':'{{AKZ_LAYER_1_TIME_END}}',
                }.items():
                    sheet.Range(address).Value2=token
                sheet.Range('B1').Formula='=A1+1'
                sheet.Range('B2').Formula='=A4-A3'
                tokens_excel(sheet,{
                    'AKZ_DATE':'19.10.2026',
                    'AKZ_LAYER_1_APPLICATION_DATE':'2026-10-19',
                    'AKZ_LAYER_1_TIME_START':'08:00',
                    'AKZ_LAYER_1_TIME_END':'10:30',
                })
                app.CalculateFullRebuild()
                pre_save=validate_sheet(sheet)
                workbook.SaveAs(str(output),FileFormat=51)
            finally:
                workbook.Close(SaveChanges=False)
            check=app.Workbooks.Open(
                Filename=str(output),ReadOnly=True,UpdateLinks=0,
                AddToMru=False,IgnoreReadOnlyRecommended=True,
            )
            try:
                app.CalculateFullRebuild()
                after_save=validate_sheet(check.Worksheets(1))
            finally:
                check.Close(SaveChanges=False)
        if pre_save!=after_save:
            raise RuntimeError('После Save/ReadOnly Open изменилось содержимое.')
    if sha(source)!=before_hash:
        raise RuntimeError('Оригинальный трёхслойный акт изменился — STOP.')
    return {
        'status':'PASS',
        'scope':'SYNTHETIC_WITH_ORIGINAL_BUILTIN_FORMAT_ONLY',
        'donor_style_ids':style_ids,
        'read_only_source_sha256_preserved':True,
        'native_save_readback':True,
        'result':after_save,
    }


if __name__=='__main__':
    args=argparse.ArgumentParser()
    args.add_argument('--format-source',required=True,
                      help='Путь к исходному акту АКЗ на 3 слоя, только чтение')
    parsed=args.parse_args()
    try:
        print(json.dumps(main(parsed.format_source),ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
