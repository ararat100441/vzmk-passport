r"""Gate A: real Excel builtin date/time formatting on a TEMPORARY COPY.

Run on Windows with the exact historical three-coat VZMK source:
    py -3.12 -B tests/production/verify_ru_excel_com.py ^
        --format-source "C:\...\1_1. Акт АКЗ (3 cлоя) 7004.xlsx"

The historical file is NEVER opened for writing by Microsoft Excel. A
throwaway copy of it preserves date format ID 14, time format ID 20,
merged cells and formulas exactly as the production engine expects.
Only four explicit placeholders are inserted in the TEMPORARY COPY,
which is deleted after the test. No official document is produced.
"""
import argparse
import hashlib
import json
import math
import shutil
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


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def verify_builtin_format_ids(path):
    """Read-only OOXML check of the user's exact historic XLSX."""
    with zipfile.ZipFile(path) as package:
        styles=E.fromstring(package.read('xl/styles.xml'))
        sheet=E.fromstring(package.read('xl/worksheets/sheet1.xml'))
    xfs=list(styles.find(NS+'cellXfs'))
    expected={'H4':14,'E22':14,'E23':20,'E24':20}
    observed={}
    for addr,required in expected.items():
        element=sheet.find('.//'+NS+'c[@r="'+addr+'"]')
        if element is None:
            raise ValueError('Не найдена исходная ячейка '+addr)
        style=int(element.get('s','0'))
        observed[addr]=int(xfs[style].get('numFmtId','0'))
        if observed[addr]!=required:
            raise ValueError('Неподходящий Excel-формат в '+addr+
                             ': '+str(observed[addr])+', ожидался '+str(required))
    return observed


def verify_sheet(sheet):
    expected={
        'H4':'19.10.2026',
        'E22':'19.10.2026',
        'E23':'08:00',
        'E24':'10:30',
    }
    visible={}
    for addr,desired in expected.items():
        cell=sheet.Range(addr)
        actual=str(cell.Text).strip()
        numeric=cell.Value2
        if type(numeric) not in (int,float) or not math.isfinite(numeric):
            raise AssertionError((addr,'Не числовая дата/время',numeric))
        accepted=(_native_time_matches(actual,desired)
                  if addr in ('E23','E24') else actual==desired)
        if not accepted:
            raise AssertionError((addr,'Не тот месяц/время',actual,desired))
        visible[addr]=actual

    actual_date=(date(2026,10,19)-date(1899,12,30)).days
    if (int(sheet.Range('H4').Value2)!=actual_date or
            int(sheet.Range('E22').Value2)!=actual_date):
        raise AssertionError('Excel Serial даты не соответствует 19.10.2026')
    hours=float(sheet.Range('E24').Value2)-float(sheet.Range('E23').Value2)
    if not math.isclose(hours,2.5/24,rel_tol=0,abs_tol=1e-10):
        raise AssertionError('Excel Time не содержит двух числовых значений')
    if not sheet.Range('E30').HasFormula:
        raise AssertionError('Исходная дата E30 потеряла формулу')
    if not math.isclose(float(sheet.Range('E30').Value2),actual_date,
                        rel_tol=0,abs_tol=1e-8):
        raise AssertionError('Формула E30 не пересчитывает дату E22')
    return {
        'dates':visible,
        'date_serial':float(sheet.Range('H4').Value2),
        'time_start_serial':float(sheet.Range('E23').Value2),
        'time_end_serial':float(sheet.Range('E24').Value2),
        'dependent_formula':'E30',
    }


def main(source_path):
    source=Path(source_path).resolve(strict=True)
    if not source.is_file() or source.suffix.lower()!='.xlsx':
        raise ValueError('Требуется исходный трехслойный XLSX.')
    original_sha=sha(source)
    formats=verify_builtin_format_ids(source)

    with tempfile.TemporaryDirectory(prefix='vzmk-native-archive-copy-') as tmp:
        dest=Path(tmp)/'synthetic-date-check.xlsx'
        shutil.copyfile(source,dest)
        if sha(dest)!=original_sha:
            raise RuntimeError('SHA временной копии не совпал с оригиналом.')

        with excel_session() as app:
            workbook=app.Workbooks.Open(
                Filename=str(dest),ReadOnly=False,
                UpdateLinks=0,AddToMru=False,
                IgnoreReadOnlyRecommended=True,Notify=False,
            )
            try:
                if workbook.ReadOnly:
                    raise RuntimeError('Тестовый XLSX неожиданно открылся только для чтения.')
                sheet=workbook.Worksheets(1)
                for addr,token in {
                    'H4':'{{AKZ_DATE}}',
                    'E22':'{{AKZ_LAYER_1_APPLICATION_DATE}}',
                    'E23':'{{AKZ_LAYER_1_TIME_START}}',
                    'E24':'{{AKZ_LAYER_1_TIME_END}}',
                }.items():
                    sheet.Range(addr).Value2=token
                tokens_excel(sheet,{
                    'AKZ_DATE':'19.10.2026',
                    'AKZ_LAYER_1_APPLICATION_DATE':'2026-10-19',
                    'AKZ_LAYER_1_TIME_START':'08:00',
                    'AKZ_LAYER_1_TIME_END':'10:30',
                })
                app.CalculateFullRebuild()
                before=verify_sheet(sheet)
                workbook.Save()
            finally:
                workbook.Close(SaveChanges=False)

            reopened=app.Workbooks.Open(
                Filename=str(dest),ReadOnly=True,
                UpdateLinks=0,AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
            try:
                app.CalculateFullRebuild()
                after=verify_sheet(reopened.Worksheets(1))
            finally:
                reopened.Close(SaveChanges=False)

    if before!=after:
        raise RuntimeError('После Excel Save/ReadOnly Open изменились значения или формат.')
    if sha(source)!=original_sha:
        raise RuntimeError('Исходный акт АКЗ изменился: STOP.')
    return {
        'status':'PASS',
        'scope':'TEMPORARY_COPY_OF_ORIGINAL_NATIVE_STYLES',
        'native_format_ids':formats,
        'original_sha256_unchanged':True,
        'native_save_readback':True,
        'result':after,
    }


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--format-source',required=True,
                        help='Точный оригинальный акт АКЗ, только чтение')
    args=parser.parse_args()
    try:
        print(json.dumps(main(args.format_source),ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
