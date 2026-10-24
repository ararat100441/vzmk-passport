"""Targeted native Windows Excel probe for Russian date/time formats.

Run BEFORE the full E2E. Creates only a temporary workbook under the OS
temporary directory. No existing Excel instance, registry or business file
is accessed. Checks local rendering, typed numeric values, formula references,
and Save -> Close -> ReadOnly Open -> verification.

Usage: py -3.12 -B tests/production/verify_ru_excel_com.py
"""
import json
import math
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from passport_app.excel import excel_session
from production_app.office import tokens_excel


def verify(sheet):
    expected={'A1':'19.10.2026','A2':'19.10.2026',
              'A3':'08:00','A4':'10:30'}
    for addr,shown in expected.items():
        cell=sheet.Range(addr)
        value=cell.Value2
        actual=str(cell.Text).strip()
        assert type(value) in (int,float) and math.isfinite(value), (
            addr, 'Not numeric', value,
        )
        assert actual==shown, (addr,'Bad native ru-RU Excel display',actual,shown)
    assert math.isclose(float(sheet.Range('A4').Value2)-
                        float(sheet.Range('A3').Value2),
                        2.5/24, rel_tol=0, abs_tol=1e-10)
    assert sheet.Range('B1').HasFormula
    assert math.isclose(float(sheet.Range('B1').Value2),
                        float(sheet.Range('A1').Value2)+1,
                        rel_tol=0, abs_tol=1e-8)
    assert sheet.Range('B2').HasFormula
    assert math.isclose(float(sheet.Range('B2').Value2),
                        float(sheet.Range('A4').Value2)-
                        float(sheet.Range('A3').Value2),
                        rel_tol=0, abs_tol=1e-10)
    return {'display':{addr:str(sheet.Range(addr).Text) for addr in expected},
            'date_serial':float(sheet.Range('A1').Value2),
            'time_start_serial':float(sheet.Range('A3').Value2),
            'time_end_serial':float(sheet.Range('A4').Value2)}


def main():
    with tempfile.TemporaryDirectory(prefix='vzmk-ru-excel-probe-') as dirname:
        out=Path(dirname)/'probe.xlsx'
        with excel_session() as app:
            b=app.Workbooks.Add(-4167)
            try:
                s=b.Worksheets(1)
                assignments={
                    'A1':'{{AKZ_DATE}}',
                    'A2':'{{AKZ_LAYER_1_APPLICATION_DATE}}',
                    'A3':'{{AKZ_LAYER_1_TIME_START}}',
                    'A4':'{{AKZ_LAYER_1_TIME_END}}',
                }
                for addr,text in assignments.items():
                    s.Range(addr).Value2=text
                s.Range('B1').Formula='=A1+1'
                s.Range('B2').Formula='=A4-A3'
                tokens_excel(s,{
                    'AKZ_DATE':'19.10.2026',
                    'AKZ_LAYER_1_APPLICATION_DATE':'2026-10-19',
                    'AKZ_LAYER_1_TIME_START':'08:00',
                    'AKZ_LAYER_1_TIME_END':'10:30',
                })
                app.CalculateFullRebuild()
                before=verify(s)
                b.SaveAs(str(out),FileFormat=51)
            finally:
                b.Close(SaveChanges=False)
            reopened=app.Workbooks.Open(
                Filename=str(out), UpdateLinks=0,ReadOnly=True,
                AddToMru=False,IgnoreReadOnlyRecommended=True,
            )
            try:
                after=verify(reopened.Worksheets(1))
            finally:
                reopened.Close(SaveChanges=False)
        assert before==after,(before,after)
        return {
            'status':'PASS',
            'scope':'SYNTHETIC_RUSSIAN_EXCEL_DATE_TIME',
            'native_save_and_readback':True,
            'original_business_files_touched':False,
            'values':after,
        }


if __name__=='__main__':
    try:
        print(json.dumps(main(),ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL: '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
