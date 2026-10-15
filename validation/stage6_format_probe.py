import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from akz_stage6.core import native_excel
with native_excel() as app:
    b=app.Workbooks.Add();s=b.Worksheets(1)
    s.Range('A1').Value2=3.989
    s.Range('A1').NumberFormatLocal='0'+str(app.DecimalSeparator)+'000'
    result={'separator':app.DecimalSeparator,'format_local':s.Range('A1').NumberFormatLocal,
            'format':s.Range('A1').NumberFormat,'text':s.Range('A1').Text,'value':s.Range('A1').Value2}
    assert s.Range('A1').Text in ('3,989','3.989')
    E=Path((ROOT/'validation/STAGE6_PART1_CURRENT.txt').read_text(encoding='utf-8'))
    (E/'native_format_probe.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(result)
    b.Close(SaveChanges=False);s=None;b=None
