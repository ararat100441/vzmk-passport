"""Explicit read-only native Excel inspection. No save, print, PDF, or recovery mode."""
import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from akz_stage6.core import native_excel,native_open,sha256
path=Path(sys.argv[1]).resolve()
before=sha256(path)
with native_excel() as app:
    app.DisplayAlerts=True
    b=native_open(app,path,read_only=True)
    try:
        app.Visible=True
        s=b.Worksheets(1)
        s.Activate()
        app.ActiveWindow.Zoom=70
        s.Range('A1').Select()
        report={'path':str(path),'input_sha256':before,'read_only':bool(b.ReadOnly),'normal_corrupt_load':0,
                'excel_version':app.Version,'sheets':[]}
        for sheet in b.Worksheets:
            sheet.Activate()
            try:
                pages=int(sheet.PageSetup.Pages.Count)
            except Exception as exc:
                pages={'status':'NOT_VERIFIED','reason':str(exc)}
            report['sheets'].append({'name':sheet.Name,'visible':sheet.Visible,
                'pages':pages,
                'print_area':sheet.PageSetup.PrintArea,'print_title_rows':sheet.PageSetup.PrintTitleRows,
                'header':sheet.PageSetup.CenterHeader,'footer':sheet.PageSetup.CenterFooter,
                'values':sheet.UsedRange.Value2})
        print(json.dumps({k:v for k,v in report.items() if k!='sheets'},ensure_ascii=False,default=str),flush=True)
        print(json.dumps([{k:v for k,v in s.items() if k!='values'} for s in report['sheets']],ensure_ascii=False),flush=True)
        if len(sys.argv)>2:
            Path(sys.argv[2]).write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
        print('READY_FOR_UI_INSPECTION; send newline to close without save',flush=True)
        sys.stdin.readline()
    finally:
        b.Close(SaveChanges=False)
        assert sha256(path)==before,'OUTPUT_CHANGED_DURING_NATIVE_REVIEW'
        if len(sys.argv)>2:
            after=Path(sys.argv[2]).with_suffix('.after.json')
            after.write_text(json.dumps({'path':str(path),'before':before,'after':sha256(path),'unchanged':True},ensure_ascii=False,indent=2),encoding='utf-8')
