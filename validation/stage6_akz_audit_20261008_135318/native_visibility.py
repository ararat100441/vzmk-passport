import sys
sys.dont_write_bytecode=True
import json, pythoncom, win32com.client
from pathlib import Path
from audit_readonly import OUT,allowed,sha,save
native=json.loads((OUT/'native_excel.json').read_text(encoding='utf-8'))
pythoncom.CoInitialize();excel=None;rows=[]
try:
    excel=win32com.client.DispatchEx('Excel.Application')
    excel.Visible=False;excel.DisplayAlerts=False;excel.EnableEvents=False;excel.AutomationSecurity=3;excel.AskToUpdateLinks=False
    for f in native['files']:
        p=Path(f['path'])
        if p.suffix.lower()!='.xls':continue
        allowed(p);before=sha(p);book=None
        try:
            book=excel.Workbooks.Open(Filename=str(p),UpdateLinks=0,ReadOnly=True,AddToMru=False,IgnoreReadOnlyRecommended=True)
            assert book.ReadOnly
            s=book.Worksheets(1); end=min(1000,s.UsedRange.Row+s.UsedRange.Rows.Count-1)
            visible=s.Range(f'B4:B{end}').SpecialCells(12)
            rr=[]
            for i in range(1,visible.Areas.Count+1):
                area=visible.Areas(i);rr.extend(range(area.Row,area.Row+area.Rows.Count))
            rec={'path':str(p),'sheet':s.Name,'visible_rows_B':rr,'filter_mode':bool(s.FilterMode),'autofilter_mode':bool(s.AutoFilterMode)}
            try:rec['autofilter_range']=s.AutoFilter.Range.Address
            except Exception:pass
            rows.append(rec);print('VISIBILITY',p.name,len(rr),flush=True)
        finally:
            if book is not None:book.Close(SaveChanges=False)
        assert before==sha(p)
    save('native_visibility.json',rows)
finally:
    if excel is not None:excel.Quit()
    pythoncom.CoUninitialize()
