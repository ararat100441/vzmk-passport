"""Read-only native Excel inspection of exact artifacts; never saves books."""
from review import *
import win32com.client
def snapshot(book):
    sheets=[]
    for s in book.Worksheets:
        u=s.UsedRange
        sheets.append({'name':s.Name,'visible':s.Visible,'used':u.Address,'values':u.Value2,'formulas':u.Formula,
                       'print_area':s.PageSetup.PrintArea,'header':s.PageSetup.CenterHeader,
                       'footer':s.PageSetup.CenterFooter,'horizontal_breaks':s.HPageBreaks.Count,
                       'vertical_breaks':s.VPageBreaks.Count,
                       'hidden_rows':[r for r in range(1,min(u.Row+u.Rows.Count,1000)) if s.Rows(r).Hidden]})
    return {'path':book.FullName,'read_only':book.ReadOnly,'sheets':sheets}
m=load(I/'EVIDENCE_MANIFEST.json');pv=load(I/'preview.json')['payload'];app=None;books=[]
try:
    app=win32com.client.DispatchEx('Excel.Application')
    app.AutomationSecurity=3;app.EnableEvents=False;app.AskToUpdateLinks=False
    app.DisplayAlerts=True;app.Visible=False
    vom=app.Workbooks.Open(Filename=pv['vom']['path'],UpdateLinks=0,ReadOnly=True,AddToMru=False,CorruptLoad=0)
    assert vom.ReadOnly
    v=snapshot(vom);s=vom.Worksheets(pv['vom']['sheet'])
    v['filter']={'auto_filter_mode':s.AutoFilterMode,'filter_mode':s.FilterMode,
                 'address':s.AutoFilter.Range.Address if s.AutoFilterMode else None,
                 'filters':[{'field':i,'on':s.AutoFilter.Filters(i).On,
                             'criteria1':str(s.AutoFilter.Filters(i).Criteria1) if s.AutoFilter.Filters(i).On else None}
                            for i in range(1,s.AutoFilter.Filters.Count+1)]}
    save('native_vom_independent.json',v);vom.Close(SaveChanges=False);vom=None;s=None
    for path in sorted(Path(m['canonical_output']).glob('*.xlsx')):
        b=app.Workbooks.Open(Filename=str(path),UpdateLinks=0,ReadOnly=True,AddToMru=False,CorruptLoad=0)
        assert b.ReadOnly and not b.LinkSources(1) and b.Connections.Count==0
        books.append(b)
    save('native_outputs_on_open.json',{'excel_version':app.Version,'books':[snapshot(b) for b in books]})
    app.CalculateFullRebuild()
    save('native_outputs_recalculated.json',{'books':[snapshot(b) for b in books]})
    app.Visible=True
    books[0].Activate()
    print('NATIVE_READY_EXACT_READ_ONLY',flush=True)
    while True:
        command=input().strip()
        if command=='close':break
        if command=='act':books[0].Activate()
        if command=='application':books[1].Activate()
        print('NATIVE_COMMAND_DONE',command,flush=True)
finally:
    for b in books:
        try:b.Close(SaveChanges=False)
        except Exception:pass
    if app is not None:app.Quit()
    save('native_hash_after.json',{str(p):sha(p) for p in Path(m['canonical_output']).glob('*.xlsx')})
