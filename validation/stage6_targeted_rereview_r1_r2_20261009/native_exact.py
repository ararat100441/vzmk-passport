from review_common import *
import win32com.client
app=None;books=[]
def snapshot(b):
    s=b.Worksheets(1)
    return {'path':b.FullName,'name':b.Name,'read_only':bool(b.ReadOnly),'sheet_count':b.Worksheets.Count,'sheet':s.Name,'values':s.UsedRange.Value2,'formulas':s.UsedRange.Formula,'hidden_rows':[i for i in range(1,22) if s.Rows(i).Hidden],'pages':s.PageSetup.Pages.Count,'print_area':s.PageSetup.PrintArea,'header':s.PageSetup.CenterHeader,'footer':s.PageSetup.CenterFooter,'links':b.LinkSources(1),'connections':b.Connections.Count}
try:
    assert {n:sha(PAIR/n) for n in EXPECTED}==EXPECTED
    app=win32com.client.DispatchEx('Excel.Application')
    app.AutomationSecurity=3;app.EnableEvents=False;app.AskToUpdateLinks=False;app.DisplayAlerts=True;app.Visible=False
    for n in (ACTNAME,APPNAME):
        b=app.Workbooks.Open(Filename=str(PAIR/n),ReadOnly=True,CorruptLoad=0,UpdateLinks=0,AddToMru=False,Notify=False,IgnoreReadOnlyRecommended=True)
        books.append(b);assert b.ReadOnly and b.Worksheets.Count==1
    before=[snapshot(b) for b in books]
    assert all(x['pages']==1 for x in before)
    assert before[1]['hidden_rows']==list(range(7,21))
    save('native_exact_open.json',{'excel_version':app.Version,'CorruptLoad':0,'DisplayAlerts':True,'books':before})
    app.CalculateFullRebuild()
    save('native_exact_recalculated.json',[snapshot(b) for b in books])
    app.Visible=True;books[1].Activate()
    print('NATIVE_READY_APPLICATION HWND',app.Hwnd,flush=True)
    while True:
        command=input().strip()
        if command=='close':break
        if command=='act':books[0].Activate()
        if command=='application':books[1].Activate()
        print('ACTIVATED',command,'HWND',app.Hwnd,flush=True)
finally:
    for b in books:
        try:b.Close(SaveChanges=False)
        except Exception:pass
    if app is not None:app.Quit()
    after={n:sha(PAIR/n) for n in EXPECTED}
    assert after==EXPECTED
    save('native_exact_close.json',{'SaveChanges':False,'hashes_after':after,'audit':AUDIT})
