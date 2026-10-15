"""Normal Excel Open in a fresh COM instance, with alerts and dialog monitoring."""
from pathlib import Path
import hashlib
import json
import subprocess
import sys
import time
import traceback

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
SHEETS=['Профили','Источники','Правила УЗК','Предпочтения','Конфликты']
sys.stdout.reconfigure(encoding='utf-8')
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,x):Path(p).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf-8')
def worker(path,folder):
    import pythoncom
    import win32com.client
    import win32process
    pythoncom.CoInitialize()
    app=book=None
    r={'opened_normally':False,'CorruptLoad':0,'ReadOnly_requested':True,'SaveChanges':False}
    try:
        app=win32com.client.DispatchEx('Excel.Application')
        app.Visible=False
        app.DisplayAlerts=True
        app.EnableEvents=False
        app.AutomationSecurity=3
        pid=win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]
        save(folder/'session.json',{'pid':pid,'hwnd':int(app.Hwnd),'Excel_version':str(app.Version),'alerts_enabled':bool(app.DisplayAlerts)})
        assert app.DisplayAlerts
        book=app.Workbooks.Open(Filename=str(path),ReadOnly=True,UpdateLinks=0,AddToMru=False,
                               IgnoreReadOnlyRecommended=True,CorruptLoad=0)
        r.update(full_name=str(book.FullName),read_only=bool(book.ReadOnly),alerts_enabled=bool(app.DisplayAlerts),
                 sheets=[book.Worksheets(i).Name for i in range(1,book.Worksheets.Count+1)],
                 captions=[str(book.Windows(i).Caption) for i in range(1,book.Windows.Count+1)],tables=[])
        assert Path(book.FullName).resolve()==path and book.ReadOnly and app.DisplayAlerts
        assert r['sheets']==SHEETS
        for i in range(1,6):
            sheet=book.Worksheets(i)
            assert sheet.ListObjects.Count==1
            t=sheet.ListObjects(1)
            r['tables'].append({'sheet':sheet.Name,'name':t.Name,'ref':t.Range.Address,
               'columns':t.ListColumns.Count,'rows':t.DataBodyRange.Rows.Count,
               'headers':[t.ListColumns(j).Name for j in range(1,t.ListColumns.Count+1)],'show_filter':bool(t.ShowAutoFilter)})
        p=r['tables'][0]
        assert p['ref']=='$A$1:$T$2850' and p['rows']==2849 and p['columns']==20
        assert p['headers'][10]=='r' and p['headers'][15]=='Радиус R'
        assert all(len(set(h.casefold() for h in t['headers']))==t['columns'] for t in r['tables'])
        assert all(t['show_filter'] for t in r['tables'])
        r['opened_normally']=True
    except BaseException:r['error']=traceback.format_exc()
    finally:
        try:
            if book is not None:book.Close(SaveChanges=False)
            if app is not None:app.Quit()
        finally:
            pythoncom.CoUninitialize()
            save(folder/'worker.json',r)

def run(path,label,expect_defect=False):
    import win32gui,win32process,win32con,win32api
    path=Path(path).resolve()
    assert path.is_relative_to(OUT) or path==ROOT/'DEV/Справочники/Профили.xlsx'
    folder=OUT/'native'/label;folder.mkdir(parents=True)
    old=sha(path)
    temp=Path.home()/'AppData/Local/Temp'
    old_logs={str(p):sha(p) for p in temp.glob('error*.xml') if p.is_file()}
    existing=set(win32process.EnumProcesses())
    p=subprocess.Popen([sys.executable,'-I','-B',str(Path(__file__).resolve()),'worker',str(path),str(folder)],stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    dialogs=[];seen=set();pid=None;timeout=False;start=time.monotonic()
    while p.poll() is None:
        try:
            session=json.loads((folder/'session.json').read_text(encoding='utf-8'))
            pid=session['pid'];assert pid not in existing,'Fresh COM instance required'
            def window(hwnd,_):
                if not win32gui.IsWindowVisible(hwnd) or win32process.GetWindowThreadProcessId(hwnd)[1]!=pid:return
                cls=win32gui.GetClassName(hwnd)
                if cls=='#32770' or cls.startswith('bosa_sdm'):
                    words=[]
                    win32gui.EnumChildWindows(hwnd,lambda child,_:words.append(win32gui.GetWindowText(child)),None)
                    if hwnd not in seen:
                        seen.add(hwnd);dialogs.append({'title':win32gui.GetWindowText(hwnd),'class':cls,'texts':words})
                    # Any dialog fails validation. Cancel it; never accept recovery.
                    win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
            win32gui.EnumWindows(window,None)
        except (FileNotFoundError,json.JSONDecodeError):pass
        if time.monotonic()-start>45:
            timeout=True;p.terminate()
            if pid is not None and pid not in existing:
                h=win32api.OpenProcess(win32con.PROCESS_TERMINATE,False,pid)
                win32api.TerminateProcess(h,1);win32api.CloseHandle(h)
            break
        time.sleep(.05)
    stdout,stderr=p.communicate(timeout=10)
    result=json.loads((folder/'worker.json').read_text(encoding='utf-8')) if (folder/'worker.json').exists() else {'opened_normally':False}
    logs=[]
    for log in temp.glob('error*.xml'):
        if str(log) not in old_logs or sha(log)!=old_logs[str(log)]:
            text=log.read_bytes().decode('utf-8-sig',errors='replace')
            if path.name in text or 'table1.xml' in text:
                dest=folder/log.name;dest.write_bytes(log.read_bytes());logs.append({'file':log.name,'sha256':sha(dest)})
    recovered=any(any(s in title.casefold() for s in ['восстановлен','recovered','repaired']) for title in result.get('captions',[]))
    result.update(dialogs=dialogs,repair_prompt=bool(dialogs),recovered_records=bool(logs),new_recovery_logs=logs,
                  recovered_title=recovered,timed_out=timeout,sha256_before=old,sha256_after=sha(path),
                  source_unchanged=old==sha(path),exit_code=p.returncode,stderr=stderr.decode('utf-8',errors='replace'),
                  expected_negative_control=expect_defect)
    good=result['opened_normally'] and not dialogs and not logs and not recovered and not timeout and old==sha(path) and p.returncode==0
    result['PASS']=((not result['opened_normally'] or bool(dialogs) or bool(logs) or recovered) and old==sha(path) and not timeout) if expect_defect else good
    save(folder/'result.json',result)
    assert result['PASS'],result
    return result
if __name__=='__main__':
    if sys.argv[1]=='worker':worker(Path(sys.argv[2]),Path(sys.argv[3]))
    else:print(json.dumps(run(Path(sys.argv[2]),sys.argv[1],'--negative' in sys.argv[3:]),ensure_ascii=False))
