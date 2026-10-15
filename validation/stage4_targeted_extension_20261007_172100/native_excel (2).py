"""Native Excel normal-open probe with alerts ON and independent dialog monitoring."""
from pathlib import Path
import hashlib
import json
import subprocess
import sys
import time
import traceback
import zipfile
from xml.etree import ElementTree as ET

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
sys.stdout.reconfigure(encoding='utf-8')
FILES=['Профили','Источники','Правила УЗК','Предпочтения','Конфликты']
MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def sha(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def save(path,value): Path(path).write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')

def worker(path,folder):
    import pythoncom
    import win32com.client
    import win32process
    pythoncom.CoInitialize()
    app=book=None
    result={'status':'FAIL','path':str(path),'corrupt_load':0,'read_only_requested':True}
    try:
        with zipfile.ZipFile(path) as package:
            expected_headers={}
            expected_ranges={}
            for name in package.namelist():
                if name.startswith('xl/tables/') and name.endswith('.xml'):
                    node=ET.fromstring(package.read(name))
                    expected_headers[node.get('name')]=[c.get('name') for c in node.find('{'+MAIN+'}tableColumns')]
                    expected_ranges[node.get('name')]=node.get('ref')
        app=win32com.client.DispatchEx('Excel.Application')
        app.Visible=False
        app.DisplayAlerts=True
        app.EnableEvents=False
        app.AutomationSecurity=3
        hwnd=int(app.Hwnd)
        pid=win32process.GetWindowThreadProcessId(hwnd)[1]
        save(folder/'session.json',{'pid':pid,'hwnd':hwnd,'display_alerts':bool(app.DisplayAlerts),'version':app.Version})
        assert bool(app.DisplayAlerts)
        book=app.Workbooks.Open(Filename=str(path),UpdateLinks=0,ReadOnly=True,
            IgnoreReadOnlyRecommended=True,AddToMru=False,CorruptLoad=0)
        result.update({'full_name':book.FullName,'read_only':bool(book.ReadOnly),
            'sheet_count':book.Worksheets.Count,'display_alerts_after_open':bool(app.DisplayAlerts),
            'connections':book.Connections.Count,
            'window_captions':[book.Windows(i).Caption for i in range(1,book.Windows.Count+1)],'tables':[]})
        try: result['repair_mode']=bool(book.RepairMode)
        except Exception: result['repair_mode']=None
        sheets=[book.Worksheets(i).Name for i in range(1,book.Worksheets.Count+1)]
        result['sheets']=sheets
        assert sheets==FILES and bool(book.ReadOnly) and bool(app.DisplayAlerts)
        assert book.Connections.Count==0
        assert str(Path(book.FullName).resolve()).casefold()==str(path).casefold()
        for index,name in enumerate(sheets,1):
            sheet=book.Worksheets(index)
            assert sheet.ListObjects.Count==1
            table=sheet.ListObjects(1)
            entry={'sheet':name,'listobjects':sheet.ListObjects.Count,'name':table.Name,
                'range':table.Range.Address,'data_rows':table.DataBodyRange.Rows.Count,
                'columns':table.ListColumns.Count,
                'column_names':[table.ListColumns(i).Name for i in range(1,table.ListColumns.Count+1)]}
            result['tables'].append(entry)
            assert entry['column_names']==expected_headers[entry['name']], 'Excel changed table headers during open'
            assert len({n.casefold() for n in entry['column_names']})==entry['columns']
            assert entry['range'].replace('$','')==expected_ranges[entry['name']]
        profiles=result['tables'][0]
        assert profiles['data_rows']==2851 and profiles['columns']==20
        try:
            formula_range=book.Worksheets(1).UsedRange.SpecialCells(-4123)
            result['profile_formulas']=formula_range.Count
        except Exception:
            result['profile_formulas']=0
        result['status']='OPENED'
    except Exception:
        result['error']=traceback.format_exc()
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()
        save(folder/'worker_result.json',result)

def run(path,label,expected_repair=False):
    import win32api
    import win32con
    import win32gui
    import win32process
    path=Path(path).resolve()
    assert path==ROOT/'DEV/Справочники/Профили.xlsx' or path.is_relative_to(OUT), 'DEV/isolated only'
    folder=OUT/'native'/label; folder.mkdir(parents=True)
    original=sha(path)
    temps=Path.home()/'AppData/Local/Temp'
    old_logs={str(p):sha(p) for p in temps.glob('error*.xml') if p.is_file()}
    existing_pids=set(win32process.EnumProcesses())
    process=subprocess.Popen([sys.executable,'-I','-B',str(Path(__file__).resolve()),'--worker',str(path),str(folder)],
                             stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    dialogs=[]; seen=set(); pid=None; timed_out=False; started=time.monotonic()
    while process.poll() is None:
        try:
            session=json.loads((folder/'session.json').read_text(encoding='utf-8'))
            pid=session['pid']
            assert pid not in existing_pids, 'Excel instance was not isolated'
            def visit(hwnd,_):
                if not win32gui.IsWindowVisible(hwnd) or win32process.GetWindowThreadProcessId(hwnd)[1]!=pid: return
                cls=win32gui.GetClassName(hwnd); title=win32gui.GetWindowText(hwnd)
                if cls=='#32770' or cls.startswith('bosa_sdm'):
                    children=[]
                    win32gui.EnumChildWindows(hwnd,lambda child,_:children.append(win32gui.GetWindowText(child)),None)
                    if hwnd not in seen:
                        seen.add(hwnd); dialogs.append({'hwnd':hwnd,'class':cls,'title':title,'texts':children,
                                                       'seconds_after_start':time.monotonic()-started})
                    # A detected dialog is always failure for fixed artifacts. Cancel
                    # it solely to clean up our isolated Excel; never accept recovery.
                    win32gui.PostMessage(hwnd,win32con.WM_CLOSE,0,0)
            win32gui.EnumWindows(visit,None)
        except (FileNotFoundError,json.JSONDecodeError): pass
        if time.monotonic()-started>75:
            timed_out=True
            process.terminate()
            if pid is not None and pid not in existing_pids:
                try:
                    handle=win32api.OpenProcess(win32con.PROCESS_TERMINATE,False,pid)
                    win32api.TerminateProcess(handle,1); win32api.CloseHandle(handle)
                except Exception: pass
            break
        time.sleep(0.05)
    stdout,stderr=process.communicate(timeout=10)
    result=json.loads((folder/'worker_result.json').read_text(encoding='utf-8')) if (folder/'worker_result.json').exists() else {'status':'FAIL'}
    new_logs=[]
    for log in temps.glob('error*.xml'):
        try:
            if str(log) not in old_logs or sha(log)!=old_logs[str(log)]:
                content=log.read_bytes()
                if any(token in content.decode('utf-8-sig',errors='replace') for token in ('table1.xml',path.name)):
                    copy=folder/log.name; copy.write_bytes(content)
                    new_logs.append({'source':str(log),'copy':str(copy),'sha256':sha(copy)})
        except OSError: pass
    recovered_title=any(any(t in caption.casefold() for t in ('восстановлен','recovered','repaired'))
                        for caption in result.get('window_captions',[]))
    repair=bool(dialogs or new_logs or result.get('repair_mode') or recovered_title)
    result.update({'dialog_monitor':True,'dialogs':dialogs,'new_repair_logs':new_logs,
        'repair_detected':repair,'recovered_title':recovered_title,'timed_out':timed_out,
        'sha256_before':original,'sha256_after':sha(path),'source_unchanged':sha(path)==original,
        'exit_code':process.returncode,'stderr':stderr.decode('utf-8',errors='replace')})
    success=(result['status']=='OPENED' and not repair and not timed_out and result['source_unchanged'])
    result['validation']='EXPECTED_DEFECT_REPRODUCED' if expected_repair and (repair or result['status']!='OPENED') else 'PASS' if success else 'FAIL'
    save(folder/'result.json',result)
    if expected_repair: assert result['validation']=='EXPECTED_DEFECT_REPRODUCED',result
    else: assert success,result
    return result

if __name__=='__main__':
    if sys.argv[1]=='--worker': worker(Path(sys.argv[2]),Path(sys.argv[3]))
    else:
        result=run(Path(sys.argv[2]),sys.argv[1],expected_repair='--expected-repair' in sys.argv[3:])
        print(json.dumps(result,ensure_ascii=False))
