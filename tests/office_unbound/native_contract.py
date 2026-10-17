"""Exercise the unchanged supervisor with real owned unbound Excel processes.

Only worker fixture routing and a negative helper proof fault are injected.
The positive case uses the stock cleanup entry and all identity/exit checks.
"""
import argparse,gc,json,os,subprocess,sys,time
from pathlib import Path
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app import supervisor,ownership
from stage9_app.workspace import save
from openpyxl import Workbook
import win32con,win32event,win32gui,win32process,pythoncom

def ready(pid):
    found=[]
    def top(hwnd,_):
        if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid:return
        def child(hwnd,_):
            if win32gui.GetClassName(hwnd)=='EXCEL7' and win32process.GetWindowThreadProcessId(hwnd)[1]==pid:found.append(hwnd)
        win32gui.EnumChildWindows(hwnd,child,None)
    win32gui.EnumWindows(top,None)
    return bool(found)

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();out=args.output.resolve();out.mkdir(parents=True,exist_ok=False)
    control_probe=out/'PREEXISTING_CONTROL.xlsx';b=Workbook();b.active['A1']='PREEXISTING TEST CONTROL - DO NOT TOUCH';b.save(control_probe);b.close()
    control,record=ownership.launch_child('Excel.Application',control_probe,os.getpid())
    save(out/'CONTROL_CREATED.json',record)
    deadline=time.monotonic()+30
    while not ready(control.pid) and time.monotonic()<deadline:time.sleep(.1)
    assert ready(control.pid),'CONTROL_NOT_READY'
    original_popen=subprocess.Popen
    results=[]
    try:
        for negative in (False,True):
            folder=out/('negative' if negative else 'positive');folder.mkdir()
            envdir=folder/'appdata'
            def routed(command,*a,**kw):
                cmd=list(command)
                if '--office-worker' in cmd:
                    cmd[cmd.index(str(ROOT/'stage9.py'))]=str(Path(__file__).with_name('fixture_worker.py'))
                elif negative and '--office-cleanup' in cmd:
                    cmd[cmd.index(str(ROOT/'stage9.py'))]=str(Path(__file__).with_name('fixture_helper_failure.py'))
                return original_popen(cmd,*a,**kw)
            began=time.monotonic()
            with patch.dict(os.environ,{'LOCALAPPDATA':str(envdir)}),patch('subprocess.Popen',side_effect=routed):
                code=supervisor.supervise(['--task','office-probe','--result',str(folder/'RESULT.json')])
            duration=time.monotonic()-began
            result=json.loads((folder/'RESULT.json').read_text(encoding='utf8'))
            summary=json.loads(Path(result['office_supervisor']).read_text(encoding='utf8'))
            assert len(summary['instances'])==1 and summary['worker_returncode']==0 and summary['normal_result']
            instance=summary['instances'][0]
            assert not instance.get('com_binding_proven') and not instance.get('quit_returned')
            assert instance['pid']!=control.pid
            assert instance['crash_cleanup']['exit_code']==0 and instance['crash_cleanup']['exit_observed_while_sta_alive'] is True
            assert instance['exit_code']==0 and not instance['still_running'] and not instance.get('forced_termination_attempted')
            assert summary['owned_orphans']==0 and not summary['ownership_errors']
            assert win32event.WaitForSingleObject(control._handle,0)!=win32con.WAIT_OBJECT_0
            assert win32process.GetProcessTimes(control._handle)['CreationTime'].isoformat()==record['creation_time']
            if negative:
                assert code==1 and result['status']=='FAIL' and result['error']=='OFFICE_OWNED_UNBOUND_CLEANUP_FAILED',result
            else:
                assert code==0 and result['status']=='PASS' and summary['owned_unbound_gracefully_recovered']==1,result
                assert summary['wait_seconds_after_worker']<60,summary
            row={'case':'negative' if negative else 'positive','test_status':'PASS','operation_status':result['status'],
                 'duration_seconds':duration,'office_wait_seconds':summary['wait_seconds_after_worker'],
                 'recovered':summary['owned_unbound_gracefully_recovered'],'forced_terminations':0,
                 'control_untouched':True,'supervisor':result['office_supervisor']}
            results.append(row);print(json.dumps(row),flush=True)
    finally:
        pythoncom.CoInitialize()
        try:
            gc.collect();closed=ownership.quit_owned_window(record,control._handle)
            save(out/'CONTROL_CLOSED.json',closed)
            assert closed['status']=='PASS' and closed['exit_code']==0,closed
            control.wait(timeout=10)
        finally:pythoncom.CoUninitialize()
    save(out/'SUMMARY.json',{'status':'PASS','cases':results,'control_shutdown':'GRACEFUL_EXIT_0'})

if __name__=='__main__':main()
