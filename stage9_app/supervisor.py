"""Keep Office handles outside the worker, including after hard worker death."""
from pathlib import Path
import json
import os
import subprocess
import sys
import time
from uuid import uuid4
from .workspace import save
from .ownership import process_image

def cleanup_entry(path):
    import win32api,win32con
    from stage8_closure import storage as S
    from .ownership import quit_owned_window
    plan=S.read(path);record=plan['record'];payload={k:v for k,v in plan.items() if k!='seal'}
    if plan['seal']!=S.digest(payload) or plan['parent_pid']!=os.getppid():raise RuntimeError('CLEANUP_PLAN_OR_PARENT_CHANGED')
    import win32process
    parent=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,os.getppid())
    try:
        if win32process.GetProcessTimes(parent)['CreationTime'].isoformat()!=plan['parent_creation_time'] or process_image(parent)!=process_image(win32api.GetCurrentProcess()):raise RuntimeError('CLEANUP_PARENT_NOT_PROVEN')
        if __import__('datetime').datetime.fromisoformat(record['creation_time']).timestamp()<win32process.GetProcessTimes(parent)['CreationTime'].timestamp():raise RuntimeError('PREEXISTING_OFFICE_NO_QUIT')
    finally:parent.Close()
    handle=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.SYNCHRONIZE,False,record['pid'])
    try:
        def trace(phase,details):save(Path(path).with_name(Path(path).stem+'_'+phase+'_'+uuid4().hex[:8]+'.json'),{'phase':phase,'details':details,'monotonic':time.monotonic()})
        try:result=quit_owned_window(record,handle,trace)
        except BaseException as exc:
            import traceback
            result={'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc()}
        save(Path(path).with_suffix('.result.json'),result)
    finally:handle.Close()
    return 0 if result['status']=='PASS' else 1

def supervise(arguments):
    import win32api,win32process,win32con,win32event,win32gui
    base=Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'VZMKPassport/diagnostics'/('job_'+uuid4().hex)
    base.mkdir(parents=True);args=list(arguments);target=None
    save(base/'JOB_REQUEST.json',{'schema':'STAGE9_JOB_REQUEST_V1','arguments':args,'supervisor_pid':os.getpid(),
                                 'supervisor_image':process_image(win32api.GetCurrentProcess()),'created_at':time.time()})
    if '--result' in args:
        n=args.index('--result');target=Path(args[n+1]);del args[n:n+2]
        if target.exists():raise RuntimeError('RESULT_ALREADY_EXISTS')
    result=base/'worker_result.json';diagnostic=base/'office.json'
    prefix=[sys.executable] if getattr(sys,'frozen',False) else [sys.executable,'-B',str(Path(__file__).resolve().parents[1]/'stage9.py')]
    before=set(win32process.EnumProcesses());handles={};records=[];seen=set();launched=set();office_children=[];ownership_errors=[]
    child=subprocess.Popen(prefix+args+['--result',str(result),'--office-worker',str(diagnostic)],
        creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0),stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    child_birth=win32process.GetProcessTimes(child._handle)['CreationTime'].timestamp()
    def collect():
        for path in base.glob('launch_*.json'):
            if path.name.endswith('.response.json') or path in launched:continue
            try:request=json.loads(path.read_text(encoding='utf8'))
            except (OSError,ValueError):continue
            launched.add(path)
            try:
                from stage8_closure import storage as S
                from .ownership import launch_child
                probe=S.no_links(request['probe'])
                if request['worker_pid']!=child.pid or request['progid'] not in ('Excel.Application','Word.Application') or probe.parent!=base.resolve() or S.sha(probe)!=request['probe_sha256']:raise RuntimeError('OFFICE_LAUNCH_REQUEST_NOT_OWNED')
                office,r=launch_child(request['progid'],request['probe'],child.pid)
                office_children.append(office)
                if r['pid'] in before:raise RuntimeError('OFFICE_CHILD_PID_PREEXISTED')
                handles[r['pid']]=office._handle;records.append(r)
                save(base/('created_'+str(r['pid'])+'.json'),r)
                response={'status':'PASS','record':r}
            except BaseException as exc:
                ownership_errors.append({'path':str(path),'error':repr(exc),'unknown_process_not_terminated':True})
                response={'status':'FAIL','error':repr(exc)}
            temporary=path.with_suffix('.pending');save(temporary,response);os.rename(temporary,path.with_suffix('.response.json'))
        for path in base.glob('office_*_START.json'):
            if path in seen:continue
            try:r=json.loads(path.read_text(encoding='utf8'))
            except (OSError,ValueError):continue
            seen.add(path)
            try:
                pid=r['pid']
                if r.get('worker_pid')!=child.pid or r.get('launcher_pid')!=os.getpid() or r.get('actual_child') is not True or r.get('owned') is not True or r.get('absent_before_dispatch') is not True or pid in before or pid not in handles:raise RuntimeError('OFFICE_OWNERSHIP_NOT_PROVEN')
                handle=handles[pid]
                born=win32process.GetProcessTimes(handle)['CreationTime']
                if born.isoformat()!=r['creation_time'] or born.timestamp()<child_birth or process_image(handle)!=r['image']:raise RuntimeError('OFFICE_PROCESS_IDENTITY_MISMATCH')
                if Path(r['image']).name.upper() not in ('EXCEL.EXE','WINWORD.EXE'):raise RuntimeError('OFFICE_IMAGE_MISMATCH')
                for current in records:
                    if current['pid']==pid:
                        current['hwnd']=r['hwnd'];current['com_binding_proven']=True
                        for key in ('user_control_before_binding','user_control_after_binding'):
                            if key in r:current[key]=r[key]
            except BaseException as exc:
                ownership_errors.append({'path':str(path),'error':repr(exc),'unknown_process_not_terminated':True})
    try:
        while child.poll() is None:
            collect()
            if ownership_errors:child.terminate();child.wait(timeout=30);break
            time.sleep(.05)
        collect();code=child.wait();normal=result.exists()
        data=json.loads(result.read_text(encoding='utf8')) if normal else {'status':'FAIL','error':'OFFICE_WORKER_DIED','worker_returncode':code}
        for path in base.glob('office_QUIT_RETURNED_*.json'):
            event=json.loads(path.read_text(encoding='utf8'))
            for r in records:
                if event['worker_pid']==child.pid and event['data']['pid']==r['pid'] and event['data']['creation_time']==r['creation_time']:
                    r.update(quit_attempted=True,quit_returned=True)
                    for key in ('workbooks_before_quit','user_control_before_quit','quit_state_diagnostic_error'):
                        if key in event['data']:r[key]=event['data'][key]
        ended=time.monotonic()
        for path in base.glob('office_EMPTY_EXCEL_WM_CLOSE_POSTED_*.json'):
            event=json.loads(path.read_text(encoding='utf8'))
            for r in records:
                if event['worker_pid']==child.pid and event['data']['pid']==r['pid'] and event['data']['creation_time']==r['creation_time']:
                    r['normal_window_close_requested']=event['data']['normal_window_close_requested']
        # An owned Office child may remain unbound (no *_START record or
        # QUIT_RETURNED) even when the document worker reports PASS. Always
        # recover such a proven child after the worker has exited, using the
        # existing sealed plan and native COM cleanup helper. Never touch
        # pre-existing Office instances or skip the identity/exit checks.
        for r in records:
            h=handles.get(r['pid'])
            if r.get('quit_returned'):continue
            if h is None or win32event.WaitForSingleObject(h,0)==win32con.WAIT_OBJECT_0:continue
            from stage8_closure.storage import digest
            plan=base/('cleanup_'+str(r['pid'])+'.json')
            payload={'record':r,'parent_pid':os.getpid(),'parent_creation_time':win32process.GetProcessTimes(win32api.GetCurrentProcess())['CreationTime'].isoformat()}
            save(plan,payload|{'seal':digest(payload)})
            helper=subprocess.Popen(prefix+['--office-cleanup',str(plan)],creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0),stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
            try:helper.wait(timeout=60)
            except subprocess.TimeoutExpired:helper.terminate();helper.wait(timeout=10);r['cleanup_helper_timeout']=True
            report=plan.with_suffix('.result.json')
            r['crash_cleanup']=json.loads(report.read_text(encoding='utf8')) if report.exists() else {'status':'FAIL','helper_returncode':helper.returncode}
        deadline=time.monotonic()+120;sample_at=ended+2;window_samples=[]
        while time.monotonic()<deadline and any(win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0 for h in handles.values()):
            if time.monotonic()>=sample_at:
                living={pid for pid,h in handles.items() if win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0}
                windows=[]
                def window(hwnd,_):
                    pid=win32process.GetWindowThreadProcessId(hwnd)[1]
                    if pid in living:windows.append({'pid':pid,'hwnd':hwnd,'class':win32gui.GetClassName(hwnd),'title':win32gui.GetWindowText(hwnd),'visible':bool(win32gui.IsWindowVisible(hwnd)),'enabled':bool(win32gui.IsWindowEnabled(hwnd))})
                win32gui.EnumWindows(window,None)
                window_samples.append({'seconds_after_worker':time.monotonic()-ended,'living_owned_pids':sorted(living),'windows':windows})
                save(base/('WAIT_WINDOWS_'+str(len(window_samples))+'.json'),window_samples[-1]);sample_at=time.monotonic()+30
            time.sleep(.1)
        for r in records:
            h=handles.get(r['pid'])
            if h is None:continue
            if win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0:
                if not r.get('actual_child') or r.get('launcher_pid')!=os.getpid() or r.get('handle_source')!='CREATE_PROCESS' or win32process.GetProcessId(h)!=r['pid'] or win32process.GetProcessTimes(h)['CreationTime'].isoformat()!=r['creation_time'] or process_image(h)!=r['image']:raise RuntimeError('OFFICE_IDENTITY_CHANGED_NO_TERMINATE')
                r['forced_termination_attempted']=True
                save(base/('office_FORCE_BEGIN_'+str(r['pid'])+'.json'),dict(r,handle_validated=True,monotonic=time.monotonic()))
                try:win32api.TerminateProcess(h,1)
                except BaseException as exc:r['termination_error']=repr(exc)
                win32event.WaitForSingleObject(h,10000)
            r['exit_code']=win32process.GetExitCodeProcess(h);r['still_running']=win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0
            r['cleanup']='GRACEFUL_EXIT' if r['exit_code']==0 and not r.get('forced_termination_attempted') and not r['still_running'] else 'ABNORMAL_OR_FORCED_EXIT' if not r['still_running'] else 'STOP_OWNED_OFFICE_STILL_RUNNING'
        # Helper success is meaningful only if it observed the verified owned
        # process exit normally before returning from its own live COM apartment.
        unbound_cleanup_failed=any(
            'crash_cleanup' in r and (
                r['crash_cleanup'].get('status')!='PASS'
                or r['crash_cleanup'].get('exit_code')!=0
                or r['crash_cleanup'].get('exit_observed_while_sta_alive') is not True)
            for r in records)
        summary={'schema':'STAGE9_SUPERVISOR_V1','worker_pid':child.pid,'worker_returncode':code,
            'wait_seconds_after_worker':time.monotonic()-ended,'instances':records,'ownership_errors':ownership_errors,
            'owned_orphans':sum(r.get('still_running',False) for r in records),'normal_result':normal,'wait_window_samples':window_samples,
            'owned_unbound_gracefully_recovered':sum(
                'crash_cleanup' in r and r['crash_cleanup'].get('status')=='PASS' and r['crash_cleanup'].get('exit_code')==0
                for r in records)}
        save(base/'SUPERVISOR.json',summary)
        if ownership_errors or summary['owned_orphans'] or code!=0:data={'status':'FAIL','error':data.get('error','OFFICE_SUPERVISOR_STOP'),'worker_returncode':code,'worker_result':data}
        elif normal and any(r.get('forced_termination_attempted') for r in records):data={'status':'FAIL','error':'OFFICE_GRACEFUL_EXIT_TIMEOUT','worker_result':data}
        elif unbound_cleanup_failed:data={'status':'FAIL','error':'OFFICE_OWNED_UNBOUND_CLEANUP_FAILED','worker_result':data}
        elif any(r.get('exit_code')!=0 for r in records):data={'status':'FAIL','error':'OFFICE_ABNORMAL_EXIT','worker_result':data}
        data['office_supervisor']=str(base/'SUPERVISOR.json');published=base/'published_result.json';save(published,data)
        if target is not None:target.parent.mkdir(parents=True,exist_ok=True);os.rename(published,target)
        elif sys.stdout:print(json.dumps(data,ensure_ascii=False,indent=2))
        return 0 if data['status']=='PASS' else 1
    finally:
        for handle in handles.values():handle.Close()
