from common import *
import ctypes,win32api,win32process,win32event
from ctypes import wintypes as W
ds=Path(SETUP['base'])/'VZMKPassport/diagnostics'
jobs=[];identities={};normal_force=[];window_close=[];unknown=[];live=[];resolved=[]
for p in sorted(ds.glob('job_*/SUPERVISOR.json')):
    s=read(p);q=read(p.parent/'JOB_REQUEST.json')
    assert Path(q['supervisor_image']).resolve()==EXE.resolve(),q
    assert s['owned_orphans']==0,s
    if s['ownership_errors']:assert any('OWNERSHIP' in e['error'] for e in s['ownership_errors'])
    jobs.append({'path':str(p),'arguments':q['arguments'],'worker_returncode':s['worker_returncode'],'ownership_errors':s['ownership_errors'],'instances':len(s['instances'])})
    for r in s['instances']:
        assert r['actual_child'] and r['owned'] and r['handle_source']=='CREATE_PROCESS'
        if r.get('forced_termination_attempted'):
            proof=read(p.parent/('office_FORCE_BEGIN_'+str(r['pid'])+'.json'))
            assert proof['handle_validated'] and all(proof[k]==r[k] for k in ('pid','retained_handle','creation_time','image','launcher_pid'))
            if s['worker_returncode']==0:normal_force.append({'job':str(p),'record':r})
        if r.get('normal_window_close_requested'):window_close.append({'job':str(p),'record':r,'worker_returncode':s['worker_returncode']})
        identities[(r['pid'],r['creation_time'])]=r
for p in ds.glob('viewer_*.json'):
    d=read(p)
    if d.get('schema')!='STAGE9_OFFICE_OWNERSHIP_V2':continue
    assert d['owned_orphans']==0,d
    for r in d['instances']:identities[(r['pid'],r['creation_time'])]=r
for p in list(OUT.glob('foreign_office_*/SOURCE_SENTINEL_CLEANUP.json'))+list(OUT.glob('foreign_pumped_*/SOURCE_SENTINEL_CLEANUP.json')):
    d=read(p);assert d['owned_orphans']==0
    for r in d['instances']:identities[(r['pid'],r['creation_time'])]=r
if (OUT/'GUI_LAUNCH.json').exists():
    r=read(OUT/'GUI_LAUNCH.json');identities[(r['pid'],r['creation_time'])]=dict(r,image=r['exe'])
query=ctypes.WinDLL('kernel32',use_last_error=True).QueryFullProcessImageNameW
query.argtypes=[W.HANDLE,W.DWORD,W.LPWSTR,ctypes.POINTER(W.DWORD)];query.restype=W.BOOL
for (pid,birth),r in identities.items():
    try:h=win32api.OpenProcess(0x100400,False,pid)
    except __import__('pywintypes').error as exc:
        if exc.winerror==87:resolved.append({'pid':pid,'creation_time':birth,'state':'ABSENT'});continue
        command="Get-CimInstance Win32_Process -Filter 'ProcessId = "+str(pid)+"' | Select-Object ProcessId,Name,@{Name='CreationTime';Expression={$_.CreationDate.ToUniversalTime().ToString('o')}} | ConvertTo-Json -Compress"
        cim=subprocess.run([os.environ['SystemRoot']+'\\System32\\WindowsPowerShell\\v1.0\\powershell.exe','-NoProfile','-Command',command],capture_output=True,timeout=30)
        try:
            current=json.loads(cim.stdout.decode('utf8-sig'))
            from datetime import datetime
            if datetime.fromisoformat(current['CreationTime'])>datetime.fromisoformat(birth):
                resolved.append({'pid':pid,'old_birth':birth,'current':current,'state':'PID_REUSED_CIM_NOT_TOUCHED'});continue
        except Exception:pass
        unknown.append({'pid':pid,'creation_time':birth,'error':repr(exc),'cim':cim.stdout.decode('utf8','replace')});continue
    try:
        current=win32process.GetProcessTimes(h)['CreationTime'].isoformat();b=ctypes.create_unicode_buffer(32768);n=W.DWORD(len(b))
        if not query(int(h),0,b,ctypes.byref(n)):raise ctypes.WinError(ctypes.get_last_error())
        if current!=birth:resolved.append({'pid':pid,'old_birth':birth,'new_birth':current,'new_image':b.value,'state':'PID_REUSED_NOT_TOUCHED'});continue
        if win32event.WaitForSingleObject(h,0)==258:live.append({'pid':pid,'birth':birth,'image':b.value})
        else:resolved.append({'pid':pid,'creation_time':birth,'state':'SIGNALED'})
    except Exception as exc:unknown.append({'pid':pid,'creation_time':birth,'error':repr(exc)})
    finally:h.Close()
result={'status':'PASS' if not live and not unknown and not normal_force else 'FAIL','jobs':jobs,'identities':len(identities),'live_owned':live,'unknown_identities':unknown,'normal_forces':normal_force,'WM_CLOSE_cases':window_close,'resolved':resolved,'no_process_termination_by_this_audit':True}
save(OUT/'FINAL_PROCESS_AUDIT.json',result)
print(json.dumps({k:result[k] for k in ('status','identities','live_owned','unknown_identities','normal_forces')},ensure_ascii=False),flush=True)
assert result['status']=='PASS',result
