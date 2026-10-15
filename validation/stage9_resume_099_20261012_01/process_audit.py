"""Read-only audit of this resume's children; preexisting GUI is preserved."""
from common import *
from datetime import datetime
import win32api,win32process,win32event,win32security,win32con,pywintypes,collections
from stage9_app.ownership import process_image
def timestamp(s):return datetime.fromisoformat(s).timestamp()
diagnostics=Path(SETUP['base'])/'VZMKPassport/diagnostics';records={};jobs=[];forces=[];closes=[]
collisions={c['job']:c for c in read(OUT/'OFFICE_DIAGNOSTIC_COLLISION_FINAL.json')['cases']}
for folder in diagnostics.glob('job_*'):
    request=read(folder/'JOB_REQUEST.json');sup=read(folder/'SUPERVISOR.json')
    assert request['supervisor_image'].casefold()==str(EXE).casefold()
    args=request['arguments'];task_name=args[args.index('--task')+1];assert task_name!='smoke','GUI_SMOKE_FORBIDDEN'
    assert sup['owned_orphans']==0
    if sup['ownership_errors']:
        tested=read(read(OUT/'OWNERSHIP_STOP_FINAL.json')['evidence'])
        assert read(tested['result'])['office_supervisor']==str(folder/'SUPERVISOR.json')
    jobs.append({'folder':str(folder),'task':task_name,'worker_returncode':sup['worker_returncode'],'ownership_errors':sup['ownership_errors'],'instances':sup['instances']})
    for r in sup['instances']:
        records[(r['pid'],r['creation_time'])]=r
        if r.get('forced_termination_attempted'):
            assert sup['worker_returncode']==73 or sup['ownership_errors'],('UNEXPECTED_NORMAL_FORCE',str(folder))
            p=folder/('office_FORCE_BEGIN_'+str(r['pid'])+'.json');proof=read(p)
            assert proof['handle_validated'] and proof['handle_source']=='CREATE_PROCESS' and proof['retained_handle']==r['retained_handle'] and proof['creation_time']==r['creation_time']
            forces.append(str(p))
        elif not sup['ownership_errors']:assert r['exit_code']==0 and not r['still_running']
    for p in folder.glob('office_EMPTY_EXCEL_WM_CLOSE_BEGIN_*.json'):
        d=read(p)['data'];assert d['handle_validated'] and d['workbooks_before_quit']==0 and d['com_interfaces']==0 and d['quit_returned']
        closes.append(str(p))
    if (folder/'office.json').exists():
        if str(folder) in collisions:
            assert sha(folder/'office.json')==collisions[str(folder)]['reserved_final_sha256']
            assert read(collisions[str(folder)]['result'])['status']=='FAIL'
        else:
            d=read(folder/'office.json');assert d['owned_orphans']==0
            assert not d.get('diagnostic_errors'),d
for p in OUT.glob('foreign_pumped_*/SOURCE_SENTINEL_CLEANUP.json'):
    for r in read(p)['instances']:records[(r['pid'],r['creation_time'])]=r
living=set(win32process.EnumProcesses());observations=[];orphans=[];unknown=[]
for (pid,birth),r in records.items():
    row={'pid':pid,'birth':birth,'image':r['image']}
    if pid not in living:row['state']='GONE'
    else:
        try:
            h=win32api.OpenProcess(0x100400,False,pid)
            try:
                if win32event.WaitForSingleObject(h,0)==0:row['state']='EXITED'
                else:
                    same=abs(timestamp(win32process.GetProcessTimes(h)['CreationTime'].isoformat())-timestamp(birth))<.001 and process_image(h).casefold()==r['image'].casefold()
                    row['state']='OWNED_ALIVE' if same else 'PID_REUSED_OR_EXITED'
            finally:h.Close()
        except (pywintypes.error,OSError) as exc:
            row['state']='GONE' if getattr(exc,'winerror',None)==87 else 'UNKNOWN';row['query_error']=repr(exc)
            if row['state']=='UNKNOWN':
                command=("$taskProcess=Get-CimInstance Win32_Process -Filter 'ProcessId = "+str(int(pid))+"'; "
                    "if ($taskProcess) { [pscustomobject]@{ProcessId=$taskProcess.ProcessId;Name=$taskProcess.Name;CreationDate=$taskProcess.CreationDate.ToString('o')} | ConvertTo-Json -Compress }")
                query=subprocess.run([str(Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe'),'-NoProfile','-Command',command],capture_output=True,text=True,timeout=30)
                if query.returncode==0:
                    if not query.stdout.strip():row['state']='GONE_AFTER_QUERY'
                    else:
                        other=json.loads(query.stdout);row['CIM']=other
                        if other['ProcessId']==pid and timestamp(other['CreationDate'])>timestamp(birth)+1:row['state']='PID_REUSED_CIM'
    if row['state']=='OWNED_ALIVE':orphans.append(row)
    if row['state']=='UNKNOWN':unknown.append(row)
    observations.append(row)
baseline=[]
for r in read(OUT/'BASELINE_PROCESSES.json'):
    row=dict(r)
    try:
        h=win32api.OpenProcess(0x100400,False,r['pid'])
        try:row['same_identity_alive']=win32event.WaitForSingleObject(h,0)==258 and abs(timestamp(win32process.GetProcessTimes(h)['CreationTime'].isoformat())-timestamp(r['creation_time']))<.001 and process_image(h).casefold()==r['image'].casefold()
        finally:h.Close()
    except (pywintypes.error,OSError) as exc:row['same_identity_alive']=False;row['observation_error']=repr(exc)
    row['controlled_by_this_resume']=False;baseline.append(row)
token=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY)
try:elevation=win32security.GetTokenInformation(token,win32security.TokenElevation)
finally:token.Close()
assert elevation==0 and not orphans and not unknown,(elevation,orphans,unknown)
result={'status':'PASS','exe_sha256':sha(EXE),'job_count':len(jobs),'office_instances':sum(len(j['instances']) for j in jobs),
    'task_counts':dict(collections.Counter(j['task'] for j in jobs)),'jobs':jobs,'force_proofs':forces,'normal_force_count':0,
    'WM_CLOSE_events':closes,'owned_orphans':orphans,'unknown_child_identities':unknown,'observations':observations,
    'baseline_preexisting_processes':baseline,'preexisting_GUI_controlled':False,'GUI_E2E':'NOT_RUN_USER_STOP','token_elevation':elevation}
save(OUT/'FINAL_PROCESS_AUDIT.json',result)
print('Jobs',len(jobs),'Office',result['office_instances'],'new child orphans 0','normal force 0','GUI not controlled')
