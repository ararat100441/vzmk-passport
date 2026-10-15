"""Read only process identities and isolated evidence; never opens production."""
from common import *
import win32api,win32process,win32event,pywintypes,collections
from stage9_app.ownership import process_image
diagnostics=Path(SETUP['base'])/'VZMKPassport/diagnostics'
identities={};current=[];force=[];incomplete=[]
for job in diagnostics.glob('job_*'):
    for path in job.glob('created_*.json'):
        r=read(path);identities[(r['pid'],r['creation_time'])]=r
    request=job/'JOB_REQUEST.json';summary=job/'SUPERVISOR.json'
    if request.exists() and read(request)['supervisor_image'].casefold()==str(EXE).casefold():
        if not summary.exists():incomplete.append(str(job));continue
        r=read(summary);args=read(request)['arguments']
        current.append({'job':str(job),'task':args[args.index('--task')+1],
            'probe_mode':args[args.index('--probe-mode')+1] if '--probe-mode' in args else None,
            'test_fault':args[args.index('--test-fault')+1] if '--test-fault' in args else None,
            'worker_returncode':r['worker_returncode'],'normal_result':r['normal_result'],
            'ownership_errors':r['ownership_errors'],'owned_orphans':r['owned_orphans'],
            'instances':r['instances']})
        assert r['owned_orphans']==0,r
        for instance in r['instances']:
            if instance.get('forced_termination_attempted'):
                assert r['worker_returncode']==73 or r['ownership_errors'],('UNEXPECTED_NORMAL_FORCE',r)
                proof=read(job/('office_FORCE_BEGIN_'+str(instance['pid'])+'.json'))
                assert proof['handle_validated'] and proof['handle_source']=='CREATE_PROCESS' and proof['retained_handle']==instance['retained_handle']
                force.append({'job':str(job),'pid':instance['pid'],'proof':str(job/('office_FORCE_BEGIN_'+str(instance['pid'])+'.json'))})
            elif not r['ownership_errors']:assert instance['exit_code']==0 and not instance['still_running'],instance
for path in OUT.glob('launch_protocol_*/*_CREATED.json'):
    r=read(path);identities[(r['pid'],r['creation_time'])]=r
for path in OUT.glob('foreign_office_*/SOURCE_SENTINEL_CLEANUP.json'):
    for r in read(path)['instances']:identities[(r['pid'],r['creation_time'])]=r
gui=read(OUT/'GUI_LAUNCH_097.json')
identities[(gui['pid'],gui['creation_time'])]={'pid':gui['pid'],'creation_time':gui['creation_time'],'image':str(EXE)}
fallback=read(OUT/'READONLY_PID_REUSE_CIM.json')
living=set(win32process.EnumProcesses());orphans=[];unknown=[];checked=[]
for (pid,birth),r in identities.items():
    observation={'pid':pid,'creation_time':birth,'image':r['image']}
    if pid not in living:observation['state']='GONE'
    else:
        try:
            h=win32api.OpenProcess(0x100400,False,pid)
            try:
                same=win32process.GetProcessTimes(h)['CreationTime'].isoformat()==birth and process_image(h).casefold()==r['image'].casefold()
                alive=win32event.WaitForSingleObject(h,0)==258
                observation['state']='OWNED_ALIVE' if same and alive else 'PID_REUSED_OR_EXITED'
                if same and alive:orphans.append(observation)
            finally:h.Close()
        except pywintypes.error as exc:
            observation['state']='GONE' if exc.winerror==87 else 'UNKNOWN';observation['query_error']=repr(exc)
            if observation['state']=='UNKNOWN':
                import json
                command=("$taskProcess=Get-CimInstance Win32_Process -Filter 'ProcessId = "+str(int(pid))+"'; "
                    "if ($taskProcess) { [pscustomobject]@{ProcessId=$taskProcess.ProcessId;Name=$taskProcess.Name;CreationDate=$taskProcess.CreationDate.ToString('o')} | ConvertTo-Json -Compress }")
                ps=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe'
                query=subprocess.run([str(ps),'-NoProfile','-Command',command],capture_output=True,text=True,timeout=30)
                if query.returncode==0:
                    if query.stdout.strip():fallback[str(pid)]=json.loads(query.stdout)
                    else:observation['state']='GONE_AFTER_QUERY'
            if observation['state']=='UNKNOWN' and str(pid) in fallback:
                from datetime import datetime
                other=fallback[str(pid)]
                if (other['ProcessId']==pid and Path(other['Name']).name.casefold()!=Path(r['image']).name.casefold()
                    and datetime.fromisoformat(other['CreationDate']).timestamp()>datetime.fromisoformat(birth).timestamp()):
                    observation['state']='PID_REUSED_CIM';observation['cim_evidence']=other
            if observation['state']=='UNKNOWN':unknown.append(observation)
    checked.append(observation)
save(OUT/'READONLY_PID_REUSE_CIM_FINAL.json',fallback)
assert not orphans and not unknown and not incomplete,(orphans,unknown,incomplete)
wmclose=[]
for row in current:
    for path in Path(row['job']).glob('office_EMPTY_EXCEL_WM_CLOSE_BEGIN_*.json'):
        d=read(path)['data'];assert d['handle_validated'] and d['com_interfaces']==0 and d['workbooks_before_quit']==0 and d['quit_returned'] and d['message']=='WM_CLOSE'
        wmclose.append(str(path))
result={'status':'PASS','exe_sha256':sha(EXE),'current_exe_jobs':len(current),'current_office_instances':sum(len(r['instances']) for r in current),
    'current_tasks':dict(collections.Counter(r['task'] for r in current)),'current_jobs':current,
    'emergency_force':force,'normal_window_close_events':wmclose,'normal_window_close_dynamic_branch_exercised':bool(wmclose),
    'all_run_identities_checked':len(checked),'identities':checked,'owned_orphans':orphans,'unknown_identities':unknown,
    'incomplete_current_jobs':incomplete,'gui_pid_exited':gui['pid'],
    'actual_user_session_continuity':'NOT_VERIFIED across whole long run: baseline 22596 and later 13368 ended while user worked; neither identity belongs to product child records',
    'controlled_foreign_office':read(OUT/'FOREIGN_OFFICE_FINAL.json'),'production_registry_opened':False}
import win32security,win32con
token=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32con.TOKEN_QUERY)
try:result['token_elevation']=win32security.GetTokenInformation(token,win32security.TokenElevation)
finally:token.Close()
assert result['token_elevation']==0,'ELEVATED_TEST_NOT_ACCEPTED'
source_copies={rel:sha(SOURCE/rel) for rel in SETUP['copied']}
assert all(source_copies[rel]==binding['sha256'] for rel,binding in SETUP['copied'].items())
save(OUT/'DEV_SOURCE_COPY_INTEGRITY.json',{'status':'PASS','source':str(SOURCE),'files':source_copies,'working_registry_included':False})
save(OUT/'FINAL_PROCESS_AUDIT_COMPLETE.json',result)
save(OUT/'GUI_UI_REVIEW_FINAL.json',{'status':'PARTIAL','exe_sha256':sha(EXE),'pid':gui['pid'],'main_window_observed':True,'dev_workspace_loaded':True,
    'project':'Гараж Казахстан','passport':'9/2026','akz_selected':[0],'gui_closed_normally':True,'gui_viewer_test':'NOT_RUN','full_gui_click_e2e':'NOT_RUN',
    'limitation':'Computer Use type_text did not populate Tk entry; native file dialog elements 156/155 rejected as unavailable in cached app state after fresh captures. Dialog cancelled; no Check/Generate invoked. Clipboard text inserted during diagnosis was discarded on closing GUI and was not used as a path or instruction.'})
save(OUT/'VISUAL_QUALITY_REVIEW_FINAL.json',{'status':'PASS','scope':'Only first generated quality document pages 1 and 2 visually inspected; not every generated Word page',
    'manifest':read(OUT/'EXE_E2E_FINAL.json')['manifests'][0],'checks':['TEST/DRAFT prominently visible','9/2026 number and dates readable','mass 15822.14','both VIK act numbers in sections 13/14','no clipped text observed','signature empty']})
print('Current jobs',len(current),'Office instances',result['current_office_instances'],'all identities',len(checked),'orphans 0','WM_CLOSE cases',len(wmclose))
