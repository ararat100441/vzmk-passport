"""Continue the sealed 0.9.12 legacy checkpoint; never regenerate old bundles."""
import copy, ctypes, hashlib, json, os, shutil, subprocess, sys, time, traceback
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4
import win32api, win32con, win32event, win32process

OUT=Path(__file__).resolve().parent
ROOT=OUT.parents[1]
PREV=ROOT/'validation/akz_settings_0911_20261013_01'
sys.path.insert(0,str(ROOT))
from stage9_app.ownership import process_image
from stage9_app import akz_auto as A
from stage8_closure import storage as S

def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:
        json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
def utc():return datetime.now(timezone.utc).isoformat()
def check_map(mapping):
    for p,h in mapping.items():assert sha(p)==(h['sha256'] if isinstance(h,dict) else h),(p,'SHA_CHANGED')

BUILD=read(OUT/'BUILD_MANIFEST.json');EXE=Path(BUILD['exe'])
OLD_SEAL='995c3fbc174695afe34314cfe14cece6237467ba8f2c53a70b0dee1401baa555'
LEGACY=Path(read(PREV/'WORKSPACE_LEGACY.json')['workspace'])
MAIN=Path(read(PREV/'WORKSPACE_MAIN.json')['workspace'])
MANIFEST=Path(read(PREV/'OLD_LEGACY_generate_e81a2d29.json')['manifest'])
PREVIEW=read(PREV/'OLD_LEGACY_check_9474b7d8.json')['preview']
AB=read(PREV/'AB_FINAL.json')
OLD=read(PREV/'SETUP.json')
BASE=Path(OLD['parent'])/'resume_local_0912_c1389d7a'
assert BASE.is_dir() and (BASE/'tmp').is_dir()
ENV=os.environ.copy();ENV['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot']
ENV.pop('PYTHONPATH',None);ENV.pop('PYTHONHOME',None)
ENV.update(LOCALAPPDATA=str(BASE),TEMP=str(BASE/'tmp'),TMP=str(BASE/'tmp'),PYTHONUTF8='1')
JOBS=[];BASELINE=[];HANDLES=[];STOPPED=False

def bind():
    assert sha(EXE)==BUILD['exe_sha256']
    assert {str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}==BUILD['files']
    check_map(read(OUT/'SOURCE_AT_BUILD.json'));check_map(read(PREV/'ACCEPTED_BEFORE.json'))
    assert sha(PREV/'EVIDENCE_MANIFEST.json')==OLD_SEAL
    seal=read(PREV/'EVIDENCE_MANIFEST.json')
    check_map({str(PREV/k):v for k,v in seal['evidence_files'].items()})
    changes={r['path']:r for r in read(OUT/'CHANGED_FILES.json') if r['before_sha256']}
    cache_drift=[]
    for p,v in seal['external_files'].items():
        if p in changes:
            assert sha(changes[p]['original_copy'])==v['sha256']
            assert sha(p)==changes[p]['after_sha256']
        elif Path(p).parent==Path(OLD['parent'])/'LocalLow/Intel/ShaderCache':
            if sha(p)!=v['sha256']:cache_drift.append({'path':p,'old_sha256':v['sha256'],'current_sha256':sha(p),'classification':'MUTABLE_DRIVER_CACHE_OUTSIDE_SOURCES_AND_BUNDLES'})
        else:assert sha(p)==v['sha256'],(p,'OLD_EXTERNAL_CHANGED')
    old=read(PREV/'BUILD_MANIFEST.json');check_map({str(Path(old['exe']).parent/p):h for p,h in old['files'].items()})
    return {'status':'PASS','at':utc(),'exe':str(EXE),'exe_sha256':sha(EXE),'onedir_files':len(BUILD['files']),
            'source_at_build_files':len(read(OUT/'SOURCE_AT_BUILD.json')),'accepted_stage1_8':len(read(PREV/'ACCEPTED_BEFORE.json')),
            'prior_seal_sha256':OLD_SEAL,'prior_evidence_files':len(seal['evidence_files']),'prior_external_files':len(seal['external_files']),
            'authorized_source_changes_preserved_before_after':list(changes),'old_bundles_and_0911_build_unchanged':True,'prior_wide_seal_runtime_cache_drift':cache_drift}

def inventory():
    code="@(Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE' OR Name='WINWORD.EXE'\" | Select-Object Name,ProcessId,ParentProcessId,ExecutablePath,@{n='creation_time';e={$_.CreationDate.ToUniversalTime().ToString('o')}}) | ConvertTo-Json -Depth 4 -Compress"
    r=subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-Command',code],capture_output=True,timeout=20)
    assert r.returncode==0,r.stderr
    raw=r.stdout.decode('utf-8-sig').strip();d=json.loads(raw) if raw else []
    return d if isinstance(d,list) else [d]

def capture():
    for r in inventory():
        h=win32api.OpenProcess(0x1000|win32con.SYNCHRONIZE,False,r['ProcessId'])
        HANDLES.append(h);born=win32process.GetProcessTimes(h)['CreationTime'].isoformat();image=process_image(h)
        assert abs(datetime.fromisoformat(born).timestamp()-datetime.fromisoformat(r['creation_time']).timestamp())<.002
        assert image.casefold()==r['ExecutablePath'].casefold()
        r.update(creation_time=born,image=image,handle=int(h),handle_access='QUERY_LIMITED_INFORMATION|SYNCHRONIZE',terminate_access=False)
        BASELINE.append(r)
    save(OUT/'OFFICE_BASELINE.json',{'at':utc(),'processes':BASELINE,'retained_readonly_handles':True})

def living():
    global STOPPED
    for r,h in zip(BASELINE,HANDLES):
        if win32event.WaitForSingleObject(h,0)==win32con.WAIT_OBJECT_0:
            STOPPED=True
            if not (OUT/'FOREIGN_OFFICE_STOP.json').exists():
                save(OUT/'FOREIGN_OFFICE_STOP.json',{'at':utc(),'identity':r,'exit_code':win32process.GetExitCodeProcess(h),'cause':'UNKNOWN','reason':'BASELINE_OFFICE_EXIT'})
            raise RuntimeError('FOREIGN_OFFICE_EXIT_STOP')
        assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time']
        assert process_image(h)==r['image']

def parent_pid(handle):
    info=(ctypes.c_void_p*6)();length=ctypes.c_ulong()
    fn=ctypes.WinDLL('ntdll').NtQueryInformationProcess
    fn.argtypes=[ctypes.c_void_p,ctypes.c_ulong,ctypes.c_void_p,ctypes.c_ulong,ctypes.c_void_p]
    fn.restype=ctypes.c_long
    assert fn(int(handle),0,ctypes.byref(info),ctypes.sizeof(info),ctypes.byref(length))==0
    return info[5]

def cancel_owned_worker(proc,result_path):
    """Only on STOP: prove a child via retained handles; leave supervisor to clean Office."""
    entries=[]
    for p in (BASE/'VZMKPassport/diagnostics').glob('job_*/JOB_REQUEST.json'):
        d=read(p)
        if d['supervisor_pid']==proc.pid and str(result_path) in d['arguments']:entries.append((p,d))
    assert len(entries)==1,'NO_UNIQUE_OWN_JOB_NO_TERMINATE'
    p,request=entries[0]
    assert process_image(proc._handle).casefold()==str(EXE).casefold()
    birth=win32process.GetProcessTimes(proc._handle)['CreationTime'].timestamp()
    assert request['created_at']>=birth and request['supervisor_image'].casefold()==str(EXE).casefold()
    candidates=set()
    for evidence in [*p.parent.glob('launch_*.json'),*p.parent.glob('office_*_START.json')]:
        d=read(evidence)
        if 'worker_pid' in d:candidates.add(d['worker_pid'])
    for pid in candidates:
        h=win32api.OpenProcess(0x1000|win32con.SYNCHRONIZE|win32con.PROCESS_TERMINATE,False,pid)
        try:
            identity={'pid':pid,'parent':parent_pid(h),'birth':win32process.GetProcessTimes(h)['CreationTime'].isoformat(),'image':process_image(h)}
            assert identity['parent']==proc.pid and datetime.fromisoformat(identity['birth']).timestamp()>=birth
            assert identity['image'].casefold()==str(EXE).casefold()
            if win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0:
                save(OUT/('STOP_OWN_WORKER_'+str(pid)+'.json'),{'identity':identity,'verified_handle':int(h),'supervisor':proc.pid,'foreign_or_Office_terminated':False})
                win32api.TerminateProcess(h,73);win32event.WaitForSingleObject(h,10000)
        finally:h.Close()
    if not candidates:raise RuntimeError('WORKER_IDENTITY_UNPROVEN_NO_TERMINATE')
    proc.wait(timeout=160)

def task(label,name,expected='PASS',**kwargs):
    living();assert sha(EXE)==BUILD['exe_sha256']
    p=OUT/(label+'_'+name+'_'+uuid4().hex[:8]+'.json')
    cmd=[str(EXE),'--task',name,'--result',str(p)]
    for k,v in kwargs.items():
        cmd+=['--'+k.replace('_','-')]
        if v is not True:cmd+=[str(v)]
    save(p.with_suffix('.command.json'),{'command':cmd,'exe_sha256':sha(EXE),'runtime':'FROZEN_EXE_0.9.12','PATH':ENV['PATH'],'LOCALAPPDATA':str(BASE),'PYTHONPATH':ENV.get('PYTHONPATH'),'PYTHONHOME':ENV.get('PYTHONHOME'),'expected':expected,'started_at':utc()})
    # Poll retained foreign handles during the job, not only after completion.
    with p.with_suffix('.stdout.log').open('xb') as stdout,p.with_suffix('.stderr.log').open('xb') as stderr:
        proc=subprocess.Popen(cmd,env=ENV,stdout=stdout,stderr=stderr,creationflags=subprocess.CREATE_NO_WINDOW)
        start=time.monotonic()
        try:
            while proc.poll() is None:
                living()
                if time.monotonic()-start>600:raise RuntimeError('OWN_JOB_TIMEOUT_STOP')
                time.sleep(.1)
        except BaseException:
            try:cancel_owned_worker(proc,p)
            except BaseException as e:save(p.with_suffix('.cancel.json'),{'error':repr(e),'traceback':traceback.format_exc(),'unknown_process_not_terminated':True})
            raise
    save(p.with_suffix('.process.json'),{'returncode':proc.returncode,'supervisor_pid':proc.pid,'ended_at':utc(),'poll_interval_seconds':.1})
    living();d=read(p);s=read(d['office_supervisor'])
    assert s['owned_orphans']==0 and not s['ownership_errors'],s
    assert all(not i.get('forced_termination_attempted') and i['exit_code']==0 and not i['still_running'] for i in s['instances']),s
    assert d['status']==expected and (proc.returncode==0)==(expected=='PASS'),d
    JOBS.append({'label':label,'task':name,'result':str(p),'expected':expected,'actual':d['status'],'supervisor':d['office_supervisor'],'office_instances':len(s['instances']),'own_orphans':0,'force':0})
    print(label,name,d['status'],'Office',len(s['instances']),'orphan0',flush=True)
    return d,p

def main():
    binding=bind()
    if not (OUT/'INITIAL_BINDING.json').exists():save(OUT/'INITIAL_BINDING.json',binding)
    if not (OUT/'SETUP.json').exists():save(OUT/'SETUP.json',{'previous_evidence':str(PREV),'legacy_workspace':str(LEGACY),'main_workspace':str(MAIN),'legacy_manifest':str(MANIFEST),'legacy_preview':PREVIEW,'base':str(BASE),'temp':str(BASE/'tmp'),'GUI':'NOT_RUN_NO_PERMISSION','minimal_proven_legacy_crlf_fix':True})
    diag=read(PREV/'OFFICE_STOP_DIAGNOSIS.json')
    if not (OUT/'PID48680_REVIEW.json').exists():save(OUT/'PID48680_REVIEW.json',{'status':'INVESTIGATED_CAUSE_UNKNOWN','identity':diag['foreign']['identity'],'observed_at_utc':diag['observed_at_utc'],'exit_code':diag['foreign']['exit_code'],'cause':'UNKNOWN','attribution':'Exit0 and absent correlated Application event do not establish the closing actor or cause.','in_prior_own_instance_lists':diag['foreign_pid_in_own_instance_lists'],'prior_own_orphans':diag['owned_orphans'],'prior_own_force':diag['own_forced_terminations'],'prior_monitor_limitation':diag['monitor_limitation'],'new_controller_polls_during_wait':True,'no_foreign_process_actions':True,'previous_report_preserved':str(PREV/'REPORT.md')})
    capture();living()
    path=LEGACY/'preferences/akz_auto/settings.json';raw=path.read_bytes();old_sha=sha(path)
    m=read(MANIFEST);assert 'akz_settings_snapshot' not in m
    before={str(MANIFEST):sha(MANIFEST),**{r['path']:sha(r['path']) for r in m['records']},**{str(p):sha(p) for p in MANIFEST.parent.rglob('PAIR.json')}}
    days={str(p):sha(p) for p in path.parent.glob('days/*.json')}
    assert days and all('settings_digest_schema' not in read(p) for p in days)
    for key in ['A_files','B_files']:check_map(AB[key])
    save(OUT/'IMMUTABLE_BEFORE.json',{'legacy_documents_manifests_receipts':before,'legacy_days':days,'A_files':AB['A_files'],'B_files':AB['B_files'],'live_settings_original_sha256':old_sha})
    (OUT/'LEGACY_LIVE_SETTINGS_BEFORE.json').write_bytes(raw)
    hold=None;source=None
    try:
        equivalent,_=task('LEGACY_EQUIVALENT_SAVE','akz-settings-save',workspace=LEGACY,akz_settings=PREV/'LEGACY_FLOAT_SETTINGS.json')
        assert equivalent.get('unchanged') is True and sha(path)==old_sha,equivalent
        generated,gp=task('LEGACY_FLOAT_DAY_REUSE','akz-auto-generate',workspace=LEGACY,preview=PREVIEW,confirm_test=True)
        check_map(days);check_map(before)
        cfg=copy.deepcopy(read(PREV/'LEGACY_FLOAT_SETTINGS.json'));cfg['climate']['air_c']=[16,26]
        save(OUT/'LEGACY_FUTURE_SETTINGS.json',cfg)
        task('LEGACY_CHANGE_FUTURE','akz-settings-save',workspace=LEGACY,akz_settings=OUT/'LEGACY_FUTURE_SETTINGS.json')
        verified,_=task('LEGACY_HISTORY','verify',workspace=LEGACY,manifest=MANIFEST)
        proof=verified['settings_validation'];assert proof['status']=='LEGACY_SETTINGS_RECOVERED_EXACT_SHA'
        assert proof['sha256']==old_sha
        source=Path(proof['source']);assert source.parent==path.parent/'history'
        source_sha=sha(source);hold=source.parent.parent/('evidence_hold_'+uuid4().hex+'.json')
        source.rename(hold)
        try:
            denied,dp=task('LEGACY_UNPROVEN','verify','FAIL',workspace=LEGACY,manifest=MANIFEST)
            assert 'LEGACY_AKZ_SETTINGS_UNPROVEN' in denied['error'] and 'TEST' in denied['error']
        finally:
            hold.rename(source);hold=None;assert sha(source)==source_sha
        restarted,_=task('LEGACY_AFTER_RESTART','verify',workspace=LEGACY,manifest=MANIFEST)
        assert restarted['settings_validation']['status']=='LEGACY_SETTINGS_RECOVERED_EXACT_SHA'
        for label,key in [('FROZEN_A_CURRENT_B','A_manifest'),('FROZEN_B_CURRENT_B','B_manifest')]:
            result,_=task(label,'verify',workspace=MAIN,manifest=AB[key])
            assert result['settings_validation']['status']=='FROZEN_SETTINGS_VERIFIED'
        check_map(before);check_map(days)
        for key in ['A_files','B_files']:check_map(AB[key])
        save(OUT/'LEGACY_FINAL.json',{'status':'PASS','legacy_created_by':'PRESERVED_FROZEN_0.9.10','validated_with':'FROZEN_0.9.12','manifest':str(MANIFEST),'legacy_settings_recovery':proof,'unproven_result':str(dp),'numeric_reuse_result':str(gp),'all_variants':['0','1','2','3'],'old_documents_manifests_receipts_unchanged':before,'old_day_files_unchanged':days,'history_restored_exactly':str(source),'A_B_frozen_files_unchanged':True,'GUI':'NOT_RUN','no_old_outputs_regenerated':True})
    finally:
        if hold is not None and hold.exists():hold.rename(source)
        # Restore only owned TEST live preferences, preserving the prior evidence seal.
        if sha(path)!=old_sha:
            with A.settings_lock(LEGACY):
                temp=path.with_name('restore_'+uuid4().hex+'.tmp');temp.write_bytes(raw);os.replace(temp,path)
        assert sha(path)==old_sha
        save(OUT/'LIVE_SETTINGS_RESTORED.json',{'path':str(path),'sha256':old_sha,'exact_original_bytes_restored':True,'frozen_manifests_receipts_touched':False})
    living()
    save(OUT/'FINAL_BINDING.json',bind())
    final=inventory();owned=[]
    for job in JOBS:
        summary=read(job['supervisor']);owned.extend(summary['instances'])
        for f in Path(job['supervisor']).parent.rglob('*'):
            if f.is_file():
                dst=OUT/'office_jobs'/f.parent.name/f.name;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(f,dst);assert sha(f)==sha(dst)
    assert not ({r['ProcessId'] for r in final}&{r['pid'] for r in owned})
    save(OUT/'OFFICE_FINAL.json',{'status':'PASS','at':utc(),'baseline_processes':[dict(r,alive=True) for r in BASELINE],'current_processes':final,'foreign_stop':False,'foreign_processes_closed':0,'own_office_instances':len(owned),'own_orphans':0,'force':0,'jobs':JOBS})
    save(OUT/'CLI_CASES.json',{'status':'PASS','runtime':'FROZEN_EXE_0.9.12','new_cases':len(JOBS),'cases':JOBS,'prior_frozen_checks':21,'no_Python_in_PATH':True,'GUI':'NOT_RUN'})
    print('RESUME_VALIDATION_PASS',len(JOBS),'CLI cases',len(owned),'native Office instances',flush=True)

if __name__=='__main__':
    try:main()
    except BaseException as exc:
        save(OUT/('STOP_'+uuid4().hex[:8]+'.json'),{'at':utc(),'error':repr(exc),'traceback':traceback.format_exc(),'completed_jobs':JOBS,'foreign_stop':STOPPED});raise
    finally:
        for h in HANDLES:h.Close()
