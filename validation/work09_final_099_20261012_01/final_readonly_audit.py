"""Finalize the stopped independent review using reads and own evidence writes only.
No Office COM connection, GUI launch, process termination, or product mutation.
"""
import concurrent.futures, hashlib, json, os, time
from collections import Counter
from datetime import datetime
from pathlib import Path
import win32api, win32event, win32process

OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
def read(p): return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
    return h.hexdigest()
def save(p,obj):
    with Path(p).open('x',encoding='utf8') as f:
        json.dump(obj,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
def key(r): return (r['pid'],r['creation_time'],r['image'].casefold())

setup=read(OUT/'SETUP.json'); base=Path(setup['base']); parent=Path(setup['parent'])
stop=read(OUT/'FOREIGN_OFFICE_STOP.json'); assert stop['status']=='STOP'
monitor=read(OUT/'OFFICE_MONITOR_FINAL.json')
mon={key(r):r for r in monitor['instances']}
e=read(OUT/'EXE_E2E_FINAL.json'); detail=read(e['evidence'])
jobs=[]; offices=[]; wm=[]; forces=[]; probes=[]; ends=[]
for s in detail['steps']:
    result_path=Path(s['result']); result=read(result_path); sup_path=Path(s['supervisor'])
    sup=read(sup_path); job=sup_path.parent; req=read(job/'JOB_REQUEST.json')
    command=read(result_path.with_suffix('.command.json'))
    proc=read(result_path.with_suffix('.process.json')); ends.append(proc['ended'])
    assert result['status']==s['expected'] and (proc['returncode']==0)==(s['expected']=='PASS')
    assert command['exe_sha256']==e['exe_sha256'] and command['command'][0]==req['supervisor_image']
    assert command['PATH'].lower()==(os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot']).lower()
    assert command['LOCALAPPDATA']==str(base)
    assert sup['owned_orphans']==0 and not sup['ownership_errors']
    if sup['worker_returncode']!=73:assert len(list(job.glob('office.json')))==1
    for r in sup['instances']:
        created=read(job/f"created_{r['pid']}.json"); started=read(job/f"office_{r['pid']}_START.json")
        for rr in (created,started):
            assert key(rr)==key(r) and rr['actual_child'] is True and rr['owned'] is True
            assert rr['launcher_pid']==req['supervisor_pid'] and rr['worker_pid']==sup['worker_pid']
            assert rr['handle_source']=='CREATE_PROCESS' and rr['retained_handle']==r['retained_handle']
        observed=mon[key(r)]
        assert observed['parent_pid']==r['launcher_pid'] and observed['exit_observed']
        assert observed['exit_code']==r['exit_code']==0 and r['still_running'] is False
        assert not r.get('forced_termination_attempted')
        h=None
        try:
            h=win32api.OpenProcess(0x1000|0x100000,False,r['pid'])
            signaled=win32event.WaitForSingleObject(h,0)==0
            if signaled: current={'state':'EXITED_OR_PID_REUSED_EXITED'}
            else:
                birth=win32process.GetProcessTimes(h)['CreationTime'].isoformat()
                assert birth!=r['creation_time'],('OWN_ORPHAN',r)
                current={'state':'PID_REUSED','creation_time':birth}
        except Exception as exc:
            if getattr(exc,'winerror',None) in (87,1168):current={'state':'ABSENT'}
            else:raise
        finally:
            if h is not None:h.Close()
        offices.append({'identity':r,'created':str(job/f"created_{r['pid']}.json"),'start':str(job/f"office_{r['pid']}_START.json"),'monitor_exact_identity_and_parent_match':True,'final_query':current})
    for p in job.glob('*EMPTY_EXCEL_WM_CLOSE_BEGIN*.json'):wm.append({'path':str(p),'sha256':sha(p),'record':read(p)})
    for p in job.glob('*FORCE_BEGIN*.json'):forces.append({'path':str(p),'sha256':sha(p)})
    jobs.append({'task':s['task'],'expected':s['expected'],'actual':result['status'],'result':str(result_path),'result_sha256':sha(result_path),'supervisor':str(sup_path),'supervisor_sha256':sha(sup_path),'job_request':str(job/'JOB_REQUEST.json'),'office_count':len(sup['instances']),'worker_returncode':sup['worker_returncode']})
assert len(jobs)==26 and len({key(x['identity']) for x in offices})==len(offices)
assert not forces
unknown=[r for r in monitor['instances'] if key(r) not in {key(x['identity']) for x in offices}]
assert not any(key(stop['unknown_office_exit'])==key(x['identity']) for x in offices)
assert stop['unknown_office_exit']['first_seen']>max(ends)
save(OUT/'FINAL_PROCESS_AUDIT.json',{'scope':'26 fresh frozen CLI tasks; runtime series stopped before independent GUI and further fault protocols','status':'PASS_OBSERVED_OWN_JOBS_WITH_EXTERNAL_STOP','jobs':jobs,'office':offices,'job_count':len(jobs),'office_count':len(offices),'normal_force':0,'all_force':0,'own_orphans':0,'own_unknown_identities':0,'validated_WM_CLOSE_events':wm,'unknown_non_review_office':unknown,'unknown_exit_after_all_jobs':True,'last_job_end_unix':max(ends),'foreign_exit_cause':'UNKNOWN','historical_Word_3132_cause':'UNKNOWN','no_global_taskkill':True,'no_unknown_process_termination':True,'monitor_access':'QUERY_LIMITED_INFORMATION|SYNCHRONIZE; all monitor handles released','full_foreign_sentinel_test':'NOT_RUN_SAFETY_STOP','full_worker_crash_probe':'NOT_RUN_SAFETY_STOP'})

# Path contracts, actual commands and sealed output references. This is not a system I/O trace.
path_options={'--workspace':(parent,),'--parent':(parent,),'--preview':(parent,OUT),'--manifest':(parent,),'--sandbox':(parent,),'--request':(parent,),'--result':(OUT,)}
request_paths=[]
for j in jobs:
    req=read(j['job_request']); args=req['arguments']; rows=[]
    assert req['supervisor_image']==str(ROOT/'stage9_dist_099/VZMK-Passport-0.9.9/VZMK-Passport-0.9.9.exe')
    for i,a in enumerate(args[:-1]):
        if a in path_options:
            p=Path(args[i+1]);assert any(p.resolve().is_relative_to(b.resolve()) for b in path_options[a]),(a,p)
            rows.append({'option':a,'path':str(p),'allowed_boundaries':[str(b) for b in path_options[a]]})
        elif a=='--source':assert Path(args[i+1])==Path(setup['source'])
        elif a=='--vom':assert Path(args[i+1])==Path(setup['vom'])
    request_paths.append({'request':j['job_request'],'paths':rows})
manifest_paths=[]
for p in e['manifests']:
    m=read(p)
    assert m['official_issue'] is False
    for r in m['records']:
        assert Path(r['path']).resolve().is_relative_to(parent.resolve()) and r['official_confirmed'] is False
    manifest_paths.append({'manifest':p,'record_paths':[r['path'] for r in m['records']]})
save(OUT/'PRODUCTION_ISOLATION_AUDIT.json',{'status':'PASS_REVIEW_SCOPE','actual_job_count':len(jobs),'requests':request_paths,'manifests':manifest_paths,'all_mutating_task_paths_confined_to_new_review_roots':True,'source_read_only':setup['source'],'source_before_after':str(OUT/'SOURCE_INTEGRITY_AFTER.json'),'working_registry_constant_SHA_required':False,'working_registry_open_write_restore_replace':False,'working_registry_historical_writer':'UNKNOWN','no_real_deductions':True,'no_official_closure':True,'no_deployment':True,'full_OS_file_IO_trace':'NOT_RUN; path contracts and immutable baseline, not continuous system-wide tracing','source_bindings':{'workspace.py':'bootstrap:28-83, bind:86-111; explicit DEV registry at37 and102','core.py':'settings:11-14, check:78-130, generate:154-233; TEST registry/run; official stop at273'}})

before=read(OUT/'PROTECTED_BEFORE.json'); expected=before['files']
# The working registry is outside ROOT and is never required to keep a stable SHA.
assert not any(Path(p).name=='ВЗМК.xlsx' and Path(p).parent==ROOT.parents[1] for p in expected)
def checked(item):
    p,h=item
    try:return p,sha(p),None
    except Exception as exc:return p,None,repr(exc)
print('Rechecking protected baseline',len(expected),flush=True)
with concurrent.futures.ThreadPoolExecutor(10) as pool: checked_files=list(pool.map(checked,expected.items()))
actual={p:h for p,h,err in checked_files if h is not None}
errors=[{'path':p,'error':err} for p,h,err in checked_files if err]
changed=[{'path':p,'before':h,'after':actual.get(p)} for p,h in expected.items() if h!=actual.get(p)]
current={str(p) for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts}
added=sorted(current-set(expected)); deleted=sorted(set(expected)-current)
save(OUT/'PROTECTED_AFTER.json',{'files':actual,'errors':errors,'production_registry_SHA_gate':False})
save(OUT/'FINAL_IMMUTABLE_INTEGRITY.json',{'status':'PASS' if not changed and not added and not deleted and not errors else 'FAIL','files_before':len(expected),'files_after':len(current),'changed':changed,'added':added,'deleted':deleted,'errors':errors,'own_evidence_excluded':str(OUT),'index_cache_excluded':'.codebase-index','production_registry_SHA_gate':False,'EXE_before':read(OUT/'BUILD_SHA256_BEFORE.json')['exe_sha256'],'EXE_after':read(OUT/'BUILD_SHA256_AFTER.json')['exe_sha256']})
assert not changed and not added and not deleted and not errors,('IMMUTABLE_DELTA',changed,added,deleted,errors)
save(OUT/'CONTROLLER_SHA256_FINAL.json',{'files':{str(p):sha(p) for p in OUT.glob('*.py')},'prepared_but_unexecuted':['negative.py','lifecycle.py','normal_overlap.py','overlap_crash_cleanup.py','foreign_pumped.py','ownership_stop.py','diagnostic_collision.py','repeat_native.py','gui_controller.py'],'exclusive_raw_result_retained':True,'inherited_raw_label_note':"e2e RESULT gui_smoke=NOT_RUN_USER_STOP and controller comments are historical copy labels; actual current GUI reason is NOT_RUN_SAFETY_STOP, independently recorded in REVIEW_RESULT and report. No raw evidence overwritten."})
print(json.dumps({'jobs':len(jobs),'Office':len(offices),'own_orphans':0,'all_force':0,'WM_CLOSE':len(wm),'unknown_non_review':len(unknown),'protected':len(expected),'integrity':'PASS'},indent=2),flush=True)
