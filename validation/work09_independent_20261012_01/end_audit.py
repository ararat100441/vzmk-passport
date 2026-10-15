import os,sys,json,hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent;sys.path.insert(0,str(ROOT))
from stage9_app.ownership import process_image
from stage8_closure.native import vba_streams
import win32api,win32process,win32event,win32security
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def save(n,d):
    with (OUT/n).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2)
data_root=Path(read(OUT/'RUNTIME_ENV.json')['LOCALAPPDATA']);work=Path(read(next(OUT.glob('lifecycle_*/import_*.json')))['workspace']);roots={data_root,work.parents[2]}
records={};reports=[]
for base in roots:
    for p in base.rglob('SUPERVISOR.json'):
        d=read(p);reports.append({'path':str(p),'orphans':d['owned_orphans'],'ownership_errors':d['ownership_errors'],'worker_returncode':d['worker_returncode']})
        for r in d['instances']:records[(r['pid'],r['creation_time'])]=r|{'report':str(p)}
alive=[]
for key,r in records.items():
    h=None
    try:
        h=win32api.OpenProcess(0x100400,False,r['pid'])
        if win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time'] and win32event.WaitForSingleObject(h,0)==258:alive.append(r)
    except win32api.error as exc:
        if exc.winerror not in (87,1168):raise
    finally:
        if h is not None:h.Close()
users=read(next(OUT.glob('lifecycle_*/USER_OFFICE_BEFORE.json')));user_states=[]
for pid,r in users.items():
    h=win32api.OpenProcess(0x100400,False,int(pid))
    try:user_states.append({'pid':int(pid),'creation_time':r['creation_time'],'same_identity':win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time'] and process_image(h)==r['image'],'alive':win32event.WaitForSingleObject(h,0)==258})
    finally:h.Close()
token=win32security.OpenProcessToken(win32api.GetCurrentProcess(),win32security.TOKEN_QUERY)
try:elevated=win32security.GetTokenInformation(token,20)
finally:token.Close()
cases=[]
for p in [*OUT.glob('lifecycle_*/office-probe_*.json'),*OUT.glob('probe_*/RESULT.json'),*OUT.glob('cold_*/RESULT.json'),OUT/'serial_normal/RESULT.json']:
    d=read(p);s=read(d['office_supervisor']);cases.append({'path':str(p),'status':d['status'],'error':d.get('error'),'worker_returncode':s['worker_returncode'],'forced_pids':[r['pid'] for r in s['instances'] if r.get('forced_termination_attempted')],'orphan':s['owned_orphans']})
negative=read(OUT/'safe_negative/RESULT.json');sandbox=Path(negative['sandbox']);streams=vba_streams(sandbox/'registry/УЗК_TEST.xlsm');seed=read(sandbox/'SANDBOX.json')['native_seed']
for name in ('BUILD_SHA256','SOURCE_SHA256'):
    assert read(OUT/(name+'_BEFORE.json'))==read(OUT/(name+'_AFTER.json')),name
assert not alive and all(s['same_identity'] and s['alive'] for s in user_states)
assert streams==seed['vba_streams']
save('FINAL_PROCESS_AND_SOURCE_AUDIT.json',{'status':'PASS','token_elevation':elevated,'office_records_checked':len(records),'owned_orphans':alive,'user_office':user_states,'reports':reports,'probes':cases,'probe_pass':sum(c['status']=='PASS' for c in cases),'probe_fail':sum(c['status']=='FAIL' for c in cases),'fixture_vba_streams_equal_to_seed':True,'closure_vba_preservation':'NOT_RUN','fixture_vba_stream_count':len(streams),'fixture_vba_stream_sha256':streams,'fixture_registry_sha256':negative['registry_sha256'],'fixture_native_seed':seed,'fixture_no_committed_close':not list((sandbox/'journal').glob('*.json')),'source_and_build_before_after_equal':True,'broad_source_integrity':read(OUT/'INDEPENDENT_AFTER.json'),'data_roots':[str(p) for p in roots]})
print('FINAL',len(records),'owned records; orphan',len(alive),'probes',len(cases),'PASS',sum(c['status']=='PASS' for c in cases),'FAIL',sum(c['status']=='FAIL' for c in cases),'VBA streams',len(streams),flush=True)
