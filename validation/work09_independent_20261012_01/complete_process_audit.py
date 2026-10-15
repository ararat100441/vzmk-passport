import json,hashlib
from pathlib import Path
import win32api,win32process,win32event
OUT=Path(__file__).resolve().parent
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
source=OUT/'FINAL_PROCESS_AND_SOURCE_AUDIT.json';d=read(source);reports={Path(r['path']).parent.name:r for r in d['reports']};records={}
for p in [*OUT.glob('lifecycle_*/office-probe_*.json'),*OUT.glob('probe_*/RESULT.json'),*OUT.glob('cold_*/RESULT.json'),OUT/'serial_normal/RESULT.json']:
    r=read(p);s=read(r['office_supervisor']);reports[Path(r['office_supervisor']).parent.name]={'path':r['office_supervisor'],'orphans':s['owned_orphans'],'ownership_errors':s['ownership_errors'],'worker_returncode':s['worker_returncode']}
for report in reports.values():
    for r in read(report['path'])['instances']:records[(r['pid'],r['creation_time'])]=r
states=[]
for key,r in records.items():
    h=None;alive=False
    try:
        h=win32api.OpenProcess(0x100400,False,r['pid']);alive=win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time'] and win32event.WaitForSingleObject(h,0)==258
    except win32api.error as exc:
        if exc.winerror not in (87,1168):raise
    finally:
        if h is not None:h.Close()
    states.append({'pid':r['pid'],'creation_time':r['creation_time'],'same_identity_alive':alive})
assert not any(r['same_identity_alive'] for r in states)
d.update(office_records_checked=len(records),reports=list(reports.values()),record_states=states,previous_subset_audit={'path':str(source),'sha256':hashlib.sha256(source.read_bytes()).hexdigest()},all_new_review_jobs_including_existing_profile_comparison=True)
with (OUT/'FINAL_PROCESS_AND_SOURCE_AUDIT_COMPLETE.json').open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2)
print('COMPLETE',len(records),'records',len(reports),'jobs, owned orphan 0',flush=True)
