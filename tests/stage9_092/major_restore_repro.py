"""Repeat the concrete normal cleanup failure on its own torn TEST fixture."""
from common import *
import time,json
fixed=len(sys.argv)>1 and sys.argv[1]=='fixed'
RUN=OUT/(('restore_fixed_' if fixed else 'restore_repro_')+uuid4().hex[:8]);RUN.mkdir();candidates=[]
for folder in OUT.glob('e2e_*'):
    for path in folder.glob('restore-preview_*.json'):
        d=read(path)
        if d.get('error')=='OFFICE_GRACEFUL_EXIT_TIMEOUT':candidates.append((folder,path,d))
folder,failed,data=candidates[0];work=read(next(folder.glob('import_*.json')))['workspace'];sandbox=data['worker_result']['plan']['payload']['sandbox']
exe=EXE if fixed else ROOT/'stage9_dist_093/VZMK-Passport-0.9.3/VZMK-Passport-0.9.3.exe';cases=[]
for n in range(3):
    path=RUN/f'case_{n}.json';started=time.time()
    p=subprocess.Popen([str(exe),'--task','restore-preview','--workspace',work,'--sandbox',sandbox,'--result',str(path)],env=ENV,creationflags=subprocess.CREATE_NO_WINDOW)
    p.wait(timeout=180);d=read(path);sup=read(d['office_supervisor']);assert not sup['ownership_errors'] and sup['owned_orphans']==0
    if fixed:assert d['status']=='PASS' and not any(r.get('forced_termination_attempted') for r in sup['instances']),d
    cases.append({'result':str(path),'supervisor':d['office_supervisor'],'returncode':p.returncode,'seconds':time.time()-started,'status':d['status']});print('restore baseline',n,d['status'],d.get('error'),flush=True)
save(RUN/'RESULT.json',{'status':'CAPTURED','exe_sha256':sha(exe),'original_failure':str(failed),'workspace':work,'sandbox':sandbox,'cases':cases})
save(OUT/('MAJOR_RESTORE_FIXED_097.json' if fixed else 'MAJOR_RESTORE_BASELINE.json'),{'evidence':str(RUN/'RESULT.json'),'cases':cases,'original_failure':str(failed)})
