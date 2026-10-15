import os,sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent;RUN=OUT/'safe_negative'
os.environ['STAGE9_VALIDATION_OUT']=str(OUT);os.environ['LOCALAPPDATA']=json.loads((OUT/'RUNTIME_ENV.json').read_text())['LOCALAPPDATA']
sys.path.insert(0,str(ROOT/'tests/stage9_091'))
from common import task,save,read,sha
work=Path(read(next(OUT.glob('lifecycle_*/import_*.json')))['workspace']);request=next((work/'validation').glob('stage8_s9_*/sandboxes/*/inputs/CLOSE_REQUEST.json'));sandbox=request.parents[1]
reg=sandbox/'registry/УЗК_TEST.xlsm';before=sha(reg);cases=[]
for p in sorted(RUN.glob('*/*.json')):
    d=read(p)
    if d.get('status')=='FAIL' and 'office_supervisor' in d:
        s=read(d['office_supervisor']);assert s['owned_orphans']==0
        cases.append({'case':p.parent.name,'status':'PASS','application_status':'FAIL','error':d.get('error'),'result':str(p),'office_instances':len(s['instances'])})
assert len(cases)==8,cases
def call(label,contains=None,**kw):
    folder=RUN/label;folder.mkdir();d,s,p=task(folder,'close-preview','FAIL',workspace=work,sandbox=sandbox,**kw)
    if contains:assert contains in d['error']
    cases.append({'case':label,'status':'PASS','application_status':'FAIL','error':d.get('error'),'result':str(p),'office_instances':len(s['instances'])})
q=read(request);control=request.with_name('UNCONFIRMED_CONTROL.json');assert read(control)['all_required_control_confirmed'] is False
q['control']={'path':str(control),'sha256':sha(control)};p=request.with_name('UNCONFIRMED_REQUEST.json');save(p,q)
call('unconfirmed_control',contains='CONTROL_NOT_CONFIRMED',request=p)
import win32file,win32con
h=win32file.CreateFile(str(reg),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
try:call('windows_exclusive_lock',request=request)
finally:h.Close()
assert sha(reg)==before and not list((sandbox/'journal').glob('*.json'))
save(RUN/'RESULT.json',{'status':'PASS','count':len(cases),'cases':cases,'registry_unchanged':True,'registry_sha256':before,'journal_empty':True,'no_close_or_deduction_performed':True,'sandbox':str(sandbox),'harness_failure_retained':'SAFE_NEGATIVE_RUN.log','resumed_without_repeating_passed_cases':True})
print('SAFE NEGATIVE',len(cases),'PASS',flush=True)
