import os,sys,json,shutil,hashlib,copy
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
os.environ['STAGE9_VALIDATION_OUT']=str(OUT)
os.environ['LOCALAPPDATA']=json.loads((OUT/'RUNTIME_ENV.json').read_text())['LOCALAPPDATA']
sys.path.insert(0,str(ROOT/'tests/stage9_091'))
from common import task,save,read,sha
RUN=OUT/'safe_negative';RUN.mkdir();cases=[]
work=Path(read(next(OUT.glob('lifecycle_*/import_*.json')))['workspace'])
def call(label,name,contains=None,workspace=work,**kw):
    folder=RUN/label;folder.mkdir();d,s,p=task(folder,name,'FAIL',workspace=workspace,**kw)
    if contains:assert contains in d.get('error',''),d
    cases.append({'case':label,'status':'PASS','application_status':d['status'],'error':d.get('error'),'result':str(p),'office_instances':len(s['instances'])})
call('missing_vom','check',vom=RUN/'missing.xlsx',contains='INPUT_LOCKED')
invalid=RUN/'invalid.xlsx';invalid.write_bytes(b'INDEPENDENT INVALID WORKBOOK')
call('invalid_vom','check',vom=invalid)
for label,mutation in [('missing_template','missing'),('changed_template','changed')]:
    target=RUN/label/'copy'/work.name;shutil.copytree(work,target)
    template=target/'documents/project/Паспорт/1_2026 - Документ о качестве.docx'
    if mutation=='missing':template.rename(template.with_suffix('.removed_for_test'))
    else:template.write_bytes(b'INDEPENDENT SHA TAMPER')
    call(label+'_check','check',workspace=target,vom=invalid,contains='TEMPLATE_OR_SAMPLE_SHA_CONFLICT' if mutation=='changed' else None)
requests=list((work/'validation').glob('stage8_s9_*/sandboxes/*/inputs/CLOSE_REQUEST.json'))
assert len(requests)==1,requests
request=requests[0];sandbox=request.parents[1];original=read(request);registry=sandbox/'registry/УЗК_TEST.xlsm';before=sha(registry)
def modified(label,change,contains=None):
    q=copy.deepcopy(original);change(q);p=request.with_name(label+'.json');save(p,q)
    call(label,'close-preview',sandbox=sandbox,request=p,contains=contains)
    assert sha(registry)==before and not list((sandbox/'journal').glob('*.json'))
modified('official_flag',lambda q:q.update(official_issue=True),'OFFICIAL_CLOSURE_FORBIDDEN')
modified('passport_9_2026',lambda q:q.update(passport='9/2026'),'PASSPORT_9_2026_DEV_TEST_FORBIDDEN')
modified('missing_required_uzk',lambda q:q.update(documents=[d for d in q['documents'] if d['kind']!='UZK']))
modified('incomplete_akz_pair',lambda q:q.update(documents=[d for d in q['documents'] if d['kind']!='AKZ_APPENDIX']))
def unconfirmed(q):
    d=read(q['control']['path']);d['all_required_control_confirmed']=False;p=request.with_name('UNCONFIRMED_CONTROL.json');save(p,d);q['control']={'path':str(p),'sha256':sha(p)}
modified('unconfirmed_control',unconfirmed)
import win32file,win32con
h=win32file.CreateFile(str(registry),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
try:call('windows_exclusive_lock','close-preview',sandbox=sandbox,request=request)
finally:h.Close()
assert sha(registry)==before and not list((sandbox/'journal').glob('*.json'))
save(RUN/'RESULT.json',{'status':'PASS','cases':cases,'count':len(cases),'registry_unchanged':True,'journal_empty':True,'no_close_or_deduction_performed':True,'sandbox':str(sandbox)})
print('SAFE NEGATIVE',len(cases),'PASS',flush=True)
