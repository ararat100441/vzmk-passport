from common import *
import traceback,copy,json
from stage8_closure.storage import digest
from stage8_closure.core import confirmation
from stage8_closure.recovery import confirmation as restore_confirmation
from openpyxl import Workbook
RUN=OUT/('negative_'+uuid4().hex[:8]);RUN.mkdir();cases=[]
e2e=read(OUT/'EXE_E2E_FINAL.json');work=e2e['workspace'];manifest=e2e['manifests'][0]
vom=ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
def call(name,expected='PASS',**args):return task(RUN,name,expected,workspace=work,**args)[0]
def case(name,fn):
    try:fn();cases.append({'name':name,'status':'PASS'});print(name,'PASS',flush=True)
    except Exception as exc:cases.append({'name':name,'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc()});raise
def save_plan(plan,label):
    p=RUN/(label+'_'+uuid4().hex[:6]+'.json');save(p,plan);return p
def seed():return call('fixture')
try:
    case('missing_vom',lambda:call('check','FAIL',vom=RUN/'missing.xlsx'))
    invalid=RUN/'invalid.xlsx';invalid.write_bytes(b'NOT AN EXCEL WORKBOOK')
    case('invalid_vom',lambda:call('check','FAIL',vom=invalid))
    def preview_tamper():
        p=call('check',vom=vom)['preview'];data=read(p);data['seal']='changed';target=Path(work)/'validation'/('tampered_'+uuid4().hex[:8]+'.json');save(target,data);call('generate','FAIL',preview=target,confirm_test=True)
    case('changed_preview',preview_tamper)
    def sha_conflict():
        p=call('check',vom=vom)['preview'];data=read(p);f=Path(data['run'])/'SOURCE_VOM.xls';f.write_bytes(b'changed owned snapshot')
        call('generate','FAIL',preview=p,confirm_test=True)
    case('sha_conflict',sha_conflict)
    def incomplete():
        m=read(manifest);m['records']=[r for r in m['records'] if r['type']!='VIK'];m.pop('seal');m['seal']=digest(m)
        target=Path(work)/'validation'/('incomplete_'+uuid4().hex[:8]+'.json');save(target,m);call('verify','FAIL',manifest=target)
    case('incomplete_document_set',incomplete)
    fixture=seed();original=read(fixture['request']);registry=Path(fixture['sandbox'])/'registry/УЗК_TEST.xlsm';before=sha(registry)
    def bad_request(label,mutate):
        q=copy.deepcopy(original);mutate(q);p=Path(fixture['request']).with_name(label+'.json');save(p,q)
        call('close-preview','FAIL',sandbox=fixture['sandbox'],request=p);assert sha(registry)==before
    case('official_flag',lambda:bad_request('official',lambda q:q.update(official_issue=True)))
    case('official_9_2026',lambda:bad_request('passport9',lambda q:q.update(passport='9/2026')))
    case('missing_required_uzk',lambda:bad_request('missingUZK',lambda q:q.update(documents=[d for d in q['documents'] if d['kind']!='UZK'])))
    def unconfirmed(q):
        control=read(q['control']['path']);control['all_required_control_confirmed']=False;p=Path(fixture['request']).with_name('UNCONFIRMED_CONTROL.json');save(p,control);q['control']={'path':str(p),'sha256':sha(p)}
    case('unconfirmed_control',lambda:bad_request('unconfirmed',unconfirmed))
    def locked():
        import win32file,win32con
        handle=win32file.CreateFile(str(registry),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
        try:call('close-preview','FAIL',sandbox=fixture['sandbox'],request=fixture['request'])
        finally:handle.Close()
        assert sha(registry)==before
    case('windows_registry_lock',locked)
    for point in ('before_save','save_io'):
        def fault_case(point=point):
            f=seed();reg=Path(f['sandbox'])/'registry/УЗК_TEST.xlsm';h=sha(reg)
            plan=call('close-preview',sandbox=f['sandbox'],request=f['request'])['plan'];p=save_plan(plan,point)
            result=call('close','FAIL',preview=p,confirm=confirmation(plan),test_fault=point)
            assert 'STAGE9_TEST_INJECTED_' in result['error'] and sha(reg)==h
            call('recover',sandbox=f['sandbox']);assert sha(reg)==h
        case('excel_'+point,fault_case)
    def damaged_backup():
        f=seed();plan=call('close-preview',sandbox=f['sandbox'],request=f['request'])['plan'];p=save_plan(plan,'after_save')
        call('close','FAIL',preview=p,confirm=confirmation(plan),test_fault='after_save')
        backup=next((Path(f['sandbox'])/'backups').rglob('*.xlsm'));backup.write_bytes(b'DAMAGED ISOLATED TEST BACKUP')
        registry=Path(f['sandbox'])/'registry/УЗК_TEST.xlsm';before=sha(registry)
        call('recover','FAIL',sandbox=f['sandbox']);assert sha(registry)==before
    case('damaged_backup_stop',damaged_backup)
    save(RUN/'RESULT.json',{'status':'PASS','exe_sha256':sha(EXE),'pass':len(cases),'fail':0,'cases':cases})
    save(OUT/'EXE_NEGATIVE_FINAL.json',{'status':'PASS','exe_sha256':sha(EXE),'evidence':str(RUN/'RESULT.json'),'pass':len(cases),'fail':0})
except Exception as exc:
    save(RUN/'FAILURE.json',{'status':'FAIL','exe_sha256':sha(EXE),'error':repr(exc),'traceback':traceback.format_exc(),'cases':cases});raise
