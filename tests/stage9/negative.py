"""Reproducible Stage 9 boundaries on fresh copies only."""
import sys,copy,json,shutil,os
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app import core
from stage9_app.workspace import bind,save
from stage8_closure import storage as S,core as close,fixtures,recovery,model,native
from stage9_app.ownership import owned_office
OUT=ROOT/'validation/stage9_20261012_01'
work=Path(S.read(OUT/'WORKSPACE.json')['path']);work,_=bind(work)
results=[]
selected=set(sys.argv[1:])
attempt=uuid4().hex[:6]
def case(name,fn):
    if selected and name not in selected:return
    try:
        data=fn();results.append({'name':name,'status':'PASS','detail':data})
    except Exception as exc:
        import traceback
        results.append({'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()})
    print(name,results[-1]['status'],flush=True)
def stopped(fn):
    try:fn()
    except Exception as exc:return {'type':type(exc).__name__,'reason':str(exc)}
    raise AssertionError('EXPECTED STOP')
def seed():return Path(fixtures.create(work/'validation'/('stage8_s9_'+uuid4().hex[:8])/'sandboxes')['sandbox'])
def plan(r):return close.preview(r,r/'inputs/CLOSE_REQUEST.json')
def balances(r):return {k:v['before'] for k,v in model.registry_rows(r/'registry/УЗК_TEST.xlsm','STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ',{'s8-a','s8-b'}).items()}
def negative_request(name,change):
    r=seed();q=S.read(r/'inputs/CLOSE_REQUEST.json');change(q,r);(r/'inputs/CLOSE_REQUEST.json').write_text(json.dumps(q,ensure_ascii=False),encoding='utf8')
    before=S.sha(r/'registry/УЗК_TEST.xlsm');stop=stopped(lambda:plan(r));assert S.sha(r/'registry/УЗК_TEST.xlsm')==before
    return stop
with owned_office(OUT/f'NEGATIVE_OFFICE_{attempt}.json'):
    case('official_9_2026_block',lambda:stopped(lambda:core.official_close('9/2026')))
    case('missing_vom',lambda:stopped(lambda:core.check(work,'Гараж Казахстан','9/2026',work/'missing.xls')))
    bad=work/f'DEV/Stage6_Fixtures/bad_{attempt}.xlsx'
    bad.write_bytes(b'not Excel')
    case('invalid_vom',lambda:stopped(lambda:core.check(work,'Гараж Казахстан','9/2026',bad)))
    def changed_preview():
        d=S.read(OUT/'E2E.json')['data'];p={'data':d,'seal':S.digest(d)}
        p['data']=copy.deepcopy(d);p['data']['quantity']='1'
        return stopped(lambda:core.verify(p))
    case('changed_preview',changed_preview)
    def sha_conflict():
        source=work/'validation/sha_probe.txt';source.write_text('before');d={'root':str(work),'run':str(work/'validation/sha_case'),'official_issue':False,'fingerprints':{str(source):S.sha(source)}}
        p={'data':d,'seal':S.digest(d)};source.write_text('after');return stopped(lambda:core.verify(p))
    case('sha_conflict',sha_conflict)
    def incomplete():
        m=S.read(S.read(OUT/'E2E.json')['manifest']);m['records']=[r for r in m['records'] if r['type']!='VIK'];m.pop('seal');m['seal']=S.digest(m)
        path=work/f'validation/incomplete_{attempt}.json';save(path,m);return stopped(lambda:core.verify_manifest(path))
    case('incomplete_document_set',incomplete)
    case('closure_official_flag',lambda:negative_request('official',lambda q,r:q.update(official_issue=True)))
    case('closure_9_2026',lambda:negative_request('9',lambda q,r:q.update(passport='9/2026')))
    case('missing_required_document',lambda:negative_request('missing',lambda q,r:q['documents'].pop()))
    def unconfirmed(q,r):
        path=Path(q['control']['path']);data=S.read(path);data['all_required_control_confirmed']=False
        path.write_text(json.dumps(data,ensure_ascii=False),encoding='utf8');q['control']['sha256']=S.sha(path)
    case('unconfirmed_control',lambda:negative_request('control',unconfirmed))
    def once():
        r=seed();p=plan(r);before=balances(r);out=close.close(p,close.confirmation(p));after=balances(r)
        assert out['status']=='CLOSED_TEST' and before=={'s8-a':12,'s8-b':10} and after=={'s8-a':9,'s8-b':6}
        sha=S.sha(r/'registry/УЗК_TEST.xlsm');repeat=close.close(p,close.confirmation(p));assert repeat['status']=='ALREADY_CLOSED_TEST' and sha==S.sha(r/'registry/УЗК_TEST.xlsm')
        package=native.package(r/'registry/УЗК_TEST.xlsm');assert 'xl/vbaProject.bin' in package
        save(OUT/f'CLOSURE_FIXTURE_{attempt}.json',{'sandbox':str(r),'request':str(r/'inputs/CLOSE_REQUEST.json')})
        return {'before':before,'after':after,'repeat':repeat['status'],'preservation':out['operation'].get('preservation')}
    case('close_once_native_xlsm_vba',once)
    def failure(point):
        r=seed();p=plan(r);sha=S.sha(r/'registry/УЗК_TEST.xlsm')
        def fault(step,target):
            if step==point:raise OSError('STAGE9_TEST_INJECTED_'+point)
        result=stopped(lambda:close.close(p,close.confirmation(p),_fault=fault));assert S.sha(r/'registry/УЗК_TEST.xlsm')==sha
        recovered=close.recover(r);assert not S.history(r)['closed'];return {'stop':result,'recovery':recovered}
    case('excel_failure_before_save',lambda:failure('before_save'))
    case('save_io_failure',lambda:failure('save_io'))
    def corrupt_backup():
        r=seed();p=plan(r)
        def fault(step,target):
            if step=='after_save':raise OSError('STOP_AFTER_BACKUP')
        stopped(lambda:close.close(p,close.confirmation(p),_fault=fault))
        backup=next((r/'backups').rglob('*.xlsm'));backup.write_bytes(b'corrupt')
        sha=S.sha(r/'registry/УЗК_TEST.xlsm');result=stopped(lambda:close.recover(r));assert sha==S.sha(r/'registry/УЗК_TEST.xlsm');return result
    case('damaged_backup_stop',corrupt_backup)
    def torn_restore():
        r=seed();p=plan(r)
        def fault(step,target):
            if step=='during_publish':raise OSError('TEST_PROCESS_INTERRUPTION')
        stopped(lambda:close.close(p,close.confirmation(p),_fault=fault));op=S.history(r)['active'];assert op['state']=='APPLYING'
        assert S.sha(r/'registry/УЗК_TEST.xlsm') not in (op['before_sha256'],op['after_sha256'])
        rp=recovery.preview(r);result=recovery.restore(rp,recovery.confirmation(rp));assert result['status']=='RECOVERED_ABORTED' and balances(r)=={'s8-a':12,'s8-b':10}
        assert not S.history(r)['closed'];return {'result':result,'new_preview':plan(r)['payload']['status']}
    case('torn_xlsm_restore_no_decrement',torn_restore)
save(OUT/('NEGATIVE_'+uuid4().hex[:6]+'.json'),{'pass':sum(r['status']=='PASS' for r in results),'fail':sum(r['status']=='FAIL' for r in results),'cases':results})
sys.exit(any(r['status']=='FAIL' for r in results))
