"""Independent review harness. Product is imported unchanged; all writes are local fixtures."""
import copy, hashlib, json, os, shutil, subprocess, sys, time, traceback
from contextlib import contextmanager
from pathlib import Path
from uuid import uuid4
from unittest.mock import patch

RUN=Path(__file__).resolve().parent; ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT)); sys.dont_write_bytecode=True
from stage8_closure import core, native, storage as S
from stage8_closure.fixtures import create
from stage8_closure.office import excel_session
from stage8_closure.model import registry_rows
ENV=os.environ|{'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1','PYTHONPATH':str(ROOT)}
def save(p,x): Path(p).write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf8')
def reg(root): return root/'registry/УЗК_TEST.xlsm'
def balances(root): return {k:v['before'] for k,v in registry_rows(reg(root),'STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ',{'s8-a','s8-b'}).items()}
def reseed(root):
    m=S.read(root/'SANDBOX.json');m.pop('seal');m['initial_registry_sha256']=S.sha(reg(root));m['seal']=S.digest(m);save(root/'SANDBOX.json',m)
def expect(fn, contains=None):
    try: fn()
    except Exception as exc:
        if contains: assert contains in str(exc),(contains,repr(exc))
        return {'type':type(exc).__name__,'message':str(exc)}
    raise AssertionError('Expected refusal')
def amend(root,name,fn):
    p=root/'inputs'/name;x=S.read(p);fn(x);save(p,x)
    q=S.read(root/'inputs/CLOSE_REQUEST.json')
    for entry in [q['shipment'],q['requirements'],q['control'],*q['documents']]:
        if Path(entry['path'])==p: entry['sha256']=S.sha(p)
    save(root/'inputs/CLOSE_REQUEST.json',q)
def request(root): return root/'inputs/CLOSE_REQUEST.json'
def plan(root): return core.preview(root,request(root))
def close(root,fault=None):
    p=plan(root);return core.close(p,core.confirmation(p),_fault=fault)
def states(root): return [e['state'] for e in S.history(root)['events']]
def fresh(root,command='recover',point=None):
    p=root/'WORKER_PREVIEW.json'
    if command!='recover':save(p,plan(root))
    args=[sys.executable,'-B','-X','faulthandler',str(__file__),'worker',command,str(root if command=='recover' else p)]
    if point: args.append(point)
    r=subprocess.run(args,env=ENV,capture_output=True,text=True,encoding='utf8',timeout=60)
    save(root/(command+'_'+str(point)+'_PROCESS.json'),{'args':args,'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr})
    return r

if len(sys.argv)>1 and sys.argv[1]=='worker':
    command,path=sys.argv[2:4]
    try:
        if command=='recover': result=core.recover(path)
        else:
            p=S.read(path)
            if command=='parallel':
                barrier=Path(sys.argv[4]);start=time.monotonic()
                while not barrier.exists():
                    if time.monotonic()-start>20: raise TimeoutError('BARRIER')
                    time.sleep(.02)
            def fault(point,target):
                if command=='crash' and point==sys.argv[4]:os._exit(73)
            result=core.close(p,core.confirmation(p),_fault=fault)
        print(json.dumps(result,ensure_ascii=False));sys.exit(0)
    except Exception as exc:print(json.dumps({'status':'STOP','error':str(exc)},ensure_ascii=False));sys.exit(2)

SEED=None;records=[]
def fixture():
    root=RUN/'sandboxes'/uuid4().hex;root.mkdir()
    for name in ('registry','journal','backups','transactions','inputs'):(root/name).mkdir()
    shutil.copyfile(reg(SEED),reg(root))
    m=S.read(SEED/'SANDBOX.json');m.pop('seal');m['sandbox_id']=root.name;m['seal']=S.digest(m);save(root/'SANDBOX.json',m)
    for p in (SEED/'inputs').glob('*.json'):
        x=S.read(p)
        if p.name=='CLOSE_REQUEST.json':
            for entry in [x['shipment'],x['requirements'],x['control'],*x['documents']]:entry['path']=str(root/'inputs'/Path(entry['path']).name)
        save(root/'inputs'/p.name,x)
    q=S.read(request(root))
    for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
    save(request(root),q)
    return root
def case(name,fn):
    print(name,'START',flush=True);t=time.monotonic()
    try: item={'name':name,'status':'PASS','detail':fn()}
    except Exception as exc:item={'name':name,'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc()}
    item['seconds']=round(time.monotonic()-t,3);records.append(item);save(RUN/'cases'/(name+'.json'),item)
    print(name,item['status'],item.get('error',''),flush=True)
    save(RUN/'INDEPENDENT_RESULTS.json',{'passed':sum(x['status']=='PASS' for x in records),'failed':sum(x['status']=='FAIL' for x in records),'cases':records})

def run():
    global SEED
    (RUN/'cases').mkdir(exist_ok=True)
    f=create(RUN/'sandboxes');SEED=Path(f['sandbox']);save(RUN/'SEED.json',f)
    def positive():
        root=fixture();before=S.sha(reg(root));vba=native.vba_streams(reg(root));p=plan(root);result=close(root)
        assert result['status']=='CLOSED_TEST' and result['official_issue'] is False
        assert balances(root)=={'s8-a':9,'s8-b':6};op=result['operation']
        assert states(root)==['INTENT','BACKUP_VERIFIED','STAGED','APPLYING','COMMITTED']
        assert S.sha(op['backup']['path'])==before
        # Independently enumerate actual stream bytes and compare package/sheets/formulas.
        import io,olefile
        from zipfile import ZipFile
        def streams(p):
            with ZipFile(p) as z,olefile.OleFileIO(io.BytesIO(z.read('xl/vbaProject.bin'))) as ole:
                return {'/'.join(n):hashlib.sha256(ole.openstream(n).read()).hexdigest() for n in ole.listdir()}
        assert streams(reg(root))==streams(op['backup']['path'])==vba and len(vba)==11
        observed=native.inspect(reg(root),{'УЗК!C24':9,'УЗК!C25':6})
        backup_native=native.inspect(op['backup']['path'],{'УЗК!C24':12,'УЗК!C25':10})
        assert observed['sheets']==backup_native['sheets']
        assert observed['formulas']['S8_FORMULAS!$A$1']['value']==23
        assert observed['formulas']['S8_FORMULAS!$A$2']['value']==5
        for k in observed['formulas']: assert observed['formulas'][k]['formula']==backup_native['formulas'][k]['formula']
        ignored={(native.sheet_part(reg(root)),'C24'),(native.sheet_part(reg(root)),'C25')}
        assert native.package(reg(root),ignored)==native.package(op['backup']['path'],ignored)
        after=S.sha(reg(root));ev=states(root)
        assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
        q=copy.deepcopy(S.read(request(root)));q['shipment_id']='OTHER-SHIPMENT'
        alias=root/'inputs/RENAMED_REQUEST.json';save(alias,q)
        assert core.preview(root,alias)['payload']['status']=='ALREADY_CLOSED_TEST'
        r=fresh(root,'close');assert r.returncode==0 and json.loads(r.stdout)['status']=='ALREADY_CLOSED_TEST'
        assert S.sha(reg(root))==after and states(root)==ev
        save(RUN/'POSITIVE_NATIVE.json',{'sandbox':str(root),'before':backup_native,'after':observed,'vba_streams':vba,'backup':op['backup'],'parts_preserved':True})
        save(RUN/'GUI_FIXTURE_CLOSED.json',{'sandbox':str(root),'request':str(request(root))})
        return {'sandbox':str(root),'balances':balances(root),'vba_streams':len(vba),'repeat_path_and_process':'ALREADY_CLOSED_TEST'}
    case('positive_exact_debit_preservation_backup_and_repeats',positive)
    def reject(name,mutation,contains):
        def test():
            root=fixture();mutation(root);before=S.sha(reg(root));e=expect(lambda:plan(root),contains)
            assert S.sha(reg(root))==before and not states(root)
            return {'sandbox':str(root),'refusal':e,'registry_unchanged':True}
        case(name,test)
    reject('changed_input_sha',lambda r:(r/'inputs/CONTROL.json').write_bytes((r/'inputs/CONTROL.json').read_bytes()+b' '),'INPUT_SHA256_CHANGED')
    reject('mandatory_vik_missing',lambda r:amend(r,'CLOSE_REQUEST.json',lambda q:q.update(documents=[x for x in q['documents'] if x['kind']!='VIK'])),'MANDATORY_DOCUMENTS_MISSING')
    reject('unconfirmed_control',lambda r:amend(r,'CONTROL.json',lambda q:q.update(all_required_control_confirmed=False)),'CONTROL_NOT_CONFIRMED')
    reject('partially_confirmed_volume',lambda r:amend(r,'CONTROL.json',lambda q:q['uzk'][0].update(quantity=2)),'UZK_CONTROL_INCOMPLETE')
    reject('sample_only_document',lambda r:amend(r,'UZK.json',lambda q:q.update(status='SAMPLE_ONLY')),'DOCUMENT_NOT_CONFIRMED')
    reject('test_document_not_official',lambda r:amend(r,'VIK.json',lambda q:q.update(status='TEST')),'DOCUMENT_NOT_CONFIRMED')
    reject('other_shipment_document',lambda r:amend(r,'VIK.json',lambda q:q.update(shipment_id='OTHER')),'EVIDENCE_MEMBERSHIP_MISMATCH')
    reject('official_promotion_forbidden',lambda r:amend(r,'CLOSE_REQUEST.json',lambda q:q.update(mode='OFFICIAL',official_issue=True)),'OFFICIAL_CLOSURE_FORBIDDEN')
    reject('real_passport_9_2026_blocked',lambda r:amend(r,'CLOSE_REQUEST.json',lambda q:q.update(passport='9/2026')),'PASSPORT_9_2026_DEV_TEST_FORBIDDEN')
    def insufficient(r):
        amend(r,'SHIPMENT.json',lambda q:q['rows'][0].update(quantity=20))
        amend(r,'REQUIREMENTS.json',lambda q:q['uzk'][0].update(quantity=13))
        amend(r,'CONTROL.json',lambda q:q['uzk'][0].update(quantity=13))
    reject('insufficient_balance',insufficient,'INSUFFICIENT_UZK_BALANCE')
    def stale():
        root=fixture();p=plan(root);before=S.sha(reg(root));amend(root,'CONTROL.json',lambda q:q.update(note='changed'))
        e=expect(lambda:core.close(p,core.confirmation(p)),'PREVIEW_STALE');assert S.sha(reg(root))==before and not states(root)
        return {'sandbox':str(root),'refusal':e}
    case('inputs_changed_after_preview',stale)
    def tamper():
        root=fixture();p=plan(root);p['payload']['deductions'][0]['after']=999
        a=expect(lambda:core.close(p,core.confirmation(p)),'PREVIEW_TAMPERED')
        p['seal']=S.digest(p['payload']);b=expect(lambda:core.close(p,core.confirmation(p)),'PREVIEW_STALE')
        assert not states(root);return {'sandbox':str(root),'unsealed':a,'resealed':b}
    case('preview_tamper_and_recomputed_seal',tamper)
    def changed_registry():
        root=fixture();p=plan(root);reg(root).write_bytes(reg(root).read_bytes()+b' ')
        e=expect(lambda:core.close(p,core.confirmation(p)),'REGISTRY_CHANGED_OR_JOURNAL_LOST');assert not states(root)
        return {'sandbox':str(root),'refusal':e}
    case('registry_changed_after_preview',changed_registry)
    def confirmation():
        root=fixture();p=plan(root);e=expect(lambda:core.close(p,''),'EXPLICIT_CONFIRMATION_REQUIRED');assert not states(root)
        return {'sandbox':str(root),'refusal':e}
    case('explicit_confirmation_required',confirmation)
    def path_escape():
        root=fixture();p=plan(root)
        e=expect(lambda:core.preview(root,ROOT/'stage8.py'),'INPUT_OUTSIDE_SANDBOX')
        alias=root/'inputs/HARDLINK.json';os.link(request(root),alias)
        h=expect(lambda:core.preview(root,alias),'HARDLINK_FORBIDDEN');assert not states(root)
        return {'sandbox':str(root),'external':e,'hardlink':h}
    case('invalid_external_path_and_hardlink',path_escape)
    def no_uzk():
        root=fixture();amend(root,'REQUIREMENTS.json',lambda q:q.update(uzk_required=False,uzk=[]));amend(root,'CONTROL.json',lambda q:q.update(uzk=[]))
        amend(root,'CLOSE_REQUEST.json',lambda q:q.update(documents=[e for e in q['documents'] if e['kind']!='UZK']))
        before=S.sha(reg(root));r=close(root);assert r['status']=='CLOSED_TEST' and S.sha(reg(root))==before
        assert not list((root/'backups').iterdir()) and not list((root/'transactions').rglob('*.xlsm'))
        return {'sandbox':str(root),'registry_unchanged':True,'backup_save_replace':0}
    case('no_uzk_required_zero_registry_writes',no_uzk)
    def backup_error():
        root=fixture();before=S.sha(reg(root))
        def fault(point,target):
            if point=='after_intent': (root/'backups'/S.history(root)['active']['operation_id']).write_bytes(b'not a directory')
        e=expect(lambda:close(root,fault));assert S.sha(reg(root))==before and states(root)==['INTENT']
        assert core.recover(root)['status']=='ABORTED_UNCHANGED'
        return {'sandbox':str(root),'real_backup_io_failure':e,'registry_unchanged':True}
    case('backup_filesystem_error_blocks_save',backup_error)
    def corrupt_backup():
        root=fixture();before=S.sha(reg(root))
        def fault(point,target):
            if point=='after_backup':target.write_bytes(b'corrupted backup')
        e=expect(lambda:close(root,fault),'BACKUP_DAMAGED');assert S.sha(reg(root))==before and states(root)==['INTENT']
        # Backup was not verified or journal-bound; unchanged source permits abort, never replay.
        assert core.recover(root)['status']=='ABORTED_UNCHANGED'
        return {'sandbox':str(root),'refusal':e,'registry_unchanged':True}
    case('corrupt_unverified_backup_prevents_save',corrupt_backup)
    for point in ('before_save','save_io','during_save','after_save','before_replace'):
        def failure(point=point):
            root=fixture();before=S.sha(reg(root))
            def fault(p,target):
                if p==point:
                    if p=='during_save':target.write_bytes(b'torn save')
                    else:raise OSError('INDEPENDENT_IO_FAILURE_'+point)
            e=expect(lambda:close(root,fault));assert S.sha(reg(root))==before and S.history(root)['active'] is not None
            stop=expect(lambda:plan(root),'INCOMPLETE_OPERATION_RECOVER_FIRST')
            r=fresh(root);assert r.returncode==0 and json.loads(r.stdout)['status']=='ABORTED_UNCHANGED'
            assert S.sha(reg(root))==before
            return {'sandbox':str(root),'refusal':e,'new_close_refused':stop,'recovered':'ABORTED_UNCHANGED','registry_unchanged':True}
        case('failure_'+point+'_and_fresh_recovery',failure)
    for point in ('before_save','after_replace'):
        def death(point=point):
            root=fixture();before=S.sha(reg(root));r=fresh(root,'crash',point);assert r.returncode==73
            current=S.sha(reg(root));assert 'COMMITTED' not in states(root)
            expect(lambda:plan(root),'INCOMPLETE_OPERATION_RECOVER_FIRST')
            recovery=fresh(root);assert recovery.returncode==0
            result=json.loads(recovery.stdout)
            if point=='before_save':assert current==before and result['status']=='ABORTED_UNCHANGED'
            else:
                assert current!=before and result['status']=='CLOSED_TEST' and result['recovered'] is True
                assert balances(root)=={'s8-a':9,'s8-b':6} and plan(root)['payload']['status']=='ALREADY_CLOSED_TEST'
            assert S.sha(reg(root))==current
            return {'sandbox':str(root),'crash_exit':73,'recovery_status':result['status'],'no_replay':True}
        case('actual_process_death_'+point,death)
    def pending_damage(kind):
        root=fixture();before=S.sha(reg(root))
        def fault(point,target):
            if point=='after_save':raise OSError('PENDING')
        expect(lambda:close(root,fault));op=S.history(root)['active']
        target=Path(op['backup']['path']) if kind=='backup' else Path(op['preservation']['path'])
        target.write_bytes(target.read_bytes()+b'broken')
        e=expect(lambda:core.recover(root),'BACKUP_DAMAGED' if kind=='backup' else 'PRESERVATION_EVIDENCE_DAMAGED')
        assert S.sha(reg(root))==before and 'COMMITTED' not in states(root)
        return {'sandbox':str(root),'refusal':e,'registry_unchanged':True}
    case('damaged_verified_backup_blocks_recovery',lambda:pending_damage('backup'))
    case('damaged_preservation_blocks_recovery',lambda:pending_damage('evidence'))
    def journal_damage(missing):
        root=fixture();close(root);after=S.sha(reg(root))
        if missing:
            for p in (root/'journal').glob('*.json'):p.unlink()
        else:
            p=root/'journal/00000001.json';p.write_bytes(b'{broken')
        e=expect(lambda:plan(root),'REGISTRY_CHANGED_OR_JOURNAL_LOST' if missing else 'UNREADABLE_JSON')
        assert S.sha(reg(root))==after;return {'sandbox':str(root),'refusal':e,'no_second_debit':True}
    case('corrupt_journal_refuses_write',lambda:journal_damage(False))
    case('missing_journal_refuses_second_debit',lambda:journal_damage(True))
    def parallel():
        root=fixture();p=plan(root);save(root/'PARALLEL_PLAN.json',p);barrier=root/'GO'
        args=[sys.executable,'-B',str(__file__),'worker','parallel',str(root/'PARALLEL_PLAN.json'),str(barrier)]
        workers=[subprocess.Popen(args,env=ENV,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf8') for _ in range(2)]
        time.sleep(.4);barrier.write_text('go');out=[]
        for worker in workers:
            stdout,stderr=worker.communicate(timeout=60);out.append({'exit':worker.returncode,'stdout':stdout,'stderr':stderr})
        result=[json.loads(x['stdout']) for x in out];assert sum(x['status']=='CLOSED_TEST' for x in result)==1
        assert all(x['status'] in ('CLOSED_TEST','ALREADY_CLOSED_TEST','STOP') for x in result)
        assert balances(root)=={'s8-a':9,'s8-b':6} and states(root).count('COMMITTED')==1
        return {'sandbox':str(root),'processes':out,'balances':balances(root)}
    case('two_actual_processes_one_debit',parallel)
    def previous_backups():
        root=fixture();r1=close(root);first=Path(r1['operation']['backup']['path']);first_sha=S.sha(first)
        for p in (root/'inputs').glob('*.json'):
            x=S.read(p);x['passport']='8002/2099';save(p,x)
        q=S.read(request(root))
        for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
        save(request(root),q);r2=close(root);assert S.sha(first)==first_sha
        assert first!=Path(r2['operation']['backup']['path']) and balances(root)=={'s8-a':6,'s8-b':2}
        return {'sandbox':str(root),'first_backup_sha256':first_sha,'backups':2,'balances':balances(root)}
    case('two_passports_previous_backups_retained',previous_backups)
    def edited_registry(kind):
        root=fixture()
        with excel_session() as excel:
            b=excel.Workbooks.Open(Filename=str(reg(root)),UpdateLinks=0,ReadOnly=False,CorruptLoad=0)
            try:
                s=b.Worksheets('УЗК')
                if kind=='duplicate':
                    s.Cells(27,1).Value2='STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ';s.Cells(27,2).Value2='S8-A';s.Cells(27,3).Value2=20
                elif kind=='negative':s.Cells(24,3).Value2=-1
                else:s.Cells(24,3).Value2=3;s.Cells(25,3).Value2=4
                b.Save()
            finally:b.Close(SaveChanges=False)
        reseed(root);before=S.sha(reg(root))
        if kind=='zero':close(root);assert balances(root)=={'s8-a':0,'s8-b':0}
        else:
            e=expect(lambda:plan(root),'UZK_MARK_MISSING_OR_DUPLICATE' if kind=='duplicate' else 'QUANTITY_INVALID')
            assert S.sha(reg(root))==before and not states(root)
        return {'sandbox':str(root),'case':kind,'balances':balances(root) if kind=='zero' else None}
    for kind in ('duplicate','negative','zero'):case('native_registry_'+kind,lambda kind=kind:edited_registry(kind))
    def com_disconnect():
        import win32process
        root=fixture();before=S.sha(reg(root));owned={}
        @contextmanager
        def tracked():
            with excel_session() as excel:
                owned['excel']=excel;owned['pid']=win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
                yield excel
        def fault(point,target):
            if point=='save_io':owned['excel'].Quit()
        with patch('stage8_closure.native.excel_session',tracked):e=expect(lambda:close(root,fault),'MICROSOFT_EXCEL_FAILURE')
        assert S.sha(reg(root))==before and states(root)==['INTENT','BACKUP_VERIFIED']
        result=fresh(root);assert result.returncode==0 and json.loads(result.stdout)['status']=='ABORTED_UNCHANGED'
        return {'sandbox':str(root),'actual_com_failure':e,'owned_excel_pid':owned['pid'],'registry_unchanged':True,'journal_no_false_committed':True,'recovery':'ABORTED_UNCHANGED'}
    case('real_excel_disconnect_at_save',com_disconnect)
    # Independent adversarial writer uses another process and the real OS rename boundary.
    def external_replace():
        root=fixture();before=S.sha(reg(root));attempt=[]
        def fault(point,target):
            if point=='before_replace':
                incoming=root/'transactions/EXTERNAL_WRITER.xlsm';shutil.copyfile(reg(SEED),incoming)
                # Different valid XLSM bytes written by a competing publisher.
                incoming.write_bytes(incoming.read_bytes()+b'EXTERNAL_CONCURRENT_PUBLISH')
                code='import os,sys;os.replace(sys.argv[1],sys.argv[2]);print("REPLACED")'
                r=subprocess.run([sys.executable,'-B','-c',code,str(incoming),str(reg(root))],capture_output=True,text=True,timeout=10)
                attempt.append({'exit':r.returncode,'stdout':r.stdout,'stderr':r.stderr,'registry_sha_at_boundary':S.sha(reg(root))})
        result=close(root,fault)
        blocked=bool(attempt and attempt[0]['exit']!=0)
        detail={'sandbox':str(root),'expected':'external publisher denied or operation STOP before overwriting its change',
                'external_process':attempt,'closure_status':result['status'],'final_balances':balances(root),'protection_passed':blocked,'initial_sha256':before}
        save(RUN/'EXTERNAL_REPLACE_EVIDENCE.json',detail)
        assert blocked, 'External os.replace succeeded under claimed guard, then product silently overwrote it and COMMITTED'
        return detail
    case('external_process_replace_before_publish',external_replace)
    save(RUN/'GUI_FIXTURE_READY.json',{'sandbox':str((r:=fixture())),'request':str(request(r))})
    for i in range(3):
        def stability(i=i):
            root=fixture();r=close(root);assert r['status']=='CLOSED_TEST' and core.recover(root)['registry_writes']==0
            return {'sandbox':str(root),'round':i+1,'native_save_reopen':'PASS'}
        case('office_stability_'+str(i+1),stability)

if __name__=='__main__':run()
