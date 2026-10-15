"""Independent TEST-only reviewer. No product files are modified."""
import sys,os,json,shutil,hashlib,subprocess,time,gc,traceback,io,copy
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile
from uuid import uuid4
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
BATCH=RUN/os.environ.get('STAGE8_REVIEW_BATCH','.')
sys.path.insert(0,str(ROOT));sys.dont_write_bytecode=True
from stage8_closure import core,recovery,native,office,storage as S
import baseline_helpers as H
import psutil
ENV=os.environ|{'PYTHONPATH':str(ROOT),'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1'}
def save(p,v):Path(p).write_bytes(json.dumps(v,ensure_ascii=False,indent=2).encode('utf8'))
reg=H.reg;req=H.req;states=H.states;balances=H.balances
def snap(r):return {'registry':S.sha(reg(r)) if reg(r).exists() else None,'journal':{p.name:S.sha(p) for p in (r/'journal').glob('*.json')}}
def proc(*a):
    w=subprocess.run([sys.executable,'-B',str(__file__),'sub',*map(str,a)],capture_output=True,text=True,encoding='utf8',env=ENV,timeout=120)
    return {'exit':w.returncode,'stdout':w.stdout,'stderr':w.stderr}
def result(w):assert w['exit']==0,w;return json.loads(w['stdout'])
def stop(fn,code=None):return H.stopped(fn,code)
def fixture():return H.fixture()
def pclose(r):return core.preview(r,req(r))
def doclose(r):p=pclose(r);return core.close(p,core.confirmation(p))
def prest(r):return recovery.preview(r)
def dorestore(r):p=prest(r);return recovery.restore(p,recovery.confirmation(p))
def partial():
    r=fixture();p=pclose(r);save(r/'CLOSE_PREVIEW.json',p)
    w=proc('close_crash',r/'CLOSE_PREVIEW.json','during_publish');assert w['exit']==73,w
    h=S.history(r);op=h['active'];assert op['state']=='APPLYING' and not h['closed']
    assert S.sha(reg(r)) not in (op['before_sha256'],op['after_sha256'])
    S.verify_backup(op['backup'],r)
    return r,w
def refuse(r,fn,code=None):
    b=snap(r);e=stop(fn,code);assert snap(r)==b
    return {'error':e,'unchanged':b}
def native_exact(r,backup):
    from xml.etree import ElementTree as E
    import olefile
    def streams(p):
        with ZipFile(p) as z,olefile.OleFileIO(io.BytesIO(z.read('xl/vbaProject.bin'))) as o:
            return {'/'.join(n):hashlib.sha256(o.openstream(n).read()).hexdigest() for n in o.listdir()}
    assert Path(backup).read_bytes()==reg(r).read_bytes()
    with ZipFile(backup) as a,ZipFile(reg(r)) as b:
        parts={n:hashlib.sha256(b.read(n)).hexdigest() for n in b.namelist()}
        assert set(a.namelist())==set(parts)
        assert all(a.read(n)==b.read(n) for n in parts)
        workbook=E.fromstring(b.read('xl/workbook.xml'))
        names=[E.tostring(n,encoding='unicode') for n in workbook if n.tag.endswith('definedNames')]
        rels={n:parts[n] for n in parts if n.endswith('.rels')}
    vba=streams(reg(r));assert len(vba)==11 and vba==streams(backup)
    info=native.inspect(reg(r),{'УЗК!C24':12,'УЗК!C25':10},diagnostics=RUN/'native_independent')
    assert info['normal_open'] and info['formulas']['S8_FORMULAS!$A$1']['value']==30
    assert info['formulas']['S8_FORMULAS!$A$2']['value']==5
    assert next(s for s in info['sheets'] if s['name']=='S8_FORMULAS')['visible']==2
    return {'all_zip_parts_exact':parts,'all_11_vba_streams':vba,'defined_names':names,'relationships':rels,'native':info}
def supported():
    r,w=partial();op=S.history(r)['active'];damaged=reg(r).read_bytes();b=snap(r)
    a=proc('recover',r);assert a['exit']==2 and 'AMBIGUOUS_OPERATION' in a['stdout'],a
    cp=proc('close_preview',r);assert cp['exit']==2,cp
    p=result(proc('restore_preview',r));assert p['payload']['phase']=='TORN'
    assert p['payload']['backup']==op['backup'] and p['payload']['verified_before_sha256']==S.sha(op['backup']['path'])
    cancel=refuse(r,lambda:recovery.restore(p,''),'EXPLICIT_CONFIRMATION')
    out=recovery.restore(p,recovery.confirmation(p));assert out['status']=='RECOVERED_ABORTED' and out['decrements']==0 and out['closed'] is False
    h=S.history(r);assert h['active'] is None and not h['closed'] and len(h['restored'])==1
    assert out['new_identity']!=out['old_identity'] and S.file_identity(reg(r))==h['expected_identity']==out['new_identity']
    assert Path(h['restored'][0]['restoration']['quarantine']).read_bytes()==damaged
    assert all(S.sha(r/'journal'/n)==v for n,v in b['journal'].items())
    n=native_exact(r,op['backup']['path']);save(RUN/'NATIVE_EXACT.json',n)
    assert states(r).count('COMMITTED')==0 and balances(r)=={'s8-a':12,'s8-b':10}
    fresh=result(proc('recover',r));assert fresh['status']=='NO_PENDING_OPERATION'
    oldclose=refuse(r,lambda:core.close(S.read(r/'CLOSE_PREVIEW.json'),core.confirmation(S.read(r/'CLOSE_PREVIEW.json'))),'FILE_ID_CHANGED')
    repeat=refuse(r,lambda:recovery.restore(p,recovery.confirmation(p)),'NO_PENDING')
    p2=result(proc('close_preview',r));save(r/'NEW_CLOSE.json',p2)
    c=result(proc('close',r/'NEW_CLOSE.json'));assert c['status']=='CLOSED_TEST'
    end=snap(r);again=result(proc('close',r/'NEW_CLOSE.json'));assert again['status']=='ALREADY_CLOSED_TEST' and snap(r)==end
    assert balances(r)=={'s8-a':9,'s8-b':6} and states(r).count('COMMITTED')==1
    save(RUN/'POSITIVE.json',{'sandbox':str(r),'restored_native':n,'restored':out})
    return {'sandbox':str(r),'crash':w,'restart':a,'close_block':cp,'cancel':cancel,'restore':out,'old_close':oldclose,'repeat_restore':repeat,'fresh_restart':fresh,'final_balances':balances(r),'states':states(r),'interrupted_operation_id':op['operation_id'],'new_close_operation_id':c['operation']['operation_id'],'old_journal_unchanged':True}
def crash_restore(point,phase):
    r,w=partial();old=reg(r).read_bytes();op=S.history(r)['active'];p=prest(r);save(r/'RESTORE.json',p)
    death=proc('restore_crash',r/'RESTORE.json',point);assert death['exit']==73,death
    assert 'COMMITTED' not in states(r);S.verify_backup(op['backup'],r)
    a=proc('recover',r);assert a['exit']==2,a
    cp=proc('close_preview',r);assert cp['exit']==2,cp
    q=result(proc('restore_preview',r));assert q['payload']['phase']==phase,(phase,q)
    h=S.history(r);rr=h['active'].get('restoration')
    if rr:assert Path(rr['quarantine'] if phase in ('VACANT','PUBLISHED') else reg(r)).read_bytes()==old
    else:assert reg(r).read_bytes()==old
    save(r/'RESUME.json',q);out=result(proc('restore',r/'RESUME.json'))
    assert out['status']=='RECOVERED_ABORTED' and out['decrements']==0
    assert S.sha(reg(r))==op['before_sha256'] and states(r).count('COMMITTED')==0
    h=S.history(r);assert Path(h['restored'][0]['restoration']['quarantine']).read_bytes()==old
    assert balances(r)=={'s8-a':12,'s8-b':10};result(proc('close_preview',r))
    return {'sandbox':str(r),'point':point,'death':death,'restart_recover_STOP':a,'restart_close_STOP':cp,'detected_phase':phase,'result':out,'backup_intact':True,'damaged_bytes_preserved':True,'states':states(r)}
def negative(kind):
    r,w=partial();op=S.history(r)['active'];p=prest(r)
    if kind=='backup_damage':Path(op['backup']['path']).write_bytes(b'BROKEN_BACKUP')
    elif kind=='backup_substitution':shutil.copyfile(H.EXTERNAL,op['backup']['path'])
    elif kind=='journal_damage':
        f=next((r/'journal').glob('*.json'));f.write_bytes(f.read_bytes()+b'BAD')
    elif kind=='unknown_identity':
        incoming=r/'transactions/unknown.xlsm';shutil.copyfile(op['backup']['path'],incoming);os.replace(incoming,reg(r))
    elif kind=='valid_unknown_inplace':reg(r).write_bytes(H.EXTERNAL.read_bytes())
    elif kind=='stage_damage':Path(op['stage']).write_bytes(b'BAD_STAGE')
    b=snap(r);x=stop(lambda:prest(r));y=stop(lambda:recovery.restore(p,recovery.confirmation(p)));assert snap(r)==b
    return {'sandbox':str(r),'kind':kind,'preview_STOP':x,'restore_STOP':y,'target_journal_unchanged':True}
def stale():
    r,w=partial();p=prest(r);save(r/'RESTORE.json',p)
    assert proc('restore_crash',r/'RESTORE.json','after_restore_intent')['exit']==73
    e=refuse(r,lambda:recovery.restore(p,recovery.confirmation(p)),'PREVIEW_STALE')
    out=dorestore(r);assert out['status']=='RECOVERED_ABORTED'
    return {'sandbox':str(r),'refusal':e,'fresh_resume':out}
def unknown_after():
    r,w=partial();out=dorestore(r);incoming=r/'transactions/unknown.xlsm';shutil.copyfile(reg(r),incoming);os.replace(incoming,reg(r))
    e=refuse(r,lambda:pclose(r),'FILE_ID_CHANGED')
    return {'sandbox':str(r),'lawful':out['new_identity'],'unknown':S.file_identity(reg(r)),'STOP':e}
def set_identity(r,passport,qa=3,qb=4):
    for f in (r/'inputs').glob('*.json'):
        q=S.read(f);q['passport']=passport
        if 'uzk' in q:
            q['uzk'][0]['quantity']=qa;q['uzk'][1]['quantity']=qb
        save(f,q)
    q=S.read(req(r))
    for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
    save(req(r),q)
def later():
    r=fixture();a=doclose(r);firstbackup=a['operation']['backup'];old=refuse(r,lambda:prest(r),'NO_PENDING')
    set_identity(r,'8002/2099');b=doclose(r);assert balances(r)=={'s8-a':6,'s8-b':2}
    kept={p.name:S.sha(p) for p in (r/'journal').glob('*.json')};before=S.sha(reg(r))
    set_identity(r,'8003/2099',1,1);p=pclose(r);save(r/'CLOSE_PREVIEW.json',p)
    assert proc('close_crash',r/'CLOSE_PREVIEW.json','during_publish')['exit']==73
    out=dorestore(r);assert S.sha(reg(r))==before and balances(r)=={'s8-a':6,'s8-b':2}
    h=S.history(r);assert len(h['closed'])==2 and len(h['restored'])==1 and h['active'] is None
    assert all(S.sha(r/'journal'/n)==v for n,v in kept.items())
    return {'sandbox':str(r),'rollback_COMMITTED_STOP':old,'later_closures_preserved':list(h['closed']),'final':balances(r),'restore':out,'old_backup_sha':firstbackup['sha256'],'expected_before':before}
def full_after():
    r=fixture();p=pclose(r);save(r/'CLOSE_PREVIEW.json',p)
    assert proc('close_crash',r/'CLOSE_PREVIEW.json','after_replace')['exit']==73
    e=refuse(r,lambda:prest(r),'FULL_RESULT')
    before=S.sha(reg(r));out=result(proc('recover',r));assert out['status']=='CLOSED_TEST' and out['registry_writes']==0 and before==S.sha(reg(r))
    assert states(r).count('COMMITTED')==1 and balances(r)=={'s8-a':9,'s8-b':6}
    return {'sandbox':str(r),'restore_forbidden':e,'recover':out,'states':states(r)}
def race(point,mode):
    r,w=partial();p=prest(r);attempts=[];foreign=b'FOREIGN_CREATOR_BYTES_NO_REPLACE'
    def fault(at,target):
        if at!=point:return
        if mode=='creator':
            with reg(r).open('xb') as f:f.write(foreign);f.flush();os.fsync(f.fileno())
            attempts.append({'foreign_created':True})
        else:
            inc=r/'transactions/incoming.xlsm';shutil.copyfile(H.EXTERNAL,inc)
            dst=reg(r).parent if mode=='parent' else reg(r)
            a=H.proc('publisher',mode,inc,dst);attempts.append(a);assert a['exit']==2,a
    if mode=='creator':
        out=stop(lambda:recovery.restore(p,recovery.confirmation(p),_fault=fault),'NO_REPLACE')
        assert reg(r).read_bytes()==foreign and attempts
        h=S.history(r);assert h['active']['state']=='RESTORE_APPLYING' and not h['closed']
        rr=h['active']['restoration'];assert Path(rr['stage']).is_file() and Path(rr['quarantine']).is_file();S.verify_backup(h['active']['backup'],r)
        diag=refuse(r,lambda:prest(r),'NAMESPACE_OR_CONTENT_AMBIGUOUS')
        return {'sandbox':str(r),'restore_STOP':out,'operator_diagnostic':diag,'foreign_preserved':True,'stage_quarantine_backup_preserved':True,'attempts':attempts}
    out=recovery.restore(p,recovery.confirmation(p),_fault=fault);assert attempts and out['status']=='RECOVERED_ABORTED'
    assert balances(r)=={'s8-a':12,'s8-b':10} and states(r).count('COMMITTED')==0
    return {'sandbox':str(r),'attempts':attempts,'restore':out,'states':states(r)}
def cli():
    r,w=partial();f=r/'CLI_RESTORE.json'
    def run(*args):
        x=subprocess.run([sys.executable,'-B',str(ROOT/'stage8.py'),*map(str,args)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=120)
        return {'exit':x.returncode,'stdout':x.stdout,'stderr':x.stderr}
    p=run('restore-preview',r,'--save-preview',f);q=result(p);b=snap(r)
    cancel=run('restore',f,'--confirm','');assert cancel['exit']==2 and snap(r)==b
    x=run('restore',f,'--confirm',q['confirmation']);out=result(x);assert out['status']=='RECOVERED_ABORTED'
    return {'sandbox':str(r),'preview_process':p,'cancel_process':cancel,'restore_process':x,'states':states(r),'balances':balances(r)}
def setup():
    snapshot=[]
    for p in psutil.process_iter(['pid','name','create_time','cmdline']):
        if (p.info['name'] or '').upper()=='EXCEL.EXE':snapshot.append(p.info)
    save(RUN/'EXCEL_BEFORE.json',snapshot)
    H.prepare_seed();print('SEED_READY',H.SEED,flush=True)
def case(name,fn):
    print('START',name,flush=True);t=time.monotonic()
    try:
        detail=fn();gc.collect()
        if isinstance(detail,dict) and detail.get('pid'):assert H.wait_exit(detail['pid'])
        item={'name':name,'status':'PASS','detail':detail}
    except Exception as e:item={'name':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    item['seconds']=round(time.monotonic()-t,3);save(BATCH/'cases'/(name+'.json'),item);RESULTS.append(item)
    save(BATCH/'RESULTS.json',{'PASS':sum(x['status']=='PASS' for x in RESULTS),'FAIL':sum(x['status']=='FAIL' for x in RESULTS),'results':RESULTS})
    assert all(S.sha(p)==h for p,h in PRODUCT.items()),'PRODUCT_CHANGED'
    print(item['status'],name,item.get('error',''),flush=True)
if __name__=='__main__':
    if sys.argv[1]=='sub':
        cmd,path=sys.argv[2:4];path=Path(path)
        try:
            with office.diagnostic_scope(RUN/'worker_office'):
                if cmd=='recover':out=core.recover(path)
                elif cmd=='close_preview':out=pclose(path)
                elif cmd=='restore_preview':out=prest(path)
                else:
                    p=S.read(path)
                    def fault(point,target):
                        if cmd.endswith('_crash') and point==sys.argv[4]:os._exit(73)
                    if cmd.startswith('close'):out=core.close(p,core.confirmation(p),_fault=fault)
                    elif cmd=='restore_crash' and sys.argv[4]=='before_prepare':
                        real=Path.open
                        def opened(self,*a,**kw):
                            if self.name=='verified.xlsm' and a and a[0]=='xb':os._exit(73)
                            return real(self,*a,**kw)
                        with patch.object(Path,'open',opened):out=recovery.restore(p,recovery.confirmation(p))
                    else:out=recovery.restore(p,recovery.confirmation(p),_fault=fault)
            print(json.dumps(out,ensure_ascii=False));sys.exit(0)
        except Exception as e:print(json.dumps({'status':'STOP','error':office._error(e)},ensure_ascii=False));sys.exit(2)
    PRODUCT=S.read(RUN/'PRODUCT_BEFORE.json');H.PRODUCT={Path(p):h for p,h in PRODUCT.items()};RESULTS=[]
    (BATCH/'cases').mkdir(parents=True,exist_ok=True)
    with office.diagnostic_scope(RUN/'office'):
        if sys.argv[1]=='setup':setup();sys.exit(0)
        H.SEED=Path(S.read(RUN/'SEED.json')['sandbox']);H.EXTERNAL=RUN/'EXTERNAL_VALID_4.xlsm'
        if sys.argv[1]=='gui_fixture':
            r,w=partial();save(RUN/'GUI_FIXTURE.json',{'sandbox':str(r),'request':str(req(r)),'before':snap(r),'crash':w});print(r,flush=True);sys.exit(0)
        case('supported_recovery_full_restart_native_reclose',supported)
        for point,phase in [('before_prepare','TORN'),('restore_copy_chunk','TORN'),('after_restore_intent','PREPARED'),('after_restore_quarantine','VACANT'),('after_restore_publish','PUBLISHED'),('before_restore_terminal','PUBLISHED')]:
            case('crash_'+point,lambda point=point,phase=phase:crash_restore(point,phase))
        for kind in ['backup_damage','backup_substitution','journal_damage','unknown_identity','valid_unknown_inplace','stage_damage']:
            case('refuse_'+kind,lambda kind=kind:negative(kind))
        case('stale_restore_preview',stale)
        case('unknown_identity_after_restore',unknown_after)
        case('COMMITTED_rollback_forbidden_later_closures_preserved',later)
        case('full_after_finalize_no_rollback',full_after)
        case('restore_foreign_creator_gap_preserved',lambda:race('after_restore_quarantine','creator'))
        case('restore_external_native_replace_blocked',lambda:race('before_restore_quarantine','native'))
        case('restore_external_native_replace_after_publish_blocked',lambda:race('after_restore_publish','native'))
        case('restore_parent_rename_gap_blocked',lambda:race('after_restore_quarantine','parent'))
        case('CLI_separate_preview_cancel_confirm',cli)
        case('regression_concurrent_native_4_stale_9',lambda:H.race('native','before_replace'))
        case('regression_external_4_before_acquisition',H.prior_publish)
        case('regression_cleanup_first_error_still_Quit',lambda:H.cleanup('CutCopyMode'))
        case('regression_cleanup_Close_still_Quit',lambda:H.cleanup('Workbooks.Close',True))
        case('regression_DEV_9_2026_forbidden',H.dev_forbidden)
    print('DONE',sum(x['status']=='PASS' for x in RESULTS),sum(x['status']=='FAIL' for x in RESULTS),flush=True)
