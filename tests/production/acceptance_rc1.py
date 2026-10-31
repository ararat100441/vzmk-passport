"""Targeted acceptance of the existing binary, without Computer Use or rebuild.

Injects an I/O failure by locking ONLY the test registry AFTER ReplaceFile.
The original EXE and its Office/confirmation guards remain unchanged.
"""
import copy,hashlib,json,os,subprocess,sys,threading,time,zipfile
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app.storage import read,save,sha
from production_app import model
import win32file,win32con

def main():
    p=__import__('argparse').ArgumentParser();p.add_argument('--base',required=True);p.add_argument('--output',required=True);args=p.parse_args()
    base=Path(args.base).resolve();out=Path(args.output).resolve();out.mkdir(exist_ok=False)
    accepted=read(base/'RESULT.json');exe=Path(accepted['exe']);binary_hash=sha(exe)
    env=dict(os.environ);env['LOCALAPPDATA']=str(base/'local');env['PATH']=str(Path(os.environ['SystemRoot'])/'System32')
    checks=[]
    def run(task,request,expected='PASS',observer=None):
        folder=out/'calls'/uuid4().hex;folder.mkdir(parents=True);save(folder/'request.json',request)
        proc=subprocess.Popen([str(exe),'--task',task,'--request',str(folder/'request.json'),'--result',str(folder/'result.json')],env=env)
        stop=threading.Event();thread=threading.Thread(target=observer,args=(stop,),daemon=True) if observer else None
        if thread:thread.start()
        try:proc.wait(timeout=600)
        finally:
            stop.set()
            if thread:thread.join(timeout=10)
        result=read(folder/'result.json');assert result['status']==expected,result
        supervisor=read(result['office_supervisor'])
        assert supervisor['owned_orphans']==0 and not supervisor['ownership_errors'],supervisor
        assert all(r['exit_code']==0 and not r.get('forced_termination_attempted') and not r['still_running'] for r in supervisor['instances']),supervisor
        checks.append({'task':task,'expected':expected,'actual':result['status'],'result_path':str(folder/'result.json'),'office':supervisor})
        print(task,result['status'],result.get('error',''),flush=True)
        return result,str(folder/'result.json')
    # Existing business documents must remain byte-identical after refused retries.
    ready=Path(accepted['manifest']).parent
    ready_hashes={str(p):sha(p) for p in ready.rglob('*') if p.is_file()}
    original_check=next(read(p) for p in (base/'calls').glob('*/request.json') if 'facts' in read(p))
    old_plan=next(read(p)['plan'] for p in (base/'calls').glob('*/result.json') if read(p).get('data',{}).get('new'))
    run('prepare',{'path':old_plan},'FAIL')
    cfg=accepted['cfg'];registry=Path(cfg['registry']);before=sha(registry)
    request=copy.deepcopy(original_check)
    next_number=model.select(model.read_registry(registry),request['project']).next_number
    request['number']=next_number;request['facts']['number']=next_number
    vom=out/'COPY_VOM.xlsx';vom.write_bytes(Path(request['vom']).read_bytes());request['vom']=str(vom)
    # TEST/SAMPLE_ONLY never qualifies for real issue, even with all UI approvals.
    bad=copy.deepcopy(request);bad['cfg']['training']=False
    run('check',bad,'FAIL');assert sha(registry)==before
    checked,_=run('check',request);plan=checked['plan'];data=read(plan)
    run('prepare',{'path':plan});run('verify',{'path':plan});assert sha(registry)==before
    # This must fail before creating an operation folder or registry backup.
    no_confirm,_=run('commit',{'path':plan,'release_result':'MISSING'},'FAIL')
    assert 'confirmation.json' in str(no_confirm) and sha(registry)==before
    assert not (registry.parent/'.vzmk-operations'/data['id']).exists()
    run('confirm',{'path':plan,'responsible':request['facts']['responsible'],'reviewed':False},'FAIL')
    assert sha(registry)==before
    original_vom=vom.read_bytes()
    with zipfile.ZipFile(vom,'a') as z:z.comment=b'OWNED TEST COPY: SHA drift injection'
    run('verify',{'path':plan},'FAIL');assert sha(registry)==before
    vom.write_bytes(original_vom)
    run('confirm',{'path':plan,'responsible':request['facts']['responsible'],'reviewed':True})
    _,released=run('release',{'path':plan});assert sha(registry)==before
    journal=registry.parent/'.vzmk-operations'/data['id']/'journal.json'
    fault={'injected':False}
    # Do not observe the atomically replaced journal. Even DELETE-sharing
    # readers can briefly leave a replaced target pending deletion on Windows.
    # Candidate creation proves commit already holds the original registry;
    # the new registry becomes openable only after the native replacement.
    candidate=journal.parent/'candidate.xlsx'
    expected_candidate=read(Path(plan).parent/'prepared.json')['snapshot_sha256']
    def lock_after_replace(stop):
        handle=None
        try:
            while not stop.is_set():
                if candidate.exists() or fault.get('candidate_seen'):
                    fault['candidate_seen']=True
                    try:handle=win32file.CreateFile(str(registry),win32con.GENERIC_READ,0,None,win32con.OPEN_EXISTING,0,None)
                    except Exception:time.sleep(.001);continue
                    chunks=[]
                    while True:
                        _,b=win32file.ReadFile(handle,1048576)
                        if not b:break
                        chunks.append(b)
                    h=hashlib.sha256(b''.join(chunks)).hexdigest()
                    if h==expected_candidate:
                        fault.update(injected=True,phase='AFTER_REPLACEFILE',observed_sha256=h)
                        stop.wait(60);break
                    handle.Close();handle=None
                time.sleep(.002)
        except BaseException as exc:fault['observer_error']=repr(exc)
        finally:
            if handle is not None:handle.Close()
    failed,_=run('commit',{'path':plan,'release_result':released},'FAIL',lock_after_replace)
    assert fault['injected'],fault
    assert sha(registry)!=before and read(journal)['phase'] in ('REGISTRY_REPLACED','REPLACE_INTENT')
    run('recover',{'journal':str(journal)})
    assert sha(registry)==before and not Path(data['target']).exists()
    run('recover',{'journal':str(journal)})
    assert all(sha(p)==h for p,h in ready_hashes.items())
    assert sha(exe)==binary_hash
    save(out/'RESULT.json',{'status':'PASS','exe':str(exe),'exe_sha256':binary_hash,'plan':plan,'journal':str(journal),
        'missing_confirmation_refused':True,'sha_drift_refused':True,'test_as_actual_refused':True,
        'existing_documents_preserved':True,'post_replace_io_fault':fault,'recovery':'SHA_EXACT_ROLLBACK',
        'checks':checks,'forced_office_terminations':0,'owned_office_orphans':0})
    print('SUPPLEMENTARY_ACCEPTANCE_PASS',flush=True)

if __name__=='__main__':main()
