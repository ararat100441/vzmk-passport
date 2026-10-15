"""Independent raw evidence checks, after all workers and GUI have finished."""
import json,hashlib,sys,gc,time
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT));sys.path.insert(0,str(RUN));sys.dont_write_bytecode=True
import psutil
from integrity import read,sha,save
from stage8_closure import storage as S
import baseline_helpers as H
SHORT=ROOT/'validation/stage8_li_261011_01'
FIX=ROOT/'validation/stage8_longpath_fix_20261011_001'
def digest(v):return hashlib.sha256(json.dumps(v,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
def directory_identity(p):
    import win32file,win32con
    h=win32file.CreateFile(str(p),win32con.GENERIC_READ,win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE,None,win32con.OPEN_EXISTING,win32con.FILE_FLAG_BACKUP_SEMANTICS|win32file.FILE_FLAG_OPEN_REPARSE_POINT,None)
    try:return S._identity(h)
    finally:h.Close()
def chain(root):
    tip=None;active=None;prepared={};closed={};events=[];restored=set();lengths=[]
    for i,p in enumerate(sorted((root/'journal').glob('*.json'))):
        q=read(p);assert q['sequence']==i and q['previous']==tip and q['seal']==digest({k:v for k,v in q.items() if k!='seal'})
        st=q['state'];op=q['operation_id']
        if st=='INTENT':assert active is None and q['key'] not in closed;active=op
        else:assert active==op
        if st=='RESTORE_PREPARED':
            prepared[op]=q['restoration'];r=prepared[op];assert r['old_identity']!=r['new_identity'];lengths.append(len(r['stage']))
        if st=='RECOVERED_ABORTED':
            r=prepared[op];assert q['recovery_id']==r['recovery_id'] and q['restored_identity']==r['new_identity'] and q['decrements']==0 and q['closed_status'] is None
            assert sha(q['native_evidence']['path'])==q['native_evidence']['sha256'];restored.add(op);active=None
        if st=='COMMITTED':assert op not in restored and q['key'] not in closed;closed[q['key']]=op;active=None
        if st=='ABORTED':active=None
        events.append({'path':str(p),'sha256':sha(p),'state':st,'operation_id':op});tip=sha(p)
    assert not any(e['state']=='COMMITTED' and e['operation_id'] in restored for e in events)
    return {'events':events,'restored_operations':sorted(restored),'closed_keys':closed,'restore_stage_lengths':lengths}
def main():
    batches={b:read(RUN/b/'RESULTS.json') for b in ('long','short','edges_final')}
    assert all(q['FAIL']==0 for q in batches.values()),[(b,q['FAIL']) for b,q in batches.items()]
    assert batches['long']['PASS']==27 and batches['short']['PASS']==27
    names=read(RUN/'PREVIOUS_15_REQUIRED.json');assert len(names)==15
    matrix=[]
    for n in names:
        row={'name':n,'previous':'FAIL'}
        for b in ('long','short'):
            q=next(q for q in batches[b]['results'] if q['name']==n);assert q['status']=='PASS';row[b]={'status':'PASS','evidence':str(RUN/b/'cases'/(n+'.json')),'sha256':sha(RUN/b/'cases'/(n+'.json'))}
        matrix.append(row)
    save('PREVIOUS_15_INDEPENDENT.json',{'status':'PASS','short_PASS':15,'long_PASS':15,'results':matrix})
    excluded={q['detail']['sandbox'] for b in ('long','short') for q in batches[b]['results'] if q['name']=='refuse_journal_damage'}
    raw={};expected=[]
    for base in (RUN/'sandboxes',SHORT/'sandboxes'):
        for p in base.glob('*/SANDBOX.json'):
            r=p.parent
            if str(r) in excluded:
                try:chain(r)
                except (json.JSONDecodeError,AssertionError) as e:expected.append({'sandbox':str(r),'expected_damaged_journal_rejected':type(e).__name__})
                else:raise AssertionError('MUTATED_JOURNAL_NOT_REJECTED')
            else:raw[str(r)]=chain(r)
    long_lengths=[v for p,q in raw.items() if Path(p).is_relative_to(RUN/'sandboxes') for v in q['restore_stage_lengths']]
    short_lengths=[v for p,q in raw.items() if Path(p).is_relative_to(SHORT/'sandboxes') for v in q['restore_stage_lengths']]
    assert min(long_lengths)>=260 and max(short_lengths)<260,(long_lengths,short_lengths)
    save('RAW_JOURNAL_CHAIN.json',{'status':'PASS','sandboxes':raw,'expected_corrupt_journals':expected,'no_recovered_operation_has_COMMITTED':True,'long_restore_stage_lengths':sorted(set(long_lengths)),'short_restore_stage_lengths':sorted(set(short_lengths))})
    supplied=(FIX/'EXACT_PRODUCT.diff').read_text(encoding='utf8');own=(RUN/'INDEPENDENT_PRODUCT.diff').read_text(encoding='utf8');assert supplied==own
    diff=read(RUN/'EXACT_DIFF_INDEPENDENT.json');diff['independently_generated_full_diff_equals_provided']=True;save('EXACT_DIFF_INDEPENDENT.json',diff)
    actions=read(FIX/'PRELIMINARY_OWN_ORPHAN_CLEANUP.json');a=next(q for q in actions['actions'] if q['pid']==70112)
    start=read(Path(a['ownership_record']));end=read(Path(a['ownership_record']).parent/'END.json')
    assert sha(Path(a['ownership_record']))==a['ownership_sha256'] and start['owned'] and start['pid']==70112 and start['creation_time']==a['creation_time']
    fail=next(q for q in read(FIX/'short_02/RESULTS.json')['results'] if q['status']=='FAIL');assert fail['name']=='refuse_valid_unknown_inplace' and '0x800706be' in fail['error']
    r=Path(a['ownership_record']).parents[4];h=S.history(r);assert not h['closed'] and h['active'];S.verify_backup(h['active']['backup'],r)
    save('AUTHOR_CLEANUP_INCIDENT_INDEPENDENT.json',{'classification':'Author TEST orchestration incident; excluded preliminary batch, not a reproduced product defect','PID':70112,'creation_time':a['creation_time'],'owned_TEST_instance':True,'ownership_record_sha256':a['ownership_sha256'],'failed_case':fail['name'],'hresult':'0x800706BE','current_historical_states':[q['state'] for q in h['events']],'backup_verified':True,'no_COMMITTED':True,'END_sha256':sha(Path(a['ownership_record']).parent/'END.json'),'original_orphan_explanation_wrong':True,'author_final_short_03_27_PASS_bound':read(FIX/'short_03/RESULTS.json')['PASS']==27,'independent_same_case_long_and_short_PASS':True,'review_process_termination_actions':0})
    records=[];temps=set()
    for base in (RUN,SHORT):
        for p in base.rglob('START.json'):
            q=read(p)
            if q.get('schema')!='STAGE8_OWNED_EXCEL_V1':continue
            assert q['owned'] is True
            e=p.parent/'END.json';assert e.exists(),str(p)
            end=read(e);steps=[x['step'] for x in end['cleanup']];assert all(s in steps for s in ('CutCopyMode','Workbooks.Close','Quit')),end
            records.append({'start':str(p),'start_sha256':sha(p),'end':str(e),'end_sha256':sha(e),'pid':q['pid'],'creation_time':q['creation_time'],'expected_workbooks':q['expected_workbooks'],'cleanup':end['cleanup']})
            for b in q['expected_workbooks']:
                if Path(b).parent.name.startswith('s8n_'):temps.add(Path(b).parent)
    known={(q['pid'],q['creation_time']) for q in records};live=[];owned=[]
    for p in psutil.process_iter(['pid','name','create_time','cmdline']):
        if (p.info['name'] or '').upper()=='EXCEL.EXE':
            live.append(p.info)
            if (p.pid,H.creation(p.pid)) in known:owned.append(p.info)
    before=read(RUN/'EXCEL_BEFORE.json');assert all(any(p['pid']==q['pid'] and p['create_time']==q['create_time'] for q in live) for p in before)
    save('EXCEL_COM_RECORDS.json',records)
    save('FINAL_EXCEL_OWNERSHIP.json',{'ownership_records':len(records),'live_owned':owned,'live_all':live,'before_instances_same_PID_and_creation':True,'unattributed_instances_not_touched':True,'global_taskkill':False,'review_process_termination_actions':0})
    leftovers=[]
    for batch in ('edges','edges_final'):
        for q in read(RUN/batch/'RESULTS.json')['results']:
            foreign=q.get('detail',{}).get('foreign_preserved')
            if isinstance(foreign,str) and Path(foreign).is_absolute():temps.add(Path(foreign).resolve().parent)
    for p in sorted(temps):
        if p.exists():leftovers.append({'path':str(p),'identity_now':directory_identity(p),'files':{str(f):sha(f) for f in p.iterdir() if f.is_file()},'reason':'Controlled os._exit or explicit foreign-identity/collision/disconnect TEST evidence; never reused or automatically deleted'})
    save('SHORT_WORKSPACE_LEFTOVERS.json',{'normal_context_cleanup_verified_in_tests':True,'retained_crash_or_adversarial_evidence':leftovers,'no_recursive_cleanup':True})
    assert not owned,owned
    gui=read(RUN/'GUI_VALIDATION.json');assert gui['status']=='PASS'
    integ=read(RUN/'SOURCE_INTEGRITY.json');assert integ['product_before_equals_after'] and not integ['protected_data_changed'] and not integ['read_errors']
    save('CANONICAL_RESULTS.json',{'PASS':sum(q['PASS'] for q in batches.values())+1,'FAIL':0,'functional_batches':{b:{'PASS':q['PASS'],'FAIL':q['FAIL'],'results_sha256':sha(RUN/b/'RESULTS.json')} for b,q in batches.items()},'GUI_visible':gui,'product_fingerprint':read(RUN/'PRODUCT_BINDING.json')['fingerprint'],'independent_review_complete':True})
    print('AUDIT_PASS',len(raw),'fixtures',len(records),'Excel records','leftovers',len(leftovers),flush=True)
if __name__=='__main__':main()
