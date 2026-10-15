import sys,os,json,hashlib,difflib,time,gc
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT));sys.path.insert(0,str(RUN))
from integrity import read,sha,save
import psutil
def digest(v):return hashlib.sha256(json.dumps(v,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
def raw_chain(root):
    tip=None;events=[];prepared={}
    for i,f in enumerate(sorted((root/'journal').glob('*.json'))):
        q=read(f);assert q['sequence']==i and q['previous']==tip and q['seal']==digest({k:v for k,v in q.items() if k!='seal'})
        if q['state']=='RESTORE_PREPARED':prepared[q['operation_id']]=q['restoration']
        if q['state']=='RECOVERED_ABORTED':
            r=prepared[q['operation_id']]
            assert q['recovery_id']==r['recovery_id'] and q['restored_identity']==r['new_identity'] and r['old_identity']!=r['new_identity']
            assert q['decrements']==0 and q['closed_status'] is None
            assert sha(q['native_evidence']['path'])==q['native_evidence']['sha256']
        events.append({'path':str(f),'sha256':sha(f),'state':q['state'],'operation_id':q['operation_id']});tip=sha(f)
    restored_ops={e['operation_id'] for e in events if e['state']=='RECOVERED_ABORTED'}
    assert not any(e['state']=='COMMITTED' and e['operation_id'] in restored_ops for e in events)
    return events
if __name__=='__main__':
    short=read(RUN/'short_run/RESULTS.json');assert short['PASS']==27 and short['FAIL']==0
    raw={}
    for q in short['results']:
        if q['name'] in ('refuse_journal_damage',):continue
        r=Path(q['detail']['sandbox']) if 'sandbox' in q['detail'] else None
        if r:raw[str(r)]=raw_chain(r)
    gui_short=read(RUN/'GUI_SHORT_FIXTURE.json');raw[gui_short['sandbox']]=raw_chain(Path(gui_short['sandbox']))
    save('RAW_JOURNAL_CHAIN.json',{'status':'PASS','sandboxes':raw,'all_prefix_links_and_seals_independently_recomputed':True,'recovered_operations_have_no_COMMITTED':True})
    from stage8_closure import storage as S,core
    e=read(RUN/'GUI_01_EMERGENCY.json');c=read(RUN/'GUI_03_CANCEL.json');good=read(RUN/'GUI_06_SHORT_SUCCESS.json');p=read(RUN/'GUI_07_NEW_PREVIEW.json')
    assert e==c and e['close']=='disabled' and e['restore']=='normal' and not e['confirmation']
    assert 'RECOVERED_ABORTED' in good['text'] and 'Списаний: 0' in good['text'] and good['plan'] is None and good['restore_plan'] is None
    assert p['plan']['payload']['status']=='READY_TEST' and p['close']=='normal' and not p['confirmation'] and p['restore']=='disabled'
    assert [(d['before'],d['quantity'],d['after']) for d in p['plan']['payload']['deductions']]==[(12,3,9),(10,4,6)]
    h=S.history(Path(gui_short['sandbox']));assert not h['closed'] and h['active'] is None and len(h['restored'])==1
    restored=h['restored'][0];assert sha(Path(gui_short['sandbox'])/'registry/УЗК_TEST.xlsm')==restored['backup']['sha256']
    long=read(RUN/'GUI_FIXTURE.json');r=Path(long['sandbox'])
    assert sha(r/'registry/УЗК_TEST.xlsm')==long['before']['registry'] and {f.name:sha(f) for f in (r/'journal').glob('*.json')}==long['before']['journal']
    assert S.history(r)['active']['state']=='APPLYING' and not S.history(r)['closed']
    save('GUI_VALIDATION.json',{'status':'PASS','real_visible_App':True,'input':'Computer Use sky, actual chooser/buttons/modal; no invoke or messagebox patches','default_no_cancel_no_changes':True,'cancel_widget_snapshot_equal':e==c,'short_restore_RECOVERED_ABORTED_zero_decrement':True,'new_preview_12_3_9_and_10_4_6_unconfirmed':True,'long_path_UI_failure_reproduced_and_target_journal_unchanged':True,'long_sandbox':str(r),'short_sandbox':gui_short['sandbox']})
    FIX=ROOT/'validation/stage8_recovery_fix_20261011_01';product=read(RUN/'PRODUCT_BEFORE.json');prior=read(ROOT/'validation/stage8_independent_targeted_20261011_01/PRODUCT_BEFORE.json');diff='';beforechecks=[]
    for p,h in product.items():
        name=Path(p).name;a=FIX/'after_product'/name;b=FIX/'before_product'/name
        if name!='recovery.py':
            assert b.exists() and sha(b)==prior[p];beforechecks.append(str(b))
        assert sha(a)==h
        aa=a.read_text(encoding='utf8').splitlines(keepends=True);bb=b.read_text(encoding='utf8').splitlines(keepends=True) if b.exists() else []
        if aa!=bb:diff+=''.join(difflib.unified_diff(bb,aa,fromfile='before/'+name,tofile='after/'+name))
    supplied=(FIX/'EXACT_PRODUCT.diff').read_text(encoding='utf8')
    # Author lists code before document; compare each complete per-file diff independent of ordering.
    def sections(text):
        lines=text.splitlines(keepends=True);parts={};current=None
        for line in lines:
            if line.startswith('--- before/'):
                current=line.rstrip().split('before/',1)[1];parts[current]=[]
            assert current is not None
            parts[current].append(line)
        return {k:''.join(v) for k,v in parts.items()}
    exact=sections(diff)==sections(supplied)
    save('EXACT_DIFF_INDEPENDENT.json',{'before_snapshots_match_prior_review_product':beforechecks,'after_snapshots_match_current_product':True,'independently_generated_diff_matches_supplied':exact,'changed_files':list(sections(diff))})
    assert exact
    records=[]
    for f in RUN.rglob('START.json'):
        q=read(f)
        if q.get('owned'):records.append({'path':str(f),**q})
    # Short fixtures have their own native diagnostics.
    for f in (ROOT/'validation/stage8_ir_261011_01').rglob('START.json'):
        q=read(f)
        if q.get('owned'):records.append({'path':str(f),**q})
    known={q['pid']:q for q in records};live=[]
    for p in psutil.process_iter(['pid','name','create_time','cmdline']):
        if (p.info['name'] or '').upper()=='EXCEL.EXE':live.append(p.info)
    own=[]
    import baseline_helpers as H
    for p in live:
        if p['pid'] in known and H.creation(p['pid'])==known[p['pid']]['creation_time']:own.append(p)
    before=read(RUN/'EXCEL_BEFORE.json')
    save('FINAL_EXCEL_OWNERSHIP.json',{'ownership_records':len(records),'live_owned':own,'live_all':live,'before_unattributed_preserved':[p for p in before if p in live],'unattributed_ownership':'NOT_VERIFIED','process_actions_on_unattributed':0,'process_termination_actions_in_review':0})
    print('RAW_CHAIN',len(raw),'DIFF',exact,'OWN_LIVE',len(own),'EXCEL_ALL',[p['pid'] for p in live],flush=True)
    assert not own
