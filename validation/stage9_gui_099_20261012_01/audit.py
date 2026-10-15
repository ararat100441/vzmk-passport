"""Audit actual frozen GUI evidence, independent of UI input delivery."""
from prepare import *
import shutil
before=read(OUT/'BINDING_BEFORE.json');after=read(OUT/'BINDING_AFTER.json')
assert before['source']==after['source'] and before['build']==after['build']
assert before['accepted_files']==after['accepted_files'] and before['test_files']==after['test_files']
assert not after['processes'],'OFFICE_OR_GUI_PROCESS_REMAINS'
baseline=read(OUT/'OFFICE_BASELINE.json');monitor=read(OUT/'OFFICE_MONITOR_FINAL.json')
assert baseline['office']==[] and not monitor['foreign_stop']
assert len(monitor['instances'])==12
assert all(r['owned'] and r['parent']==47440 and r['exit_observed'] and r['exit_code']==0 for r in monitor['instances'])
assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists()
diag=Path(r'C:\Users\arara\AppData\Local\fix098_4b9519ff\new099\VZMKPassport\diagnostics')
manifests=[read(p) for p in before['manifests']]
manifest_sets={frozenset(str(Path(r['path'])) for r in m['records']):m['passport'] for m in manifests[:2]}
directories=sorted(diag.glob('viewer_*'),key=lambda p:p.stat().st_ctime_ns)
assert len(directories)==6
rows=[];bindings={};copy_root=OUT/'frozen_viewer_diagnostics';copy_root.mkdir()
for i,p in enumerate(directories):
    def one(pattern):
        found=list(p.glob(pattern));assert len(found)==1,(str(p),pattern,len(found));return read(found[0])
    opened=one('ownership_OPENED_*.json');begin=one('ownership_OPEN_BEGIN_*.json')
    close=one('ownership_CLOSE_RESULT_*.json');ownership=read(p/'ownership.json')
    assert opened['status']=='PASS' and opened['viewer_id']==p.name==close['viewer_id']
    assert opened['open_attempt_id']==begin['open_attempt_id']==close['open_attempt_id']
    paths=begin['paths'];assert len(paths)==len(set(paths))==6
    assert sum(Path(x).suffix.lower()=='.docx' for x in paths)==3
    passport=manifest_sets[frozenset(paths)];assert passport==['9/2026','9/2026','9/2026','10/2026','9/2026','9/2026'][i]
    assert len(ownership['instances'])==2 and ownership['owned_orphans']==0
    assert not ownership['diagnostic_errors']
    for child in ownership['instances']:
        assert child['owned'] and child['actual_child'] and child['handle_source']=='CREATE_PROCESS' and child['launcher_pid']==47440
        assert child['exit_code']==0 and not child['still_running'] and not child.get('forced_termination_attempted')
        matching=[r for r in monitor['instances'] if r['pid']==child['pid'] and r['creation_time']==child['creation_time']]
        assert len(matching)==1 and matching[0]['image']==child['image']
    if i==4:
        assert ownership['status']==close['status']=='FAIL'
        error=next(e for e in ownership['cleanup_errors'] if e['step']=='Word.Close')
        assert error['hresult']==-2147023174 and error['traceback'] and 'RPC' in error['error']
        assert {'Word.Close','Word.Quit'}.issubset(e['step'] for e in close['cleanup_errors'])
    else:assert ownership['status']==close['status']=='PASS' and not close['cleanup_errors']
    assert not list(p.glob('*EMERGENCY_FORCE*'))
    assert 'FileExistsError' not in '\n'.join(f.read_text(encoding='utf8') for f in p.glob('*.json'))
    shutil.copytree(p,copy_root/p.name)
    for f in p.rglob('*'):
        if f.is_file():bindings[str(f)]=sha(f)
    rows.append({'cycle':i+1,'passport':passport,'viewer_id':p.name,'open_status':opened['status'],'close_status':close['status'],'expected_close_failure':i==4,'open_attempt_id':opened['open_attempt_id'],'close_attempt_id':close['close_attempt_id'],'instances':ownership['instances'],'diagnostic':str(p/'ownership.json')})
tree=read(OUT/'GUI_14_WORD_CLOSE_ERROR_DIALOG.json')['accessibility']['tree']
assert 'Word.Close' in tree and 'Новый просмотр не запущен' in tree and 'Диагностика:' in tree
assert all(read(OUT/('GUI_'+mode+'_EXIT.json'))['exit_code']==0 for mode in ('EXISTING','FRESH'))
assert read(OUT/'GUI_FINAL_WINDOW_LIST_CONFIRMED.json')['testAndOfficeWindows']==[]
prior_manifest=PRIOR/'EVIDENCE_MANIFEST.json';prior_seal=read(prior_manifest)
mismatches=[p for p,h in prior_seal['files'].items() if not Path(p).is_file() or sha(p)!=h]
assert not mismatches,('PRIOR_SEALED_EVIDENCE_CHANGED',mismatches)
save(OUT/'PRIOR_EVIDENCE_INTEGRITY.json',{'status':'PASS','manifest':str(prior_manifest),'sha256':sha(prior_manifest),'files':len(prior_seal['files']),'mismatches':mismatches})
save(OUT/'DIAGNOSTIC_SOURCE_BINDING.json',bindings)
save(OUT/'GUI_VALIDATION.json',{'status':'PASS','scope':'LIMITED_DEV_GUI_E2E_ON_FROZEN_099','version':'0.9.9','exe_sha256':before['build']['exe_sha256'],'viewer_cycles':rows,'normal_viewer_closes':5,'expected_negative_close_failures':1,'gui_error_display':'PASS_REAL_WORD_RPC_UNAVAILABLE','file_exists_error_count':0,'owned_office_instances':12,'owned_orphans':0,'forced_terminations':0,'foreign_office_baseline_count':0,'unexpected_foreign_exits':0,'historical_word_3132_cause':'UNKNOWN','source_integrity':'PASS','accepted_code_count':56,'test_workspace_unchanged':True,'fresh_main_window':'PASS','shutdown_first_attempt':'PASS','independent_work_review':'NOT_RUN_USER_PROHIBITED','ready_for_work09_review':True,'ready_for_user_e2e':False,'ready_for_production':False})
print('GUI AUDIT PASS: six cycles / one expected negative / 12 Office / orphan 0 / force 0 / prior seal',len(prior_seal['files']),flush=True)
