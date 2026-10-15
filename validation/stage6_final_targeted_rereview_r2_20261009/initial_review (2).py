from common_review import *
m=load(FIX/'REVIEW_MANIFEST.json');tested=load(FIX/'tested_product_hashes.json');baseline=load(FIX/'baseline.json')
current={p:sha(p) for p in tested['CURRENT']}
assert current==tested['CURRENT']==tested['BEFORE']==tested['AFTER'],'TESTED != CURRENT STOP'
assert current==m['final_tests_product']['CURRENT']
assert {n:sha(PAIR/n) for n in EXPECTED}==EXPECTED,'EXISTING PAIR CHANGED STOP'
assert sha(m['reviewed_baseline_manifest'])==m['reviewed_baseline_manifest_sha256']
prior=load(m['reviewed_baseline_manifest'])
assert baseline['product']==m['reviewed_product']==prior['product']['CURRENT']
changed=[p for p,h in baseline['product'].items() if h!=current[p]]
assert changed==[str(ROOT/'akz_stage6/core.py')]
assert sha(FIX/'BEFORE_core.py')==baseline['product'][changed[0]]
assert all(sha(p)==h for p,h in m['files'].items()),'FIX EVIDENCE SHA STOP'
guard_now={p:sha(p) for p in m['guarded_current']}
parallel=[{'path':p,'fix_manifest_sha':h,'rereview_start_sha':guard_now[p]} for p,h in m['guarded_current'].items() if guard_now[p]!=h]
assert parallel==[{'path':r['path'],'fix_manifest_sha':r['expected'],'rereview_start_sha':r['actual']} for r in load(E/'guarded_mismatches_on_start.json')]
guard={**guard_now,**m['files'],**current,str(FIX/'REVIEW_MANIFEST.json'):sha(FIX/'REVIEW_MANIFEST.json')}
oldguard=baseline['guarded'];delta=[p for p,h in oldguard.items() if sha(p)!=h]
assert set(delta)==set(changed+[r['path'] for r in parallel])
before=(FIX/'BEFORE_core.py').read_text('utf-8');after=(ROOT/'akz_stage6/core.py').read_text('utf-8')
diff=''.join(difflib.unified_diff(before.splitlines(True),after.splitlines(True),fromfile='reviewed/akz_stage6/core.py',tofile='current/akz_stage6/core.py'))
assert diff==(FIX/'product_changes.diff').read_text('utf-8')
(E/'actual_product.diff').write_text(diff,encoding='utf-8')
dump=lambda n:ast.dump(n,include_attributes=False)
ot=ast.parse(before);nt=ast.parse(after)
of={n.name:dump(n) for n in ot.body if isinstance(n,(ast.FunctionDef,ast.ClassDef))};nf={n.name:dump(n) for n in nt.body if isinstance(n,(ast.FunctionDef,ast.ClassDef))}
semantic={'changed':[n for n in of if of[n]!=nf.get(n)],'added':[n for n in nf if n not in of],'unchanged':[n for n in of if of[n]==nf.get(n)]}
assert set(semantic['changed'])=={'FileOperationError','generate'}
assert all(n in semantic['unchanged'] for n in ('write_application','write_act','matching','preview','templates','current_passport','ClimateDay','validate_saved_pair','verify_preview'))
t=load(FIX/'targeted_tests.json');results=t['results'];assert len(results)==27 and all(r['status']=='PASS' for r in results)
unit=(FIX/'unit_tests.txt').read_text('utf-8-sig');assert 'Ran 16 tests' in unit and '\nOK' in unit
assert m['written_last'] and all(Path(p).stat().st_mtime_ns<=(FIX/'REVIEW_MANIFEST.json').stat().st_mtime_ns for p in m['files'])
save('baseline_current.json',{'guarded':guard,'product':tested,'pre_fix_product':baseline['product'],'pre_fix_guard_count':len(oldguard),'pre_fix_unchanged_count':len(oldguard)-len(delta),'expected_parallel_changes':parallel,'parallel_changes_user_confirmed':'Да, изменения другой работы ожидаемы','fix_evidence_count':len(m['files']),'semantic':semantic,'existing_pair':EXPECTED,'reported_tests':{'targeted_pass':27,'unit_pass':16},'fix_manifest_last_verified':True})
save('initial_audit.json',AUDIT)
print(json.dumps({'current_product_verified':current,'only_change':changed,'fix_evidence_count':len(m['files']),'guarded_count':len(guard),'semantic':semantic,'existing_pair':'UNCHANGED','reported_targeted':'27/27 SHA authenticated'},ensure_ascii=False,indent=2))
