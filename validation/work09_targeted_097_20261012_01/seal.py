from audit import *
setup=read(OUT/'SETUP.json');proc=read(OUT/'FINAL_PROCESS_AUDIT.json');integrity=read(OUT/'FINAL_IMMUTABLE_INTEGRITY.json')
assert proc['status']=='PASS' and integrity['status']=='PASS' and integrity['exe_sha256']==EXPECTED
result={'version':'0.9.7','exe':str(EXE),'exe_sha256':EXPECTED,'verdict':'FAIL','blocker':0,'major':1,'minor':1,'findings':['W09-097-M01','W09-097-m01'],'ready_for_user_e2e':False,'ready_for_production':False,'frozen_CLI_E2E':'PASS; 28 tasks','GUI_E2E':'FAIL; actual repeated viewer open and first shutdown regression','negative':'17/17 PASS','ordinary_normal_probes':{'PASS':11,'FAIL':1,'additional_forged_ownership_mode_normal':'expected STOP'},'old_returned_quit_WM_CLOSE_fix':'PASS in targeted cases; overall reliability has retained FAIL/UNKNOWN','old_entrypoint_minor1':'PASS; viewer diagnostic loss remains separately','minor2':'PASS; 18 physical / 21 logical for three CLI bundles','process_safety':'PASS; no foreign termination found','orphan':0,'unknown_identities':0,'source_integrity':'PASS; 64536 protected files unchanged','production_operations_found':False,'historical_production_writer':'UNKNOWN','product_modified':False,'review_roots':[str(OUT),setup['base'],setup['parent'],read(OUT/'GUI_NAVIGATION_FIXTURE.json')['root']],'not_run':['full GUI TEST closure/restore and different-bundle switching after product GUI defect','viewer frozen COM/error/emergency-force injection','minimum-size/scale matrix','physical Office offline','new AKZ variants 1-3','visual review of all other Word pages','historical systemwide writer/file-IO trace; user Office continuity NOT_VERIFIED']}
save(OUT/'REVIEW_RESULT.json',result)
save(OUT/'GUI_E2E_STATUS.json',{'status':'FAIL','visible_check_generate_verify':'PASS','visible_physical_count':'6 files / 7 records','first_viewer':'PASS; 3 DOCX + 3 XLSX, ReadOnly, SHA unchanged','second_consecutive_open':'FAIL; FileExistsError on same viewer diagnostic','open_after_error_dismiss':'PASS; native REOPENED snapshot','first_shutdown_with_viewer':'FAIL; Office exited 0, main window remained','second_shutdown':'window gone and GUI identity absent','all_source_only_tests_separately_labelled':True,'controller_input_issues':'REVIEW_CONTROLLER_INCIDENTS.json','full_GUI_closure_restore':'NOT_RUN after product defect'})
product=read(OUT/'PROTECTED_AFTER.json')['files']
print('Seal verifying product bindings',len(product),flush=True)
actual=hashes([Path(p) for p in product]);assert actual==product
source=Path(setup['source']);external_roots=[Path(setup['base']),Path(setup['parent']),Path(read(OUT/'GUI_NAVIGATION_FIXTURE.json')['root']),source]
extpaths=sorted({p for root in external_roots for p in root.rglob('*') if p.is_file()})
external=hashes(extpaths)
handoff=read(ROOT/'validation/stage9_092_20261012_01/EVIDENCE_MANIFEST.json')['files']
source_matches=[]
for p,h in external.items():
    if Path(p).is_relative_to(source):
        if p in handoff:assert h==handoff[p];source_matches.append(p)
verify={'status':'PASS','product_references':len(product),'external_references':len(external),'source_handoff_matches':source_matches,'mismatches':[],'orphan':0,'unknown_identities':0,'production_SHA_gate':False,'working_registry_excluded':True,'checksum_companion_excluded_to_avoid_self_reference':True,'manifest_will_be_written_last':True}
save(OUT/'SEAL_VERIFY.json',verify)
own={str(p.relative_to(OUT)):sha(p) for p in sorted(OUT.rglob('*')) if p.is_file() and p.name not in ('EVIDENCE_MANIFEST.json','EVIDENCE_MANIFEST_SHA256.json')}
bindings={'review':own,'external':external,'product':product}
aggregate=hashlib.sha256(json.dumps(bindings,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
manifest={'schema':'WORK09_INDEPENDENT_TARGETED_097_EVIDENCE_V1','created_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'review_root':str(OUT),'product_root':str(ROOT),'actual_exe':str(EXE),'exe_sha256':EXPECTED,'review_result':result,'files':own,'external_files':external,'product_bindings':product,'aggregate_sha256':aggregate,'reference_counts':{'review':len(own),'external':len(external),'product':len(product),'total':len(own)+len(external)+len(product)},'reference_mismatches':0,'production_registry_included':False,'production_registry_constant_SHA_gate':False,'manifest_written_last':True,'checksum_companion':'EVIDENCE_MANIFEST_SHA256.json','historical_readback_evidence':'HANDOFF_READBACK.json','source_before_after_evidence':'FINAL_IMMUTABLE_INTEGRITY.json','process_final_evidence':'FINAL_PROCESS_AUDIT.json'}
payload=json.dumps(manifest,ensure_ascii=False,indent=2,default=str).encode('utf8');checksum=hashlib.sha256(payload).hexdigest()
save(OUT/'EVIDENCE_MANIFEST_SHA256.json',{'file':str(OUT/'EVIDENCE_MANIFEST.json'),'sha256':checksum,'aggregate_sha256':aggregate,'references':manifest['reference_counts'],'written_before_manifest':True})
# This is the last artifact write. Every previously listed review file is stable.
target=OUT/'EVIDENCE_MANIFEST.json'
with target.open('xb') as f:f.write(payload);f.flush();os.fsync(f.fileno())
assert sha(target)==checksum
assert all(sha(OUT/p)==h for p,h in own.items())
print(json.dumps({'status':'SEALED','manifest_sha256':checksum,'references':manifest['reference_counts'],'verdict':result['verdict'],'findings':[0,1,1]},ensure_ascii=False),flush=True)
