"""Seal current GUI evidence together with the unchanged prior CLI/source/build bindings."""
from prepare import *
save(OUT/'USER_AUTHORIZATION.json',{'source':'Direct user message in this thread','scope':'Computer Use exclusively for limited DEV GUI E2E of frozen EXE 0.9.9: launch, TEST Word/Excel, reopen cycles, switch bundles, FileExistsError absence, Word.Close error display, normal exit, orphan check','forbidden':['production writes','real UZK deductions','Stage1-8 changes','unknown ownership process termination','closing user Word/Excel','deployment','independent Work launch'],'unexpected_non_test_office_exit_rule':'Immediate STOP and diagnosis','historical_word_3132_cause':'UNKNOWN'})
v=read(OUT/'GUI_VALIDATION.json');binding=read(OUT/'BINDING_AFTER.json')
assert v['status']=='PASS' and v['ready_for_work09_review']
prior_manifest=PRIOR/'EVIDENCE_MANIFEST.json';prior=read(prior_manifest)
files=dict(prior['files'])
def add(p,h=None):
    p=str(Path(p));value=sha(p) if h is None else h
    if p in files:assert files[p]==value,('BINDING_CONFLICT',p)
    assert sha(p)==value,('HASH_MISMATCH',p)
    files[p]=value
for p in (prior_manifest,PRIOR/'EVIDENCE_MANIFEST_SHA256.json',PRIOR/'SEAL_VERIFY.json'):add(p)
for p,h in binding['source']['files'].items():add(p,h)
for p,h in binding['accepted_files'].items():add(p,h)
for p,h in binding['test_files'].items():add(p,h)
for rel,h in binding['build']['files'].items():add(Path(binding['build']['exe']).parent/rel,h)
for p,h in read(OUT/'DIAGNOSTIC_SOURCE_BINDING.json').items():add(p,h)
for p in OUT.rglob('*'):
    if p.is_file():add(p)
aggregate=hashlib.sha256(json.dumps(files,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
manifest=OUT/'EVIDENCE_MANIFEST.json'
save(manifest,{'schema':'STAGE9_099_LIMITED_GUI_E2E_SELF_EVIDENCE_V1','version':'0.9.9','exe':binding['build']['exe'],'exe_sha256':binding['build']['exe_sha256'],'gui_e2e':'PASS_LIMITED_DEV_SCOPE','cycles':6,'expected_negative_close_failures':1,'owned_office_instances':12,'owned_orphans':0,'forced_terminations':0,'prior_manifest':str(prior_manifest),'prior_manifest_sha256':sha(prior_manifest),'prior_files_unchanged':2160,'independent_review':'NOT_RUN_USER_PROHIBITED','ready_for_work09_review':True,'ready_for_user_e2e':False,'ready_for_production':False,'files':files,'file_count':len(files),'aggregate_sha256':aggregate,'seal_exclusions':['this EVIDENCE_MANIFEST.json','EVIDENCE_MANIFEST_SHA256.json','SEAL_VERIFY.json'],'scope_exclusions':['production business files','global Office addins/settings','real user documents']})
manifest_sha=sha(manifest)
save(OUT/'EVIDENCE_MANIFEST_SHA256.json',{'manifest':str(manifest),'sha256':manifest_sha,'aggregate_sha256':aggregate})
bad=[p for p,h in files.items() if not Path(p).is_file() or sha(p)!=h]
assert not bad,bad
save(OUT/'SEAL_VERIFY.json',{'status':'PASS','manifest_sha256':manifest_sha,'file_count':len(files),'mismatch_count':len(bad),'gui_e2e':'PASS_LIMITED_DEV_SCOPE','stop':True})
print(json.dumps({'status':'PASS','sealed_files':len(files),'mismatches':0,'manifest_sha256':manifest_sha,'ready_for_work09_review':True,'Work_launched':False},indent=2))
