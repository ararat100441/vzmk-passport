from audit import *
before=read(OUT/'PROTECTED_BEFORE.json')['files']
actual=hashes([Path(p) for p in before])
changed=[{'path':p,'before':h,'after':actual[p]} for p,h in before.items() if actual[p]!=h]
now={str(p) for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts}
added=sorted(now-set(before));removed=sorted(set(before)-now)
save(OUT/'PROTECTED_AFTER.json',{'files':actual,'status':'PASS' if not changed and not added and not removed else 'FAIL','changed':changed,'added':added,'removed':removed,'production_registry_excluded':True})
src=read(OUT/'DEV_SOURCE_BEFORE.json');after={p:sha(Path(src['source'])/p) for p in src['files']}
assert after==src['files'];save(OUT/'DEV_SOURCE_AFTER.json',{'status':'PASS','source':src['source'],'files':after,'production_registry_excluded':True})
save(OUT/'FINAL_IMMUTABLE_INTEGRITY.json',{'status':'PASS' if not changed and not added and not removed else 'FAIL','protected_files':len(before),'changed':changed,'added':added,'removed':removed,'exe_sha256':sha(EXE),'expected_exe_sha256':EXPECTED,'source_tests_before':read(OUT/'SOURCE_SHA256_BEFORE.json')['aggregate_sha256'],'source_tests_after':read(OUT/'SOURCE_SHA256_AFTER.json')['aggregate_sha256'],'production_registry_SHA_gate':False})
print('Final immutable audit',len(before),'changed',len(changed),'added',len(added),'removed',len(removed),'EXE',sha(EXE),flush=True)
