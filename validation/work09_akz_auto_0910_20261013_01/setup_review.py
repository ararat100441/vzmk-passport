from common import *
import concurrent.futures
if (OUT/'SETUP.json').exists():
 s=read(OUT/'SETUP.json');tag=s['tag'];parent=Path(s['parent']);base=Path(s['base'])
else:
 tag='WA10_'+uuid4().hex[:8];parent=Path.home()/tag;parent.mkdir();base=parent/'local';(base/'tmp').mkdir(parents=True);(base/'source_tests').mkdir()
old=read(PRIOR/'SETUP.json')
if not (OUT/'SETUP.json').exists():save(OUT/'SETUP.json',{'parent':str(parent),'base':str(base),'temp':str(base/'tmp'),'source':old['source'],'vom':old['vom'],'tag':tag,'GUI_authorized_by_current_user':True})
assert sha(EXE)==EXPECTED
build=read(PRIOR/'BUILD_MANIFEST.json');files={str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}
assert files==build['files'] and len(files)==1119
sources=read(PRIOR/'SOURCE_AT_BUILD.json');assert all(sha(p)==h for p,h in sources.items())
accepted=read(PRIOR/'ACCEPTED_BEFORE.json');assert len(accepted)==56 and all(sha(p)==h for p,h in accepted.items())
seed=read(PRIOR/'SEED_BEFORE.json');assert all(sha(p)==h for p,h in seed['files'].items())
for name,data in [('SOURCE_BEFORE.json',sources),('ACCEPTED_BEFORE.json',accepted),('SEED_BEFORE.json',seed),('BUILD_BEFORE.json',{'exe':str(EXE),'exe_sha256':sha(EXE),'files':files,'onedir_digest':build['onedir_digest']})]:
 if not (OUT/name).exists():save(OUT/name,data)
 else:assert read(OUT/name)==data
m=read(PRIOR/'EVIDENCE_MANIFEST.json');refs={str(PRIOR/p):r['sha256'] for p,r in m['evidence_files'].items()}|{p:r['sha256'] for p,r in m['external_files'].items()}
save(OUT/'HANDOFF_MANIFEST_READ.json',{'fields':list(m),'file_count':len(refs),'sha256':sha(PRIOR/'EVIDENCE_MANIFEST.json'),'runtime_proof':False})
with concurrent.futures.ThreadPoolExecutor(10) as pool:bad=[p for p,h,actual in pool.map(lambda x:(x[0],x[1],sha(x[0])),refs.items()) if h!=actual]
assert not bad,bad
save(OUT/'HANDOFF_READBACK.json',{'status':'PASS','files':len(refs),'mismatch':bad,'manifest_sha256':sha(PRIOR/'EVIDENCE_MANIFEST.json'),'runtime_proof':False})
paths=[p for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts]
with concurrent.futures.ThreadPoolExecutor(10) as pool:protected=dict(pool.map(lambda p:(str(p),sha(p)),paths))
save(OUT/'PROTECTED_BEFORE.json',{'files':protected,'working_registry_SHA_gate':False})
print('Setup PASS',len(protected),'protected',len(refs),'handoff refs',str(parent),flush=True)
