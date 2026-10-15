from common import *
assert read(OUT/'RESULT.json')['READY_FOR_WORK_REVIEW']=='YES'
assert read(OUT/'OFFICE_BASELINE_FINAL.json')['foreign_stop'] is False
assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists()
assert all(sha(p)==h for p,h in read(OUT/'SOURCE_AT_BUILD.json').items())
assert all(sha(p)==h for p,h in read(OUT/'ACCEPTED_BEFORE.json').items())
assert sha(EXE)==read(OUT/'BUILD_MANIFEST.json')['exe_sha256']
def walk(folder):
    files=[]
    for base,dirs,names in os.walk(folder,followlinks=False):
        dirs[:]=[d for d in dirs if not (Path(base)/d).is_symlink() and not (Path(base)/d).is_junction() and d!='__pycache__']
        for name in names:
            p=Path(base)/name
            if p.is_symlink() or p.is_junction():continue
            files.append(p)
    return sorted(files)
excluded={'EVIDENCE_MANIFEST.json','EVIDENCE_MANIFEST_SHA256.json'}
own={str(p.relative_to(OUT)):{'sha256':sha(p),'bytes':p.stat().st_size} for p in walk(OUT) if p.name not in excluded}
external={}
for p in walk(Path(SETUP['parent'])):
    if p.suffix.lower() in ('.tmp',):continue
    external[str(p)]={'sha256':sha(p),'bytes':p.stat().st_size}
for p,h in read(OUT/'SOURCE_AT_BUILD.json').items():external[p]={'sha256':h,'bytes':Path(p).stat().st_size}
for p,h in read(OUT/'ACCEPTED_BEFORE.json').items():external[p]={'sha256':h,'bytes':Path(p).stat().st_size}
for p,h in read(OUT/'SEED_BEFORE.json')['files'].items():external[p]={'sha256':h,'bytes':Path(p).stat().st_size}
for p in read(ROOT/'validation/akz_auto_20261013_01/CURRENT_PROJECT_TEMPLATES.json')['files']:external[p['path']]={'sha256':p['sha256'],'bytes':Path(p['path']).stat().st_size}
for p in walk(EXE.parent):external[str(p)]={'sha256':sha(p),'bytes':p.stat().st_size}
manifest={'schema':'STAGE9_AKZ_AUTO_REVIEW_EVIDENCE_V1','state':'SEALED','mode':'DEV_TEST_ONLY','official_issue':False,'version':'0.9.10','exe':str(EXE),'exe_sha256':sha(EXE),'generated_at_utc':time.time(),'evidence_file_count':len(own),'evidence_files':own,'external_file_count':len(external),'external_files':external,'excluded':['node_modules junctions and other links','__pycache__','external transient .tmp files','manifest self and SHA sidecar'],'GUI':'NOT_RUN_NO_NEW_PERMISSION','independent_review':'NOT_RUN_NOT_DISPATCHED','READY_FOR_WORK_REVIEW':'YES','READY_FOR_USER_E2E':'NO','READY_FOR_PRODUCTION':'NO','preserved_history':str(ROOT/'validation/akz_auto_20261013_01/EVIDENCE_MANIFEST.json')}
assert all(sha(OUT/p)==r['sha256'] for p,r in own.items())
assert all(sha(p)==r['sha256'] for p,r in external.items())
save(OUT/'EVIDENCE_MANIFEST.json',manifest)
manifest_sha=sha(OUT/'EVIDENCE_MANIFEST.json')
save(OUT/'EVIDENCE_MANIFEST_SHA256.json',{'manifest':str(OUT/'EVIDENCE_MANIFEST.json'),'sha256':manifest_sha,'status':'PASS','evidence_file_count':len(own),'external_file_count':len(external),'verification':'All sealed file SHA256 verified immediately before manifest publication; manifest self SHA verified after publication'})
assert sha(OUT/'EVIDENCE_MANIFEST.json')==manifest_sha
assert all(sha(OUT/p)==r['sha256'] for p,r in own.items())
assert all(sha(p)==r['sha256'] for p,r in external.items())
print(json.dumps({'status':'PASS','manifest_sha256':manifest_sha,'evidence_files':len(own),'external_files':len(external),'READY_FOR_WORK_REVIEW':'YES'},ensure_ascii=False),flush=True)
