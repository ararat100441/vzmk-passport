from common import *
def copy_file(source,target):
    source=Path(source);target=Path(target);assert source.is_file() and not source.is_symlink() and not source.is_junction()
    target.parent.mkdir(parents=True,exist_ok=True);assert not target.exists();shutil.copyfile(source,target);assert sha(source)==sha(target)
import shutil
jobs=read(OUT/'OFFICE_VALIDATION_FINAL.json')['jobs'];copies=[]
for job in jobs:
    folder=Path(job['supervisor']).parent
    assert folder.is_relative_to(Path(SETUP['base'])/'VZMKPassport/diagnostics')
    for p in folder.rglob('*'):
        assert not p.is_symlink() and not p.is_junction()
        if p.is_file():
            target=OUT/'office_jobs'/folder.name/p.relative_to(folder);copy_file(p,target);copies.append({'original':str(p),'copy':str(target),'sha256':sha(p)})
for label in ('FIRST','REPEAT'):
    pairs=read(OUT/('EXE_PAIRS_'+label+'.json'))
    for n,g in pairs['groups'].items():
        for p in Path(g['path']).iterdir():
            if p.is_file():copy_file(p,OUT/'review_workbooks'/label/('AKZ'+n)/p.name)
full=read(OUT/'EXE_FULL_GENERATE.json');manifest=read(full['manifest']);copy_file(full['manifest'],OUT/'full_bundle/MANIFEST.json')
seen=set()
for r in manifest['records']:
    p=Path(r['path'])
    if str(p) in seen:continue
    seen.add(str(p));relative=(Path('AKZ'+str(r['variant']))/p.name) if 'variant' in r else Path(p.name)
    copy_file(p,OUT/'full_bundle'/relative)
native=read(OUT/'EXE_TESTS_FINAL.json')['cases'][-1];evidence=read(native['result'])['evidence'];copy_file(evidence,OUT/'full_bundle/NATIVE_VALIDATION.json')
save(OUT/'EVIDENCE_COPIES.json',{'status':'PASS','office_job_files':copies,'review_pairs':16,'full_physical_files':len(seen),'copy_verification':'SHA256_EQUAL'})
print('Evidence copies:',len(copies),'Office job files; 16 review XLSX; full bundle',len(seen),'files',flush=True)
