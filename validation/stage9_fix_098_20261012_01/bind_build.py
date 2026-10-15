import os,sys,json,hashlib,marshal,types,difflib,shutil
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
from PyInstaller.archive.readers import CArchiveReader
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,v):
    with Path(p).open('x',encoding='utf8') as f:json.dump(v,f,ensure_ascii=False,indent=2);f.flush();os.fsync(f.fileno())
exe=ROOT/'stage9_dist_099/VZMK-Passport-0.9.9/VZMK-Passport-0.9.9.exe'
files={str(p.relative_to(exe.parent)):sha(p) for p in exe.parent.rglob('*') if p.is_file()}
save(OUT/'NEW_EXE_BINDING.json',{'version':'0.9.9','exe':str(exe),'exe_sha256':sha(exe),'files':files,'file_count':len(files)})
runtime=[ROOT/'stage9.py',ROOT/'stage8.py']
for package in ('stage9_app','passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure'):runtime+=list((ROOT/package).glob('*.py'))
archive=CArchiveReader(str(exe));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalize(c):return c.replace(co_filename='<SOURCE>',co_consts=tuple(normalize(v) if isinstance(v,types.CodeType) else v for v in c.co_consts))
rows=[]
for p in runtime:
    rel=p.relative_to(ROOT);name=str(rel.with_suffix('')).replace('\\','.')
    if name.endswith('.__init__'):name=name[:-9]
    present=name=='stage9' or name in pyz.toc
    if name=='stage9' or name.startswith('stage9_app'):assert present,name
    if present:
        actual=marshal.loads(archive.extract('stage9')) if name=='stage9' else pyz.extract(name)
        assert normalize(actual)==normalize(compile(p.read_text(encoding='utf8'),actual.co_filename,'exec',optimize=0)),str(p)
    assert sha(exe.parent/'_internal'/rel)==sha(p),str(p)
    rows.append({'path':str(p),'sha256':sha(p),'compiled_equal':present,'raw_equal':True})
save(OUT/'COMPILED_SOURCE_BEFORE_TESTS.json',{'status':'PASS','exe_sha256':sha(exe),'modules':rows})
sources={str(p):sha(p) for p in runtime+[ROOT/'stage9_098.spec',ROOT/'stage9_099.spec',ROOT/'tests/stage9_098/test_cleanup.py']}
save(OUT/'SOURCE_BEFORE_TESTS.json',sources)
diff=[];changed=[]
for p in (ROOT/'stage9_app').glob('*.py'):
    before=OUT/'source_before/stage9_app'/p.name;a=before.read_text(encoding='utf8') if before.exists() else '';b=p.read_text(encoding='utf8')
    if a!=b:changed.append({'path':str(p),'before_sha256':sha(before) if before.exists() else None,'after_sha256':sha(p)})
    diff+=difflib.unified_diff(a.splitlines(True),b.splitlines(True),fromfile='0.9.7/stage9_app/'+p.name,tofile='0.9.9/stage9_app/'+p.name)
save(OUT/'CHANGED_PRODUCT_FILES.json',changed)
(OUT/'EXACT_CODE_DIFF.patch').write_text(''.join(diff),encoding='utf8')
setup=json.loads((OUT/'SETUP.json').read_text(encoding='utf8'));setup['base']=str(Path(setup['base'])/'new099');setup['temp']=str(Path(setup['base'])/'tmp')
Path(setup['temp']).mkdir(parents=True);save(OUT/'SETUP_099.json',setup)
for name in ('common.py',):
    p=OUT/name;s=p.read_text(encoding='utf8').replace("OUT/'SETUP.json'","OUT/'SETUP_099.json'");p.write_text(s,encoding='utf8')
print(json.dumps({'status':'PASS','exe_sha256':sha(exe),'files':len(files),'compiled':sum(r['compiled_equal'] for r in rows),'raw':len(rows),'changes':len(changed)}),flush=True)
