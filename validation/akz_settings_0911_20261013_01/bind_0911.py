from common import *
import marshal,types
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
from PyInstaller.archive.readers import CArchiveReader
archive=CArchiveReader(str(EXE));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalize(c):return c.replace(co_filename='<SOURCE>',co_consts=tuple(normalize(v) if isinstance(v,types.CodeType) else v for v in c.co_consts))
rows=[]
sources=read(OUT/'SOURCE_AT_BUILD.json')
for path,h in sources.items():
    p=Path(path);assert sha(p)==h
    if p.suffix!='.py' or p.is_relative_to(ROOT/'tests'):continue
    rel=p.relative_to(ROOT);name=str(rel.with_suffix('')).replace('\\','.')
    if name.endswith('.__init__'):name=name[:-9]
    present=name=='stage9' or name in pyz.toc
    if name=='stage9' or name.startswith('stage9_app'):assert present,name
    if present:
        actual=marshal.loads(archive.extract('stage9')) if name=='stage9' else pyz.extract(name)
        assert normalize(actual)==normalize(compile(p.read_text(encoding='utf8'),actual.co_filename,'exec',optimize=0)),path
    assert sha(EXE.parent/'_internal'/rel)==h,path
    rows.append({'path':path,'sha256':h,'compiled_equal':present,'raw_equal':True})
files={str(p.relative_to(EXE.parent)):sha(p) for p in sorted(EXE.parent.rglob('*')) if p.is_file()}
assert not (EXE.parent/'_internal/sitecustomize.py').exists()
save(OUT/'BUILD_MANIFEST.json',{'status':'PASS','version':'0.9.11','exe':str(EXE),'exe_sha256':sha(EXE),'files':files,'file_count':len(files),'onedir_digest':hashlib.sha256(json.dumps(files,sort_keys=True,separators=(',',':')).encode()).hexdigest(),'source_modules':rows,'source_at_build':str(OUT/'SOURCE_AT_BUILD.json'),'build_environment':str(OUT/'BUILD_ENVIRONMENT.json')})
print('BIND PASS',len(files),'files;',len(rows),'raw;',sum(x['compiled_equal'] for x in rows),'compiled',flush=True)
