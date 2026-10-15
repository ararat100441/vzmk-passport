from common import *
import marshal,types
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
from PyInstaller.archive.readers import CArchiveReader
archive=CArchiveReader(str(EXE));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalize(c):return c.replace(co_filename='<SOURCE>',co_consts=tuple(normalize(v) if isinstance(v,types.CodeType) else v for v in c.co_consts))
rows=[]
for path,h in read(OUT/'SOURCE_BEFORE.json').items():
 p=Path(path);assert sha(p)==h
 if p.suffix!='.py' or p.is_relative_to(ROOT/'tests'):continue
 rel=p.relative_to(ROOT);name=str(rel.with_suffix('')).replace('\\','.')
 if name.endswith('.__init__'):name=name[:-9]
 present=name=='stage9' or name in pyz.toc
 if name=='stage9' or name.startswith('stage9_app'):assert present
 if present:
  actual=marshal.loads(archive.extract('stage9')) if name=='stage9' else pyz.extract(name)
  assert normalize(actual)==normalize(compile(p.read_text(encoding='utf8'),actual.co_filename,'exec',optimize=0)),path
 assert sha(EXE.parent/'_internal'/rel)==h
 rows.append({'path':path,'sha256':h,'compiled_equal':present,'raw_equal':True})
save(OUT/'COMPILED_SOURCE_BINDING.json',{'status':'PASS','modules':rows,'compiled_count':sum(r['compiled_equal'] for r in rows),'raw_count':len(rows),'only_co_filename_normalized':True})
print('Binding PASS',sum(r['compiled_equal'] for r in rows),'compiled',len(rows),'raw',flush=True)
