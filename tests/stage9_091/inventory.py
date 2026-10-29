from common import *
import json,platform,importlib.metadata,difflib,marshal,types,hashlib
from stage8_closure.storage import digest
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
from PyInstaller.archive.readers import CArchiveReader
label=sys.argv[1] if len(sys.argv)>1 else 'FINAL'
before=read(OUT/'PROTECTED_BEFORE.json')
bad=[p for p,h in before.items() if not Path(p).is_file() or sha(p)!=h]
save(OUT/('SOURCE_INTEGRITY_'+label+'.json'),{'status':'FAIL' if bad else 'PASS','checked':len(before),'changed_or_missing':bad})
assert not bad,bad
old=ROOT/'stage9_dist/VZMK-Passport-0.9';binding=read(OUT/'OLD_09_BUILD_BEFORE.json')
old_files={str(p.relative_to(old)):sha(p) for p in old.rglob('*') if p.is_file()}
assert old_files==binding['files'],'OLD_09_BUILD_CHANGED'
save(OUT/('OLD_09_INTEGRITY_'+label+'.json'),{'status':'PASS','files':len(old_files),'exe_sha256':binding['exe_sha256'],'aggregate_sha256':digest(old_files)})
files={str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}
documents=[p for p in files if Path(p).suffix.lower() in ('.xlsx','.xlsm','.xls','.docx')]
import docx
default=Path(docx.__file__).parent/'templates/default.docx'
assert documents==['_internal\\docx\\templates\\default.docx'] and files[documents[0]]==sha(default),documents
save(OUT/('BUILD_SHA256_'+label+'.json'),{'version':'0.9.1','exe':str(EXE),'exe_sha256':sha(EXE),'file_count':len(files),'files':files,'aggregate_sha256':digest(files),'user_documents_embedded':False})
code=[]
for name in ('passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure','stage9_app'):code+=list((ROOT/name).glob('*.py'))
code += [ROOT/'stage9.py',ROOT/'stage9.spec',ROOT/'stage9_091.spec',ROOT/'stage8.py',ROOT/'stage7_quality/render_emf.ps1',ROOT/'STAGE9_DEV.md']
code += list((ROOT/'tests/stage9').rglob('*.py'))+list((ROOT/'tests/stage9_091').rglob('*.py'))
source={str(p.relative_to(ROOT)):sha(p) for p in sorted(set(code))}
save(OUT/('SOURCE_SHA256_'+label+'.json'),{'files':source,'aggregate_sha256':digest(source)})
archive=CArchiveReader(str(EXE));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalize(code):
    return code.replace(co_filename='<SOURCE>',co_consts=tuple(normalize(c) if isinstance(c,types.CodeType) else c for c in code.co_consts))
compiled=[]
runtime_code=[ROOT/'stage9.py']
for package in ('passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure','stage9_app'):
    runtime_code+=list((ROOT/package).glob('*.py'))
for p in runtime_code:
    relative=p.relative_to(ROOT);name='stage9' if p.name=='stage9.py' else str(relative.with_suffix('')).replace('\\','.')
    if name.endswith('.__init__'):name=name[:-9]
    present=name=='stage9' or name in pyz.toc
    if name=='stage9' or name.startswith('stage9_app'):assert present,('STAGE9_MODULE_MISSING',name)
    if present:
        actual=marshal.loads(archive.extract('stage9')) if name=='stage9' else pyz.extract(name)
        expected=compile(p.read_text(encoding='utf8'),actual.co_filename,'exec',optimize=0)
        assert normalize(actual)==normalize(expected),('COMPILED_SOURCE_MISMATCH',str(p))
    copied=EXE.parent/'_internal'/relative;assert sha(copied)==sha(p),('RAW_SOURCE_MISMATCH',str(p))
    compiled.append({'path':str(relative),'source_sha256':sha(p),'raw_copy_sha256':sha(copied),'compiled_equal':True if present else None,'included_in_pyz':present})
save(OUT/('COMPILED_SOURCE_'+label+'.json'),{'status':'PASS','exe_sha256':sha(EXE),'modules':compiled})
versions={name:importlib.metadata.version(name) for name in ('pyinstaller','pyinstaller-hooks-contrib','pywin32','openpyxl','python-docx','lxml','Pillow','olefile','defusedxml')}
save(OUT/('DEPENDENCIES_'+label+'.json'),{'python':sys.version,'packages':versions,'runtime':'CPython 3.12 / Windows x64 / onedir','system_dependencies':['Microsoft Word COM','Microsoft Excel COM','Windows PowerShell 5.1'],'network':'no_network guards on parent, worker and helper; Office background network not independently blocked'})
diff=[]
for p in [ROOT/'stage9.py',*(ROOT/'stage9_app').glob('*.py')]:
    oldp=old/'_internal'/p.relative_to(ROOT)
    oldtext=oldp.read_text(encoding='utf8') if oldp.exists() else ''
    diff.extend(difflib.unified_diff(oldtext.splitlines(True),p.read_text(encoding='utf8').splitlines(True),fromfile='0.9/'+str(p.relative_to(ROOT)),tofile='0.9.1/'+str(p.relative_to(ROOT))))
(OUT/('STAGE9_DIFF_'+label+'.patch')).write_text(''.join(diff),encoding='utf8')
print(json.dumps({'protected':len(before),'old_build_files':len(old_files),'exe_sha256':sha(EXE),'runtime_raw_modules':len(compiled),'compiled_modules':sum(r['included_in_pyz'] for r in compiled),'source_sha256':digest(source)},indent=2))
