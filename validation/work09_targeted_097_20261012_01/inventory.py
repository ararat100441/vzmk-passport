# Adapted independent run controller; original SHA a95f6be9d247f7f51e2dc985781299ce41c809514772a44998ed5567ed4c5b5e
from common import *
import importlib.metadata,difflib,marshal,types,json
from stage8_closure.storage import digest
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
from PyInstaller.archive.readers import CArchiveReader
label=sys.argv[1] if len(sys.argv)>1 else 'FINAL'
accepted=read(OUT/'ACCEPTED_CODE_BEFORE.json');bad=[p for p,h in accepted.items() if not Path(p).is_file() or sha(p)!=h];assert not bad,bad
old_checks=[]
for binding_path,folder in ((OUT/'BUILD_091_BEFORE.json',ROOT/'stage9_dist_091/VZMK-Passport-0.9.1'),(ROOT/'validation/stage9_091_20261012_01/OLD_09_BUILD_BEFORE.json',ROOT/'stage9_dist/VZMK-Passport-0.9')):
    binding=read(binding_path);old_exe=folder/(folder.name+'.exe')
    files={str(p.relative_to(folder)):sha(p) for p in folder.rglob('*') if p.is_file()};assert files==binding['files'],str(folder)
    assert sha(old_exe)==binding['exe_sha256']
    old_checks.append({'exe':str(old_exe),'files':len(files),'exe_sha256':sha(old_exe),'aggregate_sha256':digest(files)})
save(OUT/('SOURCE_INTEGRITY_'+label+'.json'),{'status':'PASS','accepted_code_files':len(accepted),'changed':bad,'previous_builds':old_checks,'working_registry_sha_gate':False})
files={str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}
documents=[p for p in files if Path(p).suffix.lower() in ('.xlsx','.xlsm','.xls','.docx')]
import docx
assert documents==['_internal\\docx\\templates\\default.docx'] and files[documents[0]]==sha(Path(docx.__file__).parent/'templates/default.docx')
save(OUT/('BUILD_SHA256_'+label+'.json'),{'version':'0.9.7','exe':str(EXE),'exe_sha256':sha(EXE),'file_count':len(files),'files':files,'aggregate_sha256':digest(files),'user_documents_embedded':False})
code=[];runtime=[ROOT/'stage9.py',ROOT/'stage8.py']
for package in ('passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure','stage9_app'):runtime+=list((ROOT/package).glob('*.py'))
code+=runtime+[ROOT/'stage7_quality/render_emf.ps1',ROOT/'stage9.spec',ROOT/'stage9_091.spec',ROOT/'stage9_092.spec',ROOT/'stage9_093.spec',ROOT/'stage9_094.spec',ROOT/'stage9_095.spec',ROOT/'stage9_096.spec',ROOT/'stage9_097.spec']
for directory in ('stage9','stage9_091','stage9_092'):code+=list((ROOT/'tests'/directory).rglob('*.py'))
source={str(p.relative_to(ROOT)):sha(p) for p in sorted(set(code))};save(OUT/('SOURCE_SHA256_'+label+'.json'),{'files':source,'aggregate_sha256':digest(source)})
archive=CArchiveReader(str(EXE));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalize(c):return c.replace(co_filename='<SOURCE>',co_consts=tuple(normalize(v) if isinstance(v,types.CodeType) else v for v in c.co_consts))
compiled=[]
for p in runtime:
    rel=p.relative_to(ROOT);name=str(rel.with_suffix('')).replace('\\','.')
    if name.endswith('.__init__'):name=name[:-9]
    present=name=='stage9' or name in pyz.toc
    if name=='stage9' or name.startswith('stage9_app'):assert present,('STAGE9_MODULE_MISSING',name)
    if present:
        actual=marshal.loads(archive.extract('stage9')) if name=='stage9' else pyz.extract(name)
        assert normalize(actual)==normalize(compile(p.read_text(encoding='utf8'),actual.co_filename,'exec',optimize=0)),('COMPILED_SOURCE_MISMATCH',str(p))
    copied=EXE.parent/'_internal'/rel;assert sha(copied)==sha(p),('RAW_SOURCE_MISMATCH',str(p))
    compiled.append({'path':str(rel),'sha256':sha(p),'compiled_equal':True if present else None})
save(OUT/('COMPILED_SOURCE_'+label+'.json'),{'status':'PASS','modules':compiled,'exe_sha256':sha(EXE)})
versions={n:importlib.metadata.version(n) for n in ('pyinstaller','pyinstaller-hooks-contrib','pywin32','openpyxl','python-docx','lxml','Pillow','olefile','defusedxml')}
save(OUT/('DEPENDENCIES_'+label+'.json'),{'python':sys.version,'packages':versions,'system_dependencies':['Microsoft Word COM','Microsoft Excel COM','Windows PowerShell 5.1'],'network':'no_network guards active; physical Office network disconnection NOT_RUN'})
diff=[];changed=[]
for p in [ROOT/'stage9.py',*(ROOT/'stage9_app').glob('*.py')]:
    oldp=ROOT/'validation/stage9_092_20261012_01/source_before'/p.relative_to(ROOT);oldtext=oldp.read_text(encoding='utf8') if oldp.exists() else '';newtext=p.read_text(encoding='utf8')
    if oldtext!=newtext:changed.append(str(p.relative_to(ROOT)))
    diff+=difflib.unified_diff(oldtext.splitlines(True),newtext.splitlines(True),fromfile='0.9.1/'+str(p.relative_to(ROOT)),tofile='0.9.7/'+str(p.relative_to(ROOT)))
(OUT/('STAGE9_DIFF_'+label+'.patch')).write_text(''.join(diff),encoding='utf8')
save(OUT/('CHANGED_PRODUCT_FILES_'+label+'.json'),changed)
print(json.dumps({'accepted':len(accepted),'old_builds':old_checks,'exe':sha(EXE),'source':digest(source),'compiled':sum(r['compiled_equal'] is True for r in compiled),'raw':len(compiled)},indent=2))
