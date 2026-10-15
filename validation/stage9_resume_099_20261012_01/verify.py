"""Bind existing EXE, all onedir files, compiled/raw source and accepted code."""
from common import *
import marshal,types,importlib.metadata,collections
sys.path.insert(0,str(ROOT/'tests/stage9/build_runtime'));import sitecustomize
from PyInstaller.archive.readers import CArchiveReader
from stage8_closure.storage import digest
label=sys.argv[1] if len(sys.argv)>1 else 'BEFORE'
prior=Path(SETUP['prior_evidence']);binding=read(prior/'NEW_EXE_BINDING.json')
assert sha(EXE)==binding['exe_sha256']==EXPECTED
actual={str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()}
assert actual==binding['files'] and len(actual)==1117,'ONEDIR_CHANGED_OR_INCOMPLETE'
previous=read(prior/'SOURCE_BEFORE_TESTS.json')
assert all(sha(p)==h for p,h in previous.items()),'SOURCE_CHANGED_SINCE_099_BUILD'
runtime=[ROOT/'stage9.py',ROOT/'stage8.py']
for package in ('stage9_app','passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure'):runtime+=list((ROOT/package).glob('*.py'))
archive=CArchiveReader(str(EXE));pyz=archive.open_embedded_archive('PYZ.pyz')
def normalize(c):return c.replace(co_filename='<SOURCE>',co_consts=tuple(normalize(v) if isinstance(v,types.CodeType) else v for v in c.co_consts))
rows=[]
for p in runtime:
    rel=p.relative_to(ROOT);name=str(rel.with_suffix('')).replace('\\','.')
    if name.endswith('.__init__'):name=name[:-9]
    present=name=='stage9' or name in pyz.toc
    if name=='stage9' or name.startswith('stage9_app'):assert present,name
    if present:
        actual_code=marshal.loads(archive.extract('stage9')) if name=='stage9' else pyz.extract(name)
        assert normalize(actual_code)==normalize(compile(p.read_text(encoding='utf8'),actual_code.co_filename,'exec',optimize=0)),('COMPILED_MISMATCH',name)
    assert sha(EXE.parent/'_internal'/rel)==sha(p),('RAW_MISMATCH',str(rel))
    rows.append({'path':str(p),'sha256':sha(p),'compiled_equal':present,'raw_equal':True})
accepted=read(ROOT/'validation/stage9_092_20261012_01/ACCEPTED_CODE_BEFORE.json')
assert all(sha(p)==h for p,h in accepted.items()),'ACCEPTED_STAGE1_8_CHANGED'
old=read(prior/'OLD_EXE_BINDING.json');oldexe=Path(old['path'])
oldfiles={str(p.relative_to(oldexe.parent)):sha(p) for p in oldexe.parent.rglob('*') if p.is_file()}
assert oldfiles==old['onedir'] and sha(oldexe)==old['sha256'],'OLD_097_CHANGED'
dev=read(prior/'DEV_SOURCE_BEFORE.json');source=Path(dev['source'])
devfiles={str(p.relative_to(source)):sha(p) for p in source.rglob('*') if p.is_file()}
assert devfiles==dev['files'],'SEALED_DEV_SOURCE_CHANGED'
changed=[]
for p in [ROOT/'stage9.py',*(ROOT/'stage9_app').glob('*.py')]:
    oldp=prior/'source_before'/p.relative_to(ROOT)
    if not oldp.exists() or sha(oldp)!=sha(p):changed.append({'path':str(p),'before_sha256':sha(oldp) if oldp.exists() else None,'sha256':sha(p)})
source_files={str(p):sha(p) for p in runtime+[ROOT/'stage9_099.spec',ROOT/'tests/stage9_098/test_cleanup.py']}
save(OUT/('SOURCE_SHA256_'+label+'.json'),{'files':source_files,'aggregate_sha256':digest(source_files)})
save(OUT/('CHANGED_PRODUCT_FILES_'+label+'.json'),changed)
save(OUT/('BUILD_SHA256_'+label+'.json'),{'version':'0.9.9','exe':str(EXE),'exe_sha256':sha(EXE),'files':actual,'file_count':len(actual),'aggregate_sha256':digest(actual),'previous_build_manifest':str(prior/'NEW_EXE_BINDING.json'),'previous_build_manifest_sha256':sha(prior/'NEW_EXE_BINDING.json')})
save(OUT/('COMPILED_SOURCE_'+label+'.json'),{'status':'PASS','exe_sha256':sha(EXE),'modules':rows})
save(OUT/('SOURCE_INTEGRITY_'+label+'.json'),{'status':'PASS','accepted_code_count':len(accepted),'accepted_files':accepted,'old_097_unchanged':True,'old_097_sha256':sha(oldexe),'old_097_files':len(oldfiles),'sealed_dev_source_unchanged':True,'dev_source':str(source),'dev_files':devfiles,'working_registry_SHA_gate':False})
versions={n:importlib.metadata.version(n) for n in ('pyinstaller','pyinstaller-hooks-contrib','pywin32','openpyxl','python-docx','lxml','Pillow','olefile','defusedxml')}
save(OUT/('DEPENDENCIES_'+label+'.json'),{'python':sys.version,'packages':versions,'system':['Word/Excel COM','Windows PowerShell 5.1'],'physical_offline':'NOT_RUN','gui':'NOT_RUN_USER_STOP'})
print(json.dumps({'version':'0.9.9','exe_sha256':sha(EXE),'files':len(actual),'compiled':sum(r['compiled_equal'] for r in rows),'raw':len(rows),'accepted_code':len(accepted),'changed_product_files':len(changed),'source_aggregate':digest(source_files),'build_aggregate':digest(actual)},indent=2))
