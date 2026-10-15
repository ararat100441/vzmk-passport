"""Manifest-last handoff; all source/workbook baselines are read-only."""
import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage8_closure.storage import read,sha,digest
OUT=ROOT/'validation/stage9_20261012_01'
before=read(OUT/'PROTECTED_BEFORE.json');changed=[];missing=[]
for path,h in before.items():
    if not Path(path).is_file():missing.append(path)
    elif sha(path)!=h:changed.append(path)
save(OUT/'SOURCE_INTEGRITY.json',{'status':'PASS' if not changed and not missing else 'FAIL','checked':len(before),'changed':changed,'missing':missing})
assert not changed and not missing
code=[]
for name in ('passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure','stage9_app'):
    code+=list((ROOT/name).glob('*.py'))
code += [ROOT/'stage9.py',ROOT/'stage9.spec',ROOT/'stage8.py',ROOT/'stage7_quality/render_emf.ps1',ROOT/'STAGE9_DEV.md']
code += list((ROOT/'tests/stage9').rglob('*.py'))
source={str(p.relative_to(ROOT)):sha(p) for p in code}
save(OUT/'SOURCE_SHA256.json',{'schema':'STAGE9_SOURCE_BINDING_V1','files':source,'aggregate_sha256':digest(source)})
dist=ROOT/'stage9_dist/VZMK-Passport-0.9'
files={str(p.relative_to(dist)):sha(p) for p in dist.rglob('*') if p.is_file()}
assert not any(Path(p).suffix.lower() in ('.xlsx','.xlsm','.xls','.docx') for p in files)
save(OUT/'BUILD_SHA256.json',{'schema':'STAGE9_BUILD_BINDING_V1','version':'0.9','exe':'VZMK-Passport-0.9.exe','exe_sha256':files['VZMK-Passport-0.9.exe'],'file_count':len(files),'files':files,'aggregate_sha256':digest(files),'user_documents_embedded':False})
print('Protected',len(before),'unchanged; build files',len(files),'EXE',files['VZMK-Passport-0.9.exe'])
