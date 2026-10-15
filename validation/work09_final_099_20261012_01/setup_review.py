"""New independent review roots and protocol copies; no product mutation."""
import os,sys,json,hashlib,shutil,concurrent.futures
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def sha(p):
    h=hashlib.sha256()
    with Path(p).open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
    return h.hexdigest()
def save(p,data):
    with Path(p).open('x',encoding='utf8') as f:json.dump(data,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
PRIOR=ROOT/'validation/stage9_resume_099_20261012_01';FIX=ROOT/'validation/stage9_fix_098_20261012_01'
tag='W09F099_'+uuid4().hex[:8];parent=Path.home()/tag;base=Path(os.environ['LOCALAPPDATA'])/tag
parent.mkdir();base.mkdir();(base/'tmp').mkdir();(base/'source_tests').mkdir()
setup={'parent':str(parent),'base':str(base),'temp':str(base/'tmp'),'source_test_parent':str(base/'source_tests'),
       'source':r'C:\Users\arara\AppData\Local\VZMKStage9_092_9c0a3396\source',
       'vom':r'C:\Users\arara\AppData\Local\VZMKStage9_092_9c0a3396\source\TEST_VOM.xls','prior_evidence':str(FIX)}
save(OUT/'SETUP.json',setup)
provenance=[]
for name in ('common','verify','source_targeted','window_close_guard','extended_guard','e2e','negative','lifecycle','normal_overlap','overlap_crash_cleanup','foreign_pumped','ownership_stop','diagnostic_collision','repeat_native','semantic_audit','receipts'):
    p=PRIOR/(name+'.py');s=p.read_text(encoding='utf8')
    if name=='common':
        s=s.replace('Codex self-test','Independent Work 09 final review')
        s=s.replace("EXPECTED=read(Path(SETUP['prior_evidence'])/'NEW_EXE_BINDING.json')['exe_sha256']","EXPECTED='21ff2f7cc0d4fca746094c46d3594ae0b1dbf991c40266195e22c1030c0ccadc'")
    if name in ('verify','source_targeted'):
        s=s.replace("'gui':'NOT_RUN_USER_STOP'","'gui':'PENDING_INDEPENDENT_GUI'")
        s=s.replace('remains NOT_RUN_USER_STOP','is verified separately in this independent GUI run')
    (OUT/(name+'.py')).write_text(s,encoding='utf8')
    provenance.append({'original':str(p),'original_sha256':sha(p),'copy':str(OUT/(name+'.py')),'copy_sha256':sha(OUT/(name+'.py'))})
save(OUT/'CONTROLLER_PROVENANCE.json',{'scope':'New independent execution; historical self-check results are readback only','adapted':provenance})
exe=ROOT/'stage9_dist_099/VZMK-Passport-0.9.9/VZMK-Passport-0.9.9.exe';assert sha(exe)=='21ff2f7cc0d4fca746094c46d3594ae0b1dbf991c40266195e22c1030c0ccadc'
save(OUT/'EXE_INITIAL_SHA.json',{'exe':str(exe),'sha256':sha(exe)})
paths=[p for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts]
print('Immutable baseline discovered',len(paths),flush=True)
with concurrent.futures.ThreadPoolExecutor(10) as pool:files=dict(pool.map(lambda p:(str(p),sha(p)),paths))
save(OUT/'PROTECTED_BEFORE.json',{'files':files,'production_registry_SHA_gate':False})
print('Immutable baseline sealed',len(files),flush=True)
