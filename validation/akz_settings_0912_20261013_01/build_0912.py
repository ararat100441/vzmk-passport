import os,sys,json,hashlib,subprocess,time,importlib.metadata,difflib
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];PREV=ROOT/'validation/akz_settings_0911_20261013_01'
PY=Path(sys.executable);EXE=ROOT/'stage9_dist_0912/VZMK-Passport-0.9.12/VZMK-Passport-0.9.12.exe'
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str)
assert not EXE.parent.exists(),'PRESERVE_ALL_OLD_BUILDS'
accepted=read(PREV/'ACCEPTED_BEFORE.json');assert all(sha(p)==h for p,h in accepted.items())
old=read(PREV/'BUILD_MANIFEST.json');assert all(sha(Path(old['exe']).parent/p)==h for p,h in old['files'].items())
sources=read(PREV/'SOURCE_AT_BUILD.json');changed=[];diff=[]
for name,h in sources.items():
    p=Path(name)
    if sha(p)==h:continue
    rel=p.relative_to(ROOT)
    assert rel.as_posix() in ['stage9_app/akz_settings.py','stage9_app/__init__.py','tests/stage9_akz_auto/test_persistence.py']
    if p.is_relative_to(ROOT/'tests'):
        text=p.read_text(encoding='utf8');start=text.index('    def test_legacy_recovery_windows_writer_crlf');end=text.index('    def test_manifest_cannot_omit_snapshot_binding',start)
        before=(text[:start]+text[end:]).replace('\n','\r\n').encode('utf8')
        if hashlib.sha256(before).hexdigest()!=h:before=(text[:start]+text[end:]).encode('utf8')
    else:before=(Path(old['exe']).parent/'_internal'/rel).read_bytes()
    assert hashlib.sha256(before).hexdigest()==h,(p,'ORIGINAL_BINDING')
    copy=OUT/'before_0911'/rel;copy.parent.mkdir(parents=True,exist_ok=True);copy.write_bytes(before)
    changed.append({'path':str(p),'relative':str(rel),'before_sha256':h,'after_sha256':sha(p),'original_copy':str(copy)})
    diff.extend(difflib.unified_diff(before.decode('utf8').splitlines(True),p.read_text(encoding='utf8').splitlines(True),fromfile='0.9.11/'+rel.as_posix(),tofile='0.9.12/'+rel.as_posix()))
spec=ROOT/'stage9_0912.spec';changed.append({'path':str(spec),'relative':spec.name,'before_sha256':None,'after_sha256':sha(spec),'purpose':'New version name only; 0.9.11 spec preserved'})
save(OUT/'CHANGED_FILES.json',changed);(OUT/'EXACT_CODE_DIFF.patch').write_text(''.join(diff),encoding='utf8')
sources.pop(str(ROOT/'stage9_0911.spec'));sources[str(spec)]=sha(spec)
sources={p:sha(p) for p in sources};save(OUT/'SOURCE_AT_BUILD.json',sources)
save(OUT/'BUILD_ENVIRONMENT.json',{'python':str(PY),'python_version':sys.version,'dependencies':{d.metadata['Name']:d.version for d in importlib.metadata.distributions() if d.metadata['Name']},'build_hook':str(ROOT/'tests/stage9/build_runtime/sitecustomize.py'),'build_hook_sha256':sha(ROOT/'tests/stage9/build_runtime/sitecustomize.py'),'hook_in_product':False})
cmd=[str(PY),'-B','-m','PyInstaller','--noconfirm','--clean','--distpath',str(ROOT/'stage9_dist_0912'),'--workpath',str(ROOT/'stage9_build_0912'),str(spec)]
env=os.environ.copy();env['PYTHONPATH']=str(ROOT/'tests/stage9/build_runtime');env['PYTHONUTF8']='1'
save(OUT/'BUILD_COMMAND.json',{'command':cmd,'PYTHONPATH':env['PYTHONPATH'],'started':time.time()})
with (OUT/'BUILD_0912.log').open('xb') as f:r=subprocess.run(cmd,env=env,stdout=f,stderr=subprocess.STDOUT,timeout=900)
assert r.returncode==0;assert all(sha(p)==h for p,h in sources.items())
print('BUILD 0.9.12 PASS',sha(EXE),flush=True)
