import os,sys,json,hashlib,shutil,subprocess,concurrent.futures
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,data):
    with Path(p).open('x',encoding='utf8') as f:json.dump(data,f,ensure_ascii=False,indent=2);f.flush();os.fsync(f.fileno())
old=ROOT/'validation/work09_targeted_097_20261012_01';prior=json.loads((old/'SETUP.json').read_text(encoding='utf8'))
setup=json.loads((OUT/'SETUP.json').read_text(encoding='utf8'));base=Path(setup['base']);parent=Path(setup['parent'])
paths=[p for p in ROOT.rglob('*') if p.is_file() and not p.is_relative_to(OUT) and '.codebase-index' not in p.parts]
print('Protected files discovered',len(paths),flush=True)
with concurrent.futures.ThreadPoolExecutor(10) as pool:before=dict(pool.map(lambda p:(str(p),sha(p)),paths))
save(OUT/'PROTECTED_BEFORE.json',{'files':before,'production_registry_SHA_gate':False})
for p in [ROOT/'stage9.py',*(ROOT/'stage9_app').glob('*.py')]:
    dest=OUT/'source_before'/p.relative_to(ROOT);dest.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(p,dest)
save(OUT/'DEV_SOURCE_BEFORE.json',{'source':setup['source'],'files':{str(p.relative_to(setup['source'])):sha(p) for p in Path(setup['source']).rglob('*') if p.is_file()}})
oldexe=ROOT/'stage9_dist_097/VZMK-Passport-0.9.7/VZMK-Passport-0.9.7.exe'
assert sha(oldexe)=='cd140f337612c462903ecc35bb55173ba10839bb16fca61492c7ea7b65663cd7'
save(OUT/'OLD_EXE_BINDING.json',{'path':str(oldexe),'sha256':sha(oldexe),'onedir':{str(p.relative_to(oldexe.parent)):sha(p) for p in oldexe.parent.rglob('*') if p.is_file()}})
# Reproduce using the previous sealed TEST bundle, read-only. Only new settings
# and viewer diagnostics may be written; generation/closure are not invoked.
print(json.dumps({'status':'BASELINE_READY','files':len(before),'old_exe_sha256':sha(oldexe)}),flush=True)
sys.exit(0)
prior_e2e=json.loads((old/'EXE_E2E_FINAL.json').read_text(encoding='utf8'))
cfg=base/'VZMKPassport';cfg.mkdir();save(cfg/'settings.json',{'workspace':prior_e2e['workspace'],'version':'0.9.7'})
env=os.environ.copy();env['LOCALAPPDATA']=str(base);env['TEMP']=env['TMP']=str(base/'tmp');env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
import win32process
p=subprocess.Popen([str(oldexe)],env=env,creationflags=subprocess.CREATE_NO_WINDOW,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
save(OUT/'OLD_GUI_LAUNCH.json',{'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'exe':str(oldexe),'sha256':sha(oldexe),'read_only_workspace':prior_e2e['workspace'],'manifest':prior_e2e['manifests'][0],'LOCALAPPDATA':str(base)})
print(json.dumps({'protected':len(before),'old_gui_pid':p.pid,'manifest':prior_e2e['manifests'][0],'base':str(base)},ensure_ascii=False),flush=True)
