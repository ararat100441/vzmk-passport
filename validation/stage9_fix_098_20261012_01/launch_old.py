import os,json,subprocess,hashlib,win32process
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
setup=json.loads((OUT/'SETUP.json').read_text(encoding='utf8'));old=ROOT/'validation/work09_targeted_097_20261012_01'
e2e=json.loads((old/'EXE_E2E_FINAL.json').read_text(encoding='utf8'))
cfg=Path(setup['base'])/'VZMKPassport';cfg.mkdir(exist_ok=True)
with (cfg/'settings.json').open('x',encoding='utf8') as f:json.dump({'workspace':e2e['workspace'],'version':'0.9.7'},f)
exe=ROOT/'stage9_dist_097/VZMK-Passport-0.9.7/VZMK-Passport-0.9.7.exe'
env=os.environ.copy();env['LOCALAPPDATA']=setup['base'];env['TEMP']=env['TMP']=setup['temp'];env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
p=subprocess.Popen([str(exe)],env=env,creationflags=subprocess.CREATE_NO_WINDOW,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
record={'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'exe':str(exe),'sha256':hashlib.sha256(exe.read_bytes()).hexdigest(),'read_only_workspace':e2e['workspace'],'manifest':e2e['manifests'][0],'LOCALAPPDATA':setup['base']}
with (OUT/'OLD_GUI_LAUNCH.json').open('x',encoding='utf8') as f:json.dump(record,f,ensure_ascii=False,indent=2)
print(json.dumps(record,ensure_ascii=False),flush=True)
