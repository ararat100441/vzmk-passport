import os,sys,json,subprocess,time
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
EXE=ROOT/'stage9_dist_091/VZMK-Passport-0.9.1/VZMK-Passport-0.9.1.exe'
work=json.loads(next(OUT.glob('lifecycle_*/import_*.json')).read_text(encoding='utf8'))['workspace']
env=os.environ.copy();env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
results=[]
for label,local in [('existing_profile',os.environ['LOCALAPPDATA']),('isolated_repeat',json.loads((OUT/'RUNTIME_ENV.json').read_text())['LOCALAPPDATA'])]:
    folder=OUT/('probe_'+label);folder.mkdir();target=folder/'RESULT.json';env['LOCALAPPDATA']=local
    command=[str(EXE),'--task','office-probe','--workspace',work,'--probe-mode','normal','--result',str(target)]
    start=time.monotonic();p=subprocess.run(command,env=env,cwd=folder,timeout=360,capture_output=True)
    d=json.loads(target.read_text(encoding='utf8'));s=json.loads(Path(d['office_supervisor']).read_text(encoding='utf8'))
    data={'label':label,'LOCALAPPDATA':local,'command':command,'cwd':str(folder),'code':p.returncode,'seconds':time.monotonic()-start,'result':d,'supervisor':s}
    (folder/'COMPARISON.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8');results.append(data)
    print(label,p.returncode,d.get('error'),'orphan',s['owned_orphans'],flush=True)
(OUT/'PROBE_COMPARISON.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf8')
