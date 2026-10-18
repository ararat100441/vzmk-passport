from pathlib import Path
import os,subprocess,sys,json
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from production_app.storage import save,read
out=ROOT/'validation/production_100/acceptance_sources_final/smoke_exe';out.mkdir(exist_ok=False)
exe=read(ROOT/'validation/production_100/DISTRIBUTION.json')['exe']
save(out/'request.json',{})
env=dict(os.environ);env['LOCALAPPDATA']=str(out/'local');env['PATH']=str(Path(os.environ['SystemRoot'])/'System32')
p=subprocess.Popen([exe,'--task','smoke','--request',str(out/'request.json'),'--result',str(out/'RESULT.json')],env=env)
p.wait(timeout=120);r=read(out/'RESULT.json');assert p.returncode==0 and r['status']=='PASS' and r['version']=='1.0.0-rc1' and r['gui'] and r['frozen'],r
print(json.dumps(r,ensure_ascii=False,indent=2))
