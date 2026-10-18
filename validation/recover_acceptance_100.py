from pathlib import Path
import os,subprocess,sys,json
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from production_app.storage import read,save,sha
base=Path(r'C:\Temp\vzmk100-acceptance-20261018');extra=Path(r'C:\Temp\vzmk100-acceptance-extra-20261018')
accepted=read(base/'RESULT.json');cfg=accepted['cfg']
plans=[read(p)['plan'] for p in (extra/'calls').glob('*/result.json') if read(p).get('data',{}).get('new')]
plan=read(plans[0]);journal=Path(cfg['registry']).parent/'.vzmk-operations'/plan['id']/'journal.json'
out=ROOT/'validation/production_100/acceptance_resume/recover_monitor_collision';out.mkdir(exist_ok=False)
save(out/'request.json',{'journal':str(journal)})
env=dict(os.environ);env['LOCALAPPDATA']=str(base/'local')
p=subprocess.Popen([accepted['exe'],'--task','recover','--request',str(out/'request.json'),'--result',str(out/'RESULT.json')],env=env)
p.wait(timeout=120)
result=read(out/'RESULT.json');print(json.dumps(result,ensure_ascii=False,indent=2));assert p.returncode==0 and result['status']=='PASS'
