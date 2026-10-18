from pathlib import Path
import os,subprocess,json,sys
root=Path(__file__).resolve().parents[1];out=root/'validation/production_100/gui'
out.mkdir(exist_ok=False)
cfg=json.loads(Path(r'C:\Temp\vzmk100-source-final\RESULT.json').read_text(encoding='utf8'))['cfg']
local=out/'local';folder=local/'VZMKPassportProduction';folder.mkdir(parents=True)
(folder/'settings.json').write_text(json.dumps(cfg,ensure_ascii=False),encoding='utf8')
env=dict(os.environ);env['LOCALAPPDATA']=str(local)
log=(out/'GUI.log').open('w')
p=subprocess.Popen([sys.executable,'-B',str(root/'production.py')],env=env,stdout=log,stderr=log)
(out/'PROCESS.json').write_text(json.dumps({'pid':p.pid,'cfg':cfg,'local':str(local)}),encoding='utf8')
print(p.pid)
