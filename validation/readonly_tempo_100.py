from pathlib import Path
import json,os,subprocess,sys
root=Path(__file__).resolve().parents[1]
out=root/'validation/production_100/tempo';out.mkdir(exist_ok=False)
base=Path(r'C:\Users\arara\Desktop\ВЗМК рабочий стол\Работа')
request={'registry_path':str(base/'ВЗМК (2).xlsx'),'project':'ТЭМПО 1','number':'1/1/2026',
    'folder':str(base/'Документы/ТЭМПО/1 выдача/Паспорта/1 от 02.09.2026')}
(out/'request.json').write_text(json.dumps(request,ensure_ascii=False),encoding='utf8')
env=dict(os.environ);env['LOCALAPPDATA']=str(out/'local')
p=subprocess.Popen([sys.executable,'-B',str(root/'production.py'),'--task','inspect-existing','--request',str(out/'request.json'),'--result',str(out/'RESULT.json')],env=env)
p.wait(timeout=600)
print((out/'RESULT.json').read_text(encoding='utf8'))
assert p.returncode==0
