"""Fresh authorized DEV fixture. Preserve the previous STOP and 0.9.9 evidence."""
import sys,os,json,hashlib,subprocess
from pathlib import Path
from uuid import uuid4
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
before=read(ROOT/'validation/akz_auto_20261013_01/ACCEPTED_BEFORE.json');assert all(sha(p)==h for p,h in before.items())
save(OUT/'ACCEPTED_BEFORE.json',before)
token=uuid4().hex[:8];parent=Path.home()/('AA10_'+token);parent.mkdir();base=parent/'local';(base/'tmp').mkdir(parents=True)
source=Path(r'C:\Users\arara\AppData\Local\VZMKStage9_092_9c0a3396\source')
exe=ROOT/'stage9_dist_099/VZMK-Passport-0.9.9/VZMK-Passport-0.9.9.exe';env=os.environ.copy();env['LOCALAPPDATA']=str(base);env['TEMP']=env['TMP']=str(base/'tmp');env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
result=OUT/'IMPORT_099.json';cmd=[str(exe),'--task','import','--source',str(source),'--parent',str(parent),'--result',str(result)]
proc=subprocess.run(cmd,env=env,timeout=100,capture_output=True);assert proc.returncode==0,proc.stderr
workspace=Path(read(result)['workspace'])
save(OUT/'SETUP.json',{'parent':str(parent),'base':str(base),'temp':str(base/'tmp'),'workspace':str(workspace),'source':str(source),'vom':str(source/'TEST_VOM.xls'),'exe_099':str(exe),'exe_099_sha256':sha(exe),'gui_permission_for_new_version':False})
save(OUT/'USER_CLARIFICATION.json',{'source':'Direct user reply 2026-10-13','system_samples_1_2_3_authorized':True,'foreign_requisites':'replace verified current project','historical_control':'never transfer','assignments':'catalog № АКЗ; variant is not layer count','formula':'gamma=ln(H/100)+17.62*T/(243.12+T); Td=243.12*gamma/(17.62-gamma)','formula_scope':'water approximation; TEST values not actual measurements','stage1_8_changes':False})
save(OUT/'SEED_BEFORE.json',{'source':str(source),'files':{str(p):sha(p) for p in source.rglob('*') if p.is_file()}})
print(json.dumps(read(OUT/'SETUP.json'),ensure_ascii=False,indent=2))
