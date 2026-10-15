"""Exercise actual frozen COM cleanup three times, never an existing Office PID."""
import sys,os,subprocess,json
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage8_closure.storage import read,sha
OUT=ROOT/'validation/stage9_20261012_01'/('ownership_'+uuid4().hex[:8]);OUT.mkdir()
EXE=ROOT/'stage9_dist/VZMK-Passport-0.9/VZMK-Passport-0.9.exe'
workspace=read(ROOT/'validation/stage9_20261012_01/WORKSPACE.json')['path']
vom=ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
env=os.environ.copy();env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONHOME',None);env.pop('PYTHONPATH',None)
diagnostics=Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/diagnostics'
cases=[]
for n in range(3):
    result=OUT/f'check_{n}.json'
    before=set(diagnostics.glob('office_*.json'))
    p=subprocess.run([str(EXE),'--task','check','--workspace',workspace,'--vom',str(vom),'--passport','9/2026','--variants','0','--result',str(result)],env=env,timeout=180)
    data=read(result);assert p.returncode==0 and data['status']=='PASS',data
    paths=[p for p in set(diagnostics.glob('office_*.json'))-before if read(p).get('schema')=='STAGE9_OFFICE_OWNERSHIP_V1'];assert len(paths)==1,paths
    diagnostic=read(paths[0]);assert diagnostic['owned_orphans']==0
    assert diagnostic['instances'] and all(r['quit_attempted'] and not r['still_running'] and r['exit_code']!=259 for r in diagnostic['instances'])
    cases.append({'status':'PASS','result':str(result),'office_diagnostic':str(paths[0])});print(n,'PASS',flush=True)
save(OUT/'RESULT.json',{'status':'PASS','cases':cases,'exe_sha256':sha(EXE)})
