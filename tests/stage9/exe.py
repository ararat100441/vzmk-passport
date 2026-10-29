"""Real frozen EXE E2E with no Python executable on PATH and network denied."""
import os,sys,json,subprocess,hashlib,traceback
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage8_closure.storage import read,sha
OUT=ROOT/'validation/stage9_20261012_01';RUN=OUT/('exe_'+uuid4().hex[:8]);RUN.mkdir()
EXE=ROOT/'stage9_dist/VZMK-Passport-0.9/VZMK-Passport-0.9.exe'
env=os.environ.copy();system_root=os.environ['SystemRoot'];env['PATH']=system_root+'\\System32;'+system_root;env.pop('PYTHONHOME',None);env.pop('PYTHONPATH',None);env['PYTHONUTF8']='1'
results=[]
def task(name,**args):
    path=RUN/(name+'_'+uuid4().hex[:6]+'.json');cmd=[str(EXE),'--task',name,'--result',str(path)]
    for key,value in args.items():
        cmd.append('--'+key.replace('_','-'))
        if value is not True:cmd.append(str(value))
    proc=subprocess.run(cmd,env=env,timeout=300,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0),capture_output=True)
    data=read(path);results.append({'task':name,'returncode':proc.returncode,'result':str(path),'status':data['status']})
    save(RUN/('step_'+str(len(results))+'.json'),results[-1]);print(name,data['status'],flush=True)
    assert proc.returncode==0 and data['status']=='PASS',data
    return data
try:
    work=task('import',source=ROOT,parent=Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/workspaces')['workspace']
    smoke=task('smoke',workspace=work);assert smoke['frozen'] and smoke['python_in_path'] is None
    vom=ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
    preview=task('check',workspace=work,vom=vom,passport='9/2026',variants='0')
    d=preview['data'];assert d['visible_rows']==104 and d['hidden_excluded']==397 and d['quantity']=='288'
    generated=task('generate',workspace=work,preview=preview['preview'],confirm_test=True)
    task('verify',workspace=work,manifest=generated['manifest'])
    fixture=task('fixture',workspace=work)
    closed=task('close-preview',workspace=work,sandbox=fixture['sandbox'],request=fixture['request'])
    p=RUN/'CLOSE_PREVIEW.json';save(p,closed['plan'])
    from stage8_closure.core import confirmation
    task('close',workspace=work,preview=p,confirm=confirmation(closed['plan']))
    again=task('close',workspace=work,preview=p,confirm=confirmation(closed['plan']))
    assert again['operation_status']=='ALREADY_CLOSED_TEST'
    task('recover',workspace=work,sandbox=fixture['sandbox'])
    torn=task('fixture',workspace=work)
    rp=task('close-preview',workspace=work,sandbox=torn['sandbox'],request=torn['request'])
    crash_plan=RUN/'CRASH_PLAN.json';save(crash_plan,rp['plan'])
    crashed=subprocess.run([sys.executable,'-B',str(ROOT/'tests/stage9/frozen_worker.py'),str(EXE.parent/'_internal'),work,str(crash_plan)],timeout=150,capture_output=True)
    assert crashed.returncode==73,crashed.stderr.decode('utf8','replace')
    restored=task('restore-preview',workspace=work,sandbox=torn['sandbox'])
    restore_plan=RUN/'RESTORE_PLAN.json';save(restore_plan,restored['plan'])
    from stage8_closure.recovery import confirmation as restore_confirmation
    restored_result=task('restore',workspace=work,preview=restore_plan,confirm=restore_confirmation(restored['plan']))
    assert restored_result['operation_status']=='RECOVERED_ABORTED' and restored_result['decrements']==0
    after_restore=task('close-preview',workspace=work,sandbox=torn['sandbox'],request=torn['request'])
    assert after_restore['plan']['payload']['status']=='READY_TEST'
    task('smoke',workspace=work)
    new=task('check',workspace=work,vom=vom,passport='10/2026',variants='0',new=True)
    new_output=task('generate',workspace=work,preview=new['preview'],confirm_test=True)
    task('verify',workspace=work,manifest=new_output['manifest'])
    task('verify',workspace=work,manifest=generated['manifest'])
    save(RUN/'RESULT.json',{'status':'PASS','exe_sha256':sha(EXE),'workspace':work,'manifest':generated['manifest'],'new_passport_manifest':new_output['manifest'],'steps':results,'network_policy':'Python sockets/DNS denied in entry point','python_in_path':None})
    save(OUT/'EXE_FINAL_RESULT.json',{'status':'PASS','evidence':str(RUN/'RESULT.json'),'exe_sha256':sha(EXE),'manifest':generated['manifest'],'workspace':work})
except Exception as exc:
    save(RUN/'FAILURE.json',{'status':'FAIL','error':str(exc),'traceback':traceback.format_exc(),'steps':results,'exe_sha256':sha(EXE)})
    raise
