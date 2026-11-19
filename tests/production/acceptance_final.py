"""Windows frozen-binary acceptance; production inputs stay read-only."""
import argparse,json,os,subprocess,sys,time
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app.storage import read,save,sha

def main():
    p=argparse.ArgumentParser();p.add_argument('--exe',required=True);p.add_argument('--output',required=True);a=p.parse_args()
    exe=Path(a.exe).resolve();out=Path(a.output).resolve();out.mkdir(parents=True,exist_ok=False)
    results=[];env=dict(os.environ);env['PYTHONIOENCODING']='utf-8'
    def run(label,command,environment=None):
        with (out/(label+'.log')).open('w',encoding='utf8') as log:
            proc=subprocess.Popen(command,cwd=ROOT,env=environment or env,stdout=log,stderr=subprocess.STDOUT)
            try:code=proc.wait(timeout=1200)
            except subprocess.TimeoutExpired:
                save(out/'CONTROLLER_TIMEOUT.json',{'pid':proc.pid,'supervisor_left_alive':True,'label':label});raise
        result={'label':label,'exit_code':code,'log':str(out/(label+'.log'))};results.append(result)
        save(out/'progress.json',results);print(label,code,flush=True)
        if code:raise RuntimeError(label+' failed; see '+str(out/(label+'.log')))
    run('safety',[sys.executable,'-B','tests/production/test_safety.py'])
    run('registry_preservation',[sys.executable,'-B','tests/production/test_registry_preservation.py'])
    isolated=env|{'LOCALAPPDATA':str(out/'smoke-local'),'PATH':str(Path(os.environ['SystemRoot'])/'System32')}
    run('frozen_smoke',[str(exe),'--task','smoke','--result',str(out/'smoke.json')],isolated)
    assert read(out/'smoke.json')['frozen'] is True
    work=Path(r'C:\Users\PUBLIC_USER\Desktop\ВЗМК рабочий стол\Работа')
    request={'registry_path':str(work/'ВЗМК (2).xlsx'),'project':'ТЭМПО 1','number':'1/1/2026',
             'folder':str(work/'Документы/ТЭМПО/1 выдача/Паспорта/1 от 02.09.2026')}
    save(out/'existing_request.json',request)
    run('real_tempo_readonly',[str(exe),'--task','inspect-existing','--request',str(out/'existing_request.json'),'--result',str(out/'existing.json')],isolated)
    run('frozen_e2e',[sys.executable,'-B','tests/production/e2e.py','--exe',str(exe),'--root',str(out/'e2e'),
        '--reuse-forms',r'C:\VZMKPassportChecks\20261018\production_source\documents\project\forms'])
    run('frozen_fault_recovery',[sys.executable,'-B','tests/production/acceptance_rc1.py','--base',str(out/'e2e'),'--output',str(out/'fault')])
    run('gui_components',[sys.executable,'-B','tests/production/settings_rc1.py','--exe',str(exe),'--base',str(out/'e2e'),'--output',str(out/'gui-components')])
    run('foreign_excel',[sys.executable,'-B','tests/production/foreign_excel.py','--base',str(out/'e2e'),'--output',str(out/'foreign-excel')])
    inputs=read(ROOT/'validation/production100-final2/BUILD_INPUTS.json')
    assert all(sha(ROOT/path)==h for path,h in inputs['files'].items())
    save(out/'RESULT.json',{'status':'PASS','exe':str(exe),'exe_sha256':sha(exe),'checks':results,
         'frozen_without_python_path':True,'actual_issue':'NOT_RUN','production_registry_written':False,
         'build_inputs_unchanged':True,'release_readiness':'BLOCKED_REAL_FORMS_AND_SHIPMENT_INPUTS'})
    print('ACCEPTANCE_PASS; real release remains BLOCKED',flush=True)
if __name__=='__main__':main()
