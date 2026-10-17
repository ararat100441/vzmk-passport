"""One new targeted Stage 9 TEST run, source or frozen, with Office evidence."""
import argparse,hashlib,json,os,subprocess,sys,time
from decimal import Decimal
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output',type=Path,required=True);parser.add_argument('--exe',type=Path)
    args=parser.parse_args();out=args.output.resolve();out.mkdir(parents=True,exist_ok=False)
    workspace=Path(r'C:\Users\arara\AppData\Local\VZMKPassport\workspaces\19c900ed8f7a42f3be9fbc8aa70052f4')
    prefix=[str(args.exe.resolve())] if args.exe else [sys.executable,'-B',str(ROOT/'stage9.py')]
    env=dict(os.environ);env['LOCALAPPDATA']=str(out/'appdata')
    if args.exe:env['PATH']=str(Path(os.environ['SystemRoot'])/'System32')
    reports=[]
    def run(task,arguments):
        target=out/(task+'.json');began=time.monotonic()
        p=subprocess.Popen(prefix+['--task',task,'--workspace',str(workspace),*arguments,'--result',str(target)],env=env,
                           stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        # Do not kill the supervisor on a controller deadline: it owns the
        # original Office handles and must remain alive to finish safe cleanup.
        try:p.wait(timeout=600)
        except subprocess.TimeoutExpired:
            (out/(task+'_CONTROLLER_TIMEOUT.json')).write_text(json.dumps({'status':'FAIL','error':'TEST_CONTROLLER_WAIT_LIMIT',
                'supervisor_pid':p.pid,'supervisor_not_terminated':True}),encoding='utf8')
            raise RuntimeError('TEST_CONTROLLER_WAIT_LIMIT; supervisor left alive for safe cleanup')
        result=json.loads(target.read_text(encoding='utf8'))
        assert p.returncode==0 and result['status']=='PASS',result
        summary=json.loads(Path(result['office_supervisor']).read_text(encoding='utf8'))
        assert summary['owned_orphans']==0 and not summary['ownership_errors'],summary
        assert all(r['exit_code']==0 and not r.get('forced_termination_attempted') and not r['still_running'] for r in summary['instances']),summary
        assert 'OFFICE_GRACEFUL_EXIT_TIMEOUT' not in json.dumps(result),result
        reports.append({'task':task,'duration_seconds':time.monotonic()-began,'office':summary,'supervisor_path':result['office_supervisor']})
        print(task,'PASS','office_wait',round(summary['wait_seconds_after_worker'],3),'recovered',summary['owned_unbound_gracefully_recovered'],flush=True)
        return result
    check=run('check',['--project','Гараж Казахстан','--passport','9/2026','--vom',r'C:\Users\arara\Downloads\ВОМ отгрузочный .xlsx','--variants','0'])
    data=check['data']
    assert data['visible_rows']==104 and Decimal(data['quantity'])==288 and Decimal(data['mass_kg'])==Decimal('15822.14') and data['variants']==[0],data
    generated=run('generate',['--preview',check['preview'],'--confirm-test'])
    verified=run('verify',['--manifest',generated['manifest']])
    assert verified['physical_files']==6 and verified['document_records']==7,verified
    manifest=json.loads(Path(generated['manifest']).read_text(encoding='utf8'))
    assert manifest['official_issue'] is False and manifest['visible_rows']==104 and Decimal(manifest['quantity'])==288
    result={'status':'PASS','manifest':generated['manifest'],'files':6,'records':7,'marks':104,'quantity':288,'mass_kg':'15822.14',
            'office':reports,'recovered':sum(r['office']['owned_unbound_gracefully_recovered'] for r in reports),
            'forced_terminations':0,'office_graceful_exit_timeout':False,'official_issue':False}
    (out/'RESULT.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
    print('TARGET_E2E_PASS',generated['manifest'],flush=True)

if __name__=='__main__':main()
