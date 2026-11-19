"""Check frozen GUI startup on an existing DEV stand with isolated preferences."""
import argparse,json,os,subprocess
from pathlib import Path

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe',type=Path,required=True)
    parser.add_argument('--workspace',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();out=args.output.resolve();out.mkdir(parents=True,exist_ok=False)
    env=dict(os.environ);env['LOCALAPPDATA']=str(out/'appdata')
    env['PATH']=str(Path(os.environ['SystemRoot'])/'System32')
    completed=subprocess.run([str(args.exe.resolve()),'--task','smoke','--workspace',str(args.workspace),
                              '--result',str(out/'RESULT.json')],env=env,timeout=90)
    result=json.loads((out/'RESULT.json').read_text(encoding='utf8'))
    assert completed.returncode==0 and result['status']=='PASS',result
    assert result['version']=='0.9.14' and result['frozen'] and result['gui'],result
    assert result['python_in_path'] is None,result
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
