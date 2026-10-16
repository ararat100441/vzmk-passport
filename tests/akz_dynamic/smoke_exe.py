"""Launch the separate DEV/TEST EXE with copied AKZ inputs and isolated settings.

This is a version/packaging/GUI startup check, not full passport generation.
"""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from uuid import uuid4

sys.path.insert(0,str(Path(__file__).resolve().parents[2]))
from stage8_closure import storage as S
from stage9_app.workspace import save

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--exe',type=Path,required=True)
    parser.add_argument('--source',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();out=args.output.resolve();out.mkdir(parents=True,exist_ok=False)
    root=out/uuid4().hex;root.mkdir()
    for name in ('DEV','documents','validation'):(root/name).mkdir()
    inputs={};source=args.source.resolve()
    files=[(source/'DEV/ВЗМК_DEV.xlsx','DEV/ВЗМК_DEV.xlsx')]
    project=source.parent/'Гараж казахстан/Паспорт'
    names=['Акт № 4-2026.xlsx','1_2026 - Приложение к АКЗ.xlsx']
    for n in (1,2,3):names += [f'АКЗ-{n}.xlsx',f'Приложение к АКЗ-{n}.xlsx']
    files += [(project/name,'documents/project/Паспорт/'+name) for name in names]
    with S.frozen([path for path,_ in files]):
        for path,relative in files:
            dst=root/relative;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(path,dst)
            assert S.sha(path)==S.sha(dst)
            inputs[relative]={'source':str(path),'sha256':S.sha(path)}
    meta={'schema':'STAGE9_WORKSPACE_V1','workspace_id':root.name,'official_issue':False,
          'project':'Гараж Казахстан','inputs':inputs,'test_purpose':'AKZ_ONLY_GUI_STARTUP'}
    save(root/'STAGE9_WORKSPACE.json',meta|{'seal':S.digest(meta)})
    env=dict(os.environ);env['LOCALAPPDATA']=str(out/'appdata')
    env['PATH']=str(Path(env['SystemRoot'])/'System32')
    completed=subprocess.run([str(args.exe.resolve()),'--task','smoke','--workspace',str(root),
                              '--result',str(out/'RESULT.json')],env=env,timeout=90)
    result=json.loads((out/'RESULT.json').read_text(encoding='utf8'))
    assert completed.returncode==0 and result['status']=='PASS',result
    assert result['version']=='0.9.13' and result['frozen'] and result['gui'],result
    assert result['python_in_path'] is None,result
    assert all(S.sha(path)==inputs[relative]['sha256'] for path,relative in files)
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
