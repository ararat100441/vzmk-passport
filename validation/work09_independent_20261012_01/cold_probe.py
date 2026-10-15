import os,sys,json,subprocess,time
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
EXE=ROOT/'stage9_dist_091/VZMK-Passport-0.9.1/VZMK-Passport-0.9.1.exe'
work=json.loads(next(OUT.glob('lifecycle_*/import_*.json')).read_text(encoding='utf8'))['workspace']
base=Path(json.loads((OUT/'RUNTIME_ENV.json').read_text())['LOCALAPPDATA'])
results=[]
for n in (1,2):
    folder=OUT/('cold_'+str(n));folder.mkdir();temp=base/('temp_'+str(n));temp.mkdir()
    env=os.environ.copy();env.update(LOCALAPPDATA=str(base),TEMP=str(temp),TMP=str(temp),PATH=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot']);env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
    target=folder/'RESULT.json';start=time.monotonic()
    p=subprocess.run([str(EXE),'--task','office-probe','--workspace',work,'--probe-mode','normal','--result',str(target)],env=env,cwd=folder,timeout=360,capture_output=True)
    d=json.loads(target.read_text(encoding='utf8'));s=json.loads(Path(d['office_supervisor']).read_text(encoding='utf8'))
    data={'case':n,'TEMP':str(temp),'code':p.returncode,'seconds':time.monotonic()-start,'result':d,'supervisor':s}
    (folder/'DETAIL.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf8');results.append(data)
    print('cold',n,p.returncode,d.get('error'),'orphan',s['owned_orphans'],flush=True)
(OUT/'COLD_PROBE.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf8')
