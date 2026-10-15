import os,sys,json,subprocess,time
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
EXE=ROOT/'stage9_dist_091/VZMK-Passport-0.9.1/VZMK-Passport-0.9.1.exe'
work=json.loads(next(OUT.glob('lifecycle_*/import_*.json')).read_text(encoding='utf8'))['workspace']
base=Path(json.loads((OUT/'RUNTIME_ENV.json').read_text())['LOCALAPPDATA'])
env=os.environ.copy();env.update(LOCALAPPDATA=str(base),PATH=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot']);env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
if sys.argv[1]=='serial':
    folder=OUT/'serial_normal';folder.mkdir();target=folder/'RESULT.json';start=time.monotonic()
    p=subprocess.run([str(EXE),'--task','office-probe','--workspace',work,'--probe-mode','normal','--result',str(target)],env=env,cwd=folder,timeout=360,capture_output=True)
    d=json.loads(target.read_text(encoding='utf8'));s=json.loads(Path(d['office_supervisor']).read_text(encoding='utf8'))
    (folder/'DETAIL.json').write_text(json.dumps({'code':p.returncode,'seconds':time.monotonic()-start,'result':d,'supervisor':s,'no_other_review_office_scenario_running':True},ensure_ascii=False,indent=2),encoding='utf8')
    print(p.returncode,d.get('error'),'orphan',s['owned_orphans'],flush=True)
else:
    local=base/'gui';cfg=local/'VZMKPassport/settings.json';cfg.parent.mkdir(parents=True);cfg.write_text(json.dumps({'workspace':work,'version':'0.9.1'}),encoding='utf8')
    env['LOCALAPPDATA']=str(local);cwd=OUT/'gui_cwd';cwd.mkdir()
    p=subprocess.Popen([str(EXE)],env=env,cwd=cwd)
    import win32process
    proof={'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'exe':str(EXE),'cwd':str(cwd),'LOCALAPPDATA':str(local),'path':env['PATH'],'workspace':work}
    (OUT/'GUI_LAUNCH.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(proof,ensure_ascii=False),flush=True)
