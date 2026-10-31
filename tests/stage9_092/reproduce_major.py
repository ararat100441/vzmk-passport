"""Run the immutable 0.9.1 executable concurrently and sample owned window state."""
import sys,os,time,subprocess,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage8_closure import storage as S
from stage9_app.workspace import save
import win32gui,win32process,win32api,win32event
OUT=ROOT/'validation/stage9_092_20261012_01';setup=S.read(OUT/'SETUP.json');run=OUT/(sys.argv[1] if len(sys.argv)>1 else 'major_reproduction');run.mkdir()
exe=ROOT/'stage9_dist_091/VZMK-Passport-0.9.1/VZMK-Passport-0.9.1.exe';assert S.sha(exe)=='677a111e1762e72949125e74977ca3468f114895a2de92f705c49264aec61c88'
env=os.environ.copy();env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None);env['LOCALAPPDATA']=setup['base']
cases=[]
for group in range(3):
    active=[]
    for n in range(int(sys.argv[3]) if len(sys.argv)>3 else 2):
        e=env.copy();temp=Path(setup['base'])/f'{run.name}_temp_{group}_{n}';temp.mkdir();e['TEMP']=e['TMP']=str(temp)
        result=run/f'case_{group}_{n}.json';started=time.time()
        p=subprocess.Popen([str(exe),'--task','office-probe','--workspace',setup['probe_workspace'],'--result',str(result)],env=e,creationflags=subprocess.CREATE_NO_WINDOW)
        active.append((p,result,started))
        if len(sys.argv)>2:time.sleep(float(sys.argv[2]))
    samples=[]
    while any(p.poll() is None for p,_,_ in active):
        for created in (Path(setup['base'])/'VZMKPassport/diagnostics').glob('job_*/created_*.json'):
            r=S.read(created)
            try:
                h=win32api.OpenProcess(0x100400,False,r['pid'])
                try:
                    if win32process.GetProcessTimes(h)['CreationTime'].isoformat()!=r['creation_time'] or win32event.WaitForSingleObject(h,0)!=258:continue
                    windows=[]
                    def window(hwnd,_):
                        if win32process.GetWindowThreadProcessId(hwnd)[1]==r['pid']:
                            windows.append({'hwnd':hwnd,'class':win32gui.GetClassName(hwnd),'title':win32gui.GetWindowText(hwnd),'visible':bool(win32gui.IsWindowVisible(hwnd)),'enabled':bool(win32gui.IsWindowEnabled(hwnd))})
                    win32gui.EnumWindows(window,None)
                    samples.append({'time':time.time(),'record':str(created),'windows':windows})
                finally:h.Close()
            except Exception:pass
        if time.time()-min(s for _,_,s in active)>180:raise RuntimeError('OLD_EXE_CONTROLLER_TIMEOUT')
        time.sleep(1)
    for p,path,started in active:
        data=S.read(path);sup=S.read(data['office_supervisor']);cases.append({'result':str(path),'returncode':p.returncode,'seconds':time.time()-started,'supervisor':sup})
        print(group,path.name,data['status'],data.get('error'),flush=True)
    save(run/f'windows_{group}.json',samples)
    if any(c['returncode'] for c in cases):break
save(run/'RESULT.json',{'old_exe_sha256':S.sha(exe),'cases':cases,'reproduced':any(c['returncode'] for c in cases)})
