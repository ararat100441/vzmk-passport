"""Read-only binding and Office monitor; no UI automation or process termination."""
import os,sys,json,time,hashlib,subprocess
from pathlib import Path
from datetime import datetime,timezone
import win32api,win32con,win32event,win32process
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
PRIOR=ROOT/'validation/stage9_resume_099_20261012_01'
def read(p):return json.loads(Path(p).read_text(encoding='utf8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(p,d):
    with Path(p).open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str);f.flush();os.fsync(f.fileno())
def processes():
    result=subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-Command',"Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='EXCEL.EXE' OR Name='VZMK-Passport-0.9.9.exe'\" | Select-Object Name,ProcessId,ParentProcessId,CreationDate,ExecutablePath,CommandLine | ConvertTo-Json -Depth 3"],capture_output=True,check=True,encoding='utf8',errors='replace',creationflags=subprocess.CREATE_NO_WINDOW)
    data=json.loads(result.stdout) if result.stdout.strip() else []
    return [data] if isinstance(data,dict) else data
if __name__=='__main__':
    phase=sys.argv[1] if len(sys.argv)>1 else 'BEFORE'
    if phase!='MONITOR':
        source=read(PRIOR/'SOURCE_SHA256_AFTER.json');build=read(PRIOR/'BUILD_SHA256_AFTER.json');accepted=read(PRIOR/'SOURCE_INTEGRITY_AFTER.json')['accepted_files']
        assert all(sha(p)==h for p,h in source['files'].items())
        exe=Path(build['exe']);files={str(p.relative_to(exe.parent)):sha(p) for p in exe.parent.rglob('*') if p.is_file()}
        assert files==build['files'] and sha(exe)==build['exe_sha256']
        assert all(sha(p)==h for p,h in accepted.items())
        e2e=read(PRIOR/'EXE_E2E_FINAL.json');workspace=Path(e2e['workspace'])
        test_files={str(p):sha(p) for p in workspace.rglob('*') if p.is_file()}
        if phase=='AFTER':assert test_files==read(OUT/'BINDING_BEFORE.json')['test_files'],'TEST_WORKSPACE_CHANGED'
        save(OUT/('BINDING_'+phase+'.json'),{'status':'PASS','at':datetime.now(timezone.utc).isoformat(),'source':source,'build':build,'accepted_files':accepted,'workspace':str(workspace),'manifests':e2e['manifests'],'test_files':test_files,'processes':processes(),'historical_word_3132_cause':'UNKNOWN'})
        print('BINDING',phase,'PASS',len(files),'onedir',len(accepted),'accepted',flush=True)
    else:
        handles={};records=[];log=(OUT/'OFFICE_MONITOR.jsonl').open('x',encoding='utf8')
        def emit(d):log.write(json.dumps(d,ensure_ascii=False,default=str)+'\n');log.flush();os.fsync(log.fileno())
        diag_roots=[Path(r'C:\Users\arara\AppData\Local\fix098_4b9519ff\VZMKPassport\diagnostics'),Path(r'C:\Users\arara\AppData\Local\fix098_4b9519ff\new099\VZMKPassport\diagnostics')]
        initial=True
        try:
            while not (OUT/'MONITOR_STOP.json').exists():
                rows=processes();owned={}
                for root in diag_roots:
                    if root.exists():
                        for p in root.glob('viewer_*/ownership_*_START.json'):
                            try:
                                r=read(p)
                                if r.get('launcher_pid')==47440 and r.get('actual_child') is True and r.get('owned') is True and r.get('handle_source')=='CREATE_PROCESS':owned[r['pid']]=r
                            except (OSError,ValueError):pass
                for r in rows:
                    if r['Name'].upper() not in ('WINWORD.EXE','EXCEL.EXE'):continue
                    pid=r['ProcessId']
                    if pid not in handles:
                        h=win32api.OpenProcess(win32con.SYNCHRONIZE|0x1000,False,pid)
                        birth=win32process.GetProcessTimes(h)['CreationTime'].isoformat()
                        item={'pid':pid,'creation_time':birth,'image':r['ExecutablePath'],'parent':r['ParentProcessId'],'baseline':initial,'owned':False,'exit_observed':False}
                        handles[pid]=(h,item);records.append(item);emit({'event':'OBSERVED','record':dict(item),'at':time.time()})
                for pid,(h,item) in handles.items():
                    if pid in owned:
                        r=owned[pid]
                        if datetime.fromisoformat(item['creation_time']).timestamp()==datetime.fromisoformat(r['creation_time']).timestamp() and item['image'].casefold()==r['image'].casefold() and item['parent']==r['launcher_pid']:
                            if not item['owned']:emit({'event':'OWNERSHIP_CONFIRMED','record':r,'at':time.time()})
                            item['owned']=True
                    if not item['exit_observed'] and win32event.WaitForSingleObject(h,0)==win32con.WAIT_OBJECT_0:
                        item['exit_observed']=True;item['exit_code']=win32process.GetExitCodeProcess(h);emit({'event':'EXIT','record':dict(item),'at':time.time()})
                        if not item['owned']:
                            save(OUT/'FOREIGN_OFFICE_STOP.json',{'reason':'UNEXPECTED_NON_TEST_OFFICE_EXIT','record':item,'cause':'UNKNOWN'});print('STOP FOREIGN OFFICE EXIT',pid,flush=True);raise RuntimeError('FOREIGN_OFFICE_EXIT_STOP')
                if initial:save(OUT/'OFFICE_BASELINE.json',{'at':time.time(),'office':[dict(v) for h,v in handles.values()],'historical_word_3132_cause':'UNKNOWN'});print('MONITOR READY baseline',len(handles),flush=True)
                initial=False;time.sleep(.25)
        finally:
            for h,item in handles.values():h.Close()
            save(OUT/'OFFICE_MONITOR_FINAL.json',{'instances':records,'foreign_stop':(OUT/'FOREIGN_OFFICE_STOP.json').exists(),'monitor_only':True,'no_termination_api':True});log.close()
