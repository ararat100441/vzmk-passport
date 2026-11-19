"""Read-only final process, Office version and evidence readback."""
from common import *
import json
import win32api,win32process,win32event
from stage9_app.ownership import process_image,office_image
from datetime import datetime,timezone
assert not (OUT/'FINAL_AUDIT.json').exists()
finals={name:read(OUT/(name+'.json')) for name in ('LIFECYCLE_FINAL','OWNERSHIP_STOP_FINAL','EXE_E2E_FINAL','EXE_NEGATIVE_FINAL')}
assert all(d['status']=='PASS' and d['exe_sha256']==sha(EXE) for d in finals.values())
diagnostics=Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/diagnostics'
records={}
for pattern in ('job_*/created_*.json','job_*/office_*_START.json','viewer_*_START.json'):
    for p in diagnostics.glob(pattern):
        try:
            r=read(p)
            if r.get('owned') is True:records[(r['pid'],r['creation_time'])]=r|{'record':str(p)}
        except (OSError,ValueError,KeyError):pass
active=[];owned_alive=[];record_states=[]
for (pid,birth),r in records.items():
    handle=None
    try:
        handle=win32api.OpenProcess(0x100400,False,pid)
        same=win32process.GetProcessTimes(handle)['CreationTime'].isoformat()==birth
        alive=same and win32event.WaitForSingleObject(handle,0)==258
        if alive:owned_alive.append(r)
        record_states.append({'pid':pid,'creation_time':birth,'same_process_alive':alive,'record':r['record']})
    except win32api.error as exc:
        if exc.winerror not in (87,1168):raise
        record_states.append({'pid':pid,'creation_time':birth,'same_process_alive':False,'record':r['record'],'open_error':exc.winerror})
    finally:
        if handle is not None:handle.Close()
for pid in win32process.EnumProcesses():
    handle=None
    try:
        handle=win32api.OpenProcess(0x100400,False,pid);image=process_image(handle)
        if Path(image).name.upper() in ('EXCEL.EXE','WINWORD.EXE'):
            active.append({'pid':pid,'creation_time':win32process.GetProcessTimes(handle)['CreationTime'].isoformat(),'image':image})
    except win32api.error:pass
    finally:
        if handle is not None:handle.Close()
users=read(finals['LIFECYCLE_FINAL']['evidence'])['users']
for pid,r in users.items():assert any(a['pid']==int(pid) and a['creation_time']==r['creation_time'] and a['image']==r['image'] for a in active),('USER_OFFICE_CHANGED',pid)
assert not owned_alive,owned_alive
versions={}
for name in ('Excel.Application','Word.Application'):
    p=office_image(name);info=win32api.GetFileVersionInfo(p,'\\');ms,ls=info['FileVersionMS'],info['FileVersionLS']
    versions[name]={'path':p,'version':'.'.join(str(n) for n in (ms>>16,ms&65535,ls>>16,ls&65535))}
result={'status':'PASS','timestamp_utc':datetime.now(timezone.utc).isoformat(),'exe_sha256':sha(EXE),'final_results':finals,'owned_records_checked':len(records),'owned_orphans':0,'record_states':record_states,'user_office_preserved':True,'user_baseline':users,'active_office':active,'active_unknown_processes_not_attached_or_terminated':True,'office_versions':versions,'independent_work_review':'NOT_RUN','interactive_user_gui_e2e':'NOT_RUN'}
save(OUT/'FINAL_AUDIT.json',result)
print(json.dumps({k:v for k,v in result.items() if k not in ('record_states','final_results')},ensure_ascii=False,indent=2))
