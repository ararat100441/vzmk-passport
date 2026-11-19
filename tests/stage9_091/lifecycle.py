from common import *
import win32api,win32process,win32event,traceback
from stage9_app.ownership import process_image
RUN=OUT/('lifecycle_'+uuid4().hex[:8]);RUN.mkdir()
users={};old_owned={}
for p in (Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/diagnostics').glob('job_*/office_*_START.json'):
    try:
        r=read(p)
        if r.get('owned') is True:old_owned[(r['pid'],r['creation_time'])]=str(p)
    except Exception:pass
for pid in win32process.EnumProcesses():
    try:
        h=win32api.OpenProcess(0x100400,False,pid);image=process_image(h)
        if Path(image).name.upper() in ('EXCEL.EXE','WINWORD.EXE'):
            born=win32process.GetProcessTimes(h)['CreationTime'].isoformat()
            if (pid,born) not in old_owned:users[pid]={'creation_time':born,'image':image}
        h.Close()
    except Exception:pass
save(RUN/'USER_OFFICE_BEFORE.json',users)
try:
    work,_,_=task(RUN,'import',source=ROOT,parent=Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/workspaces')
    work=work['workspace'];cases=[]
    for mode in ('normal','error','crash','normal'):
        data,sup,path=task(RUN,'office-probe',expected='PASS' if mode=='normal' else 'FAIL',workspace=work,probe_mode=mode)
        assert len(sup['instances'])==2 and all(r.get('actual_child') and r['handle_source']=='CREATE_PROCESS' and r['launcher_pid']!=sup['worker_pid'] for r in sup['instances']),sup
        if mode!='crash':assert all(not r.get('forced_termination_attempted') for r in sup['instances']),sup
        if mode=='normal':assert all(r['quit_returned'] for r in sup['instances'])
        if mode=='crash':
            assert sup['worker_returncode']==73 and not sup['normal_result']
            assert all(not r.get('termination_error') and not r['still_running'] for r in sup['instances']),sup
            for r in sup['instances']:
                if r.get('forced_termination_attempted'):
                    proof=read(Path(data['office_supervisor']).parent/('office_FORCE_BEGIN_'+str(r['pid'])+'.json'))
                    assert proof['handle_validated'] and proof['retained_handle']==r['retained_handle'] and proof['handle_source']=='CREATE_PROCESS' and proof['creation_time']==r['creation_time']
        if mode=='error':assert 'STAGE9_TEST_OFFICE_ERROR' in data['error']
        for pid,before in users.items():
            h=win32api.OpenProcess(0x100400,False,pid)
            assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==before['creation_time'] and win32event.WaitForSingleObject(h,0)==258
            h.Close()
        cases.append({'mode':mode,'result':str(path),'supervisor':data['office_supervisor'],'status':'PASS','user_office_preserved':True})
    save(RUN/'RESULT.json',{'status':'PASS','exe_sha256':sha(EXE),'workspace':work,'cases':cases,'users':users})
    save(OUT/'LIFECYCLE_FINAL.json',{'status':'PASS','evidence':str(RUN/'RESULT.json'),'exe_sha256':sha(EXE),'workspace':work})
except Exception as exc:
    save(RUN/'FAILURE.json',{'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc(),'exe_sha256':sha(EXE)});raise
