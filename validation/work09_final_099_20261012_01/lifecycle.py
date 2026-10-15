# Adapted independent run controller; original SHA 352a1db841dc2bcadc6d0292062a35d6719e7aa5141e9e293accd189fa6a6412
from common import *
import win32api,win32process,win32event,traceback
from stage9_app.ownership import process_image
RUN=OUT/('lifecycle_'+uuid4().hex[:8]);RUN.mkdir()
users={}
for pid in win32process.EnumProcesses():
    try:
        h=win32api.OpenProcess(0x100400,False,pid)
        try:
            image=process_image(h)
            if Path(image).name.upper() in ('EXCEL.EXE','WINWORD.EXE'):
                users[pid]={'image':image,'creation_time':win32process.GetProcessTimes(h)['CreationTime'].isoformat()}
        finally:h.Close()
    except __import__('pywintypes').error:pass
save(RUN/'USER_OFFICE_BEFORE.json',users)
try:
    work,_,_=task(RUN,'import',source=SOURCE,parent=WORKSPACES)
    work=work['workspace'];cases=[]
    for mode in ('normal','error','cleanup-error','crash','normal'):
        data,sup,path=task(RUN,'office-probe',expected='PASS' if mode=='normal' else 'FAIL',workspace=work,probe_mode=mode)
        assert len(sup['instances'])==2 and all(r.get('actual_child') and r['handle_source']=='CREATE_PROCESS' and r['launcher_pid']!=sup['worker_pid'] for r in sup['instances']),sup
        if mode!='crash':assert all(not r.get('forced_termination_attempted') for r in sup['instances']),sup
        if mode!='crash':
            worker_office=read(Path(data['office_supervisor']).parent/'office.json')
            assert all(r['exit_observed_while_sta_alive'] and r['exit_code']==0 for r in worker_office['instances']),worker_office
        if mode=='normal':assert all(r['quit_returned'] for r in sup['instances'])
        for r in sup['instances']:
            if r['progid']=='Excel.Application':
                assert isinstance(r['user_control_before_binding'],bool) and r['user_control_after_binding'] is False
                if mode!='crash':assert r['user_control_before_quit'] is False and r['workbooks_before_quit']==0
        if mode=='crash':
            assert sup['worker_returncode']==73 and not sup['normal_result']
            assert all(not r.get('termination_error') and not r['still_running'] for r in sup['instances']),sup
            for r in sup['instances']:
                if r.get('forced_termination_attempted'):
                    proof=read(Path(data['office_supervisor']).parent/('office_FORCE_BEGIN_'+str(r['pid'])+'.json'))
                    assert proof['handle_validated'] and proof['retained_handle']==r['retained_handle'] and proof['handle_source']=='CREATE_PROCESS' and proof['creation_time']==r['creation_time']
        if mode in ('error','cleanup-error'):assert 'STAGE9_TEST_OFFICE_ERROR' in data['error']
        if mode=='cleanup-error':
            assert data['worker_result']['cleanup_errors']==[{'step':'Word.Close','error':"RuntimeError('STAGE9_TEST_WORD_CLOSE_ERROR')"}]
            assert all(r['quit_returned'] for r in sup['instances'])
        cases.append({'mode':mode,'result':str(path),'supervisor':data['office_supervisor'],'status':'PASS','user_session_continuity':'OBSERVATIONAL_ONLY; controlled foreign test separately'})
    save(RUN/'RESULT.json',{'status':'PASS','exe_sha256':sha(EXE),'workspace':work,'cases':cases,'users':users})
    save(OUT/'LIFECYCLE_FINAL.json',{'status':'PASS','evidence':str(RUN/'RESULT.json'),'exe_sha256':sha(EXE),'workspace':work})
except Exception as exc:
    save(RUN/'FAILURE.json',{'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc(),'exe_sha256':sha(EXE)});raise
