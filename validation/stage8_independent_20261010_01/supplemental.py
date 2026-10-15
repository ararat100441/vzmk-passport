import json,os,sys,time,shutil,subprocess,traceback
from pathlib import Path
from contextlib import contextmanager
from unittest.mock import patch
import independent as T
from stage8_closure import core,storage as S
from stage8_closure.office import excel_session
T.SEED=Path(S.read(T.RUN/'SEED.json')['sandbox'])
out=[]
def run(name,fn):
    try:r={'name':name,'status':'PASS','detail':fn()}
    except Exception as exc:r={'name':name,'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()}
    out.append(r);T.save(T.RUN/'cases'/(name+'.json'),r);T.save(T.RUN/'SUPPLEMENTAL_RESULTS.json',out);print(name,r['status'],r.get('error',''),flush=True)
def publisher(mode):
    root=T.fixture();attempt=[]
    def fault(point,target):
        if point!='before_replace':return
        incoming=root/'transactions/OTHER_PUBLISHER.xlsm';shutil.copyfile(T.reg(T.SEED),incoming)
        incoming.write_bytes(incoming.read_bytes()+b'OTHER_PUBLISHER_BYTES')
        if mode=='rename':code='import os,sys;os.rename(sys.argv[2],sys.argv[2]+".external_old");os.replace(sys.argv[1],sys.argv[2]);print("REPLACED")'
        else:code='import sys;from stage8_closure.storage import replace;replace(sys.argv[1],sys.argv[2]);print("REPLACED")'
        r=subprocess.run([sys.executable,'-B','-c',code,str(incoming),str(T.reg(root))],env=T.ENV,capture_output=True,text=True,timeout=15)
        attempt.append({'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr,'sha_after_external_publish':S.sha(T.reg(root))})
    result=T.close(root,fault)
    detail={'sandbox':str(root),'method':mode,'external_writer':attempt,'closure_status':result['status'],'balances':T.balances(root),'journal':T.states(root)}
    T.save(root/'EXTERNAL_PUBLISH_RESULT.json',detail)
    assert attempt[0]['exit_code']!=0, 'Concurrent replacement succeeded; Stage8 overwrote it silently and COMMITTED. Evidence: '+str(root/'EXTERNAL_PUBLISH_RESULT.json')
    return detail
run('external_rename_then_replace_before_publish',lambda:publisher('rename'))
run('external_native_atomic_replace_before_publish',lambda:publisher('native'))

def cleanup():
    import win32process,win32api,win32con,pythoncom
    owned={};root=T.fixture()
    class Proxy:
        def __init__(self,real):object.__setattr__(self,'real',real)
        def __getattr__(self,k):return getattr(self.real,k)
        def __setattr__(self,k,v):
            if k=='CutCopyMode':raise pythoncom.com_error(-2147418111,'INDEPENDENT_TRANSIENT_CLEANUP_REJECTION',None,None)
            setattr(self.real,k,v)
    # Only a cleanup property assignment fails; all workbook operations use real Excel.
    @contextmanager
    def tracked():
        import win32com.client
        original=win32com.client.DispatchEx
        def dispatch(*a,**k):
            real=original(*a,**k);owned['real']=real;owned['pid']=win32process.GetWindowThreadProcessId(real.Hwnd)[1]
            return Proxy(real)
        with patch('win32com.client.DispatchEx',side_effect=dispatch):
            with excel_session() as excel:yield excel
    try:
        with tracked() as excel:
            b=excel.Workbooks.Open(Filename=str(T.reg(root)),ReadOnly=True,UpdateLinks=0,CorruptLoad=0)
            b.Close(SaveChanges=False)
        h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,owned['pid'])
        try:alive=win32process.GetExitCodeProcess(h)==259
        finally:h.Close()
        T.save(root/'CLEANUP_FAILURE.json',{'sandbox':str(root),'pid':owned['pid'],'owned_excel_alive':alive,'error_propagated':False,'Quit_skipped':True})
        assert not alive,'First cleanup COM failure is swallowed; Quit is skipped; own Excel remains alive. Evidence: '+str(root/'CLEANUP_FAILURE.json')
    finally:
        owned['real'].Workbooks.Close();owned['real'].Quit()
run('cleanup_com_failure_must_still_attempt_quit',cleanup)
