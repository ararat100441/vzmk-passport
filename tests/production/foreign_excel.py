"""A separately owned Excel instance locks only a disposable registry copy."""
import argparse,copy,json,os,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from production_app import core,model
from production_app.storage import read,save,sha

def main():
    p=argparse.ArgumentParser();p.add_argument('--base',required=True);p.add_argument('--output',required=True);a=p.parse_args()
    base=Path(a.base);out=Path(a.output);out.mkdir(exist_ok=False)
    accepted=read(base/'RESULT.json');cfg=core.training_copy(accepted['cfg'],out)
    request=next(read(p) for p in (base/'calls').glob('*/request.json') if 'facts' in read(p))
    request=copy.deepcopy(request);request['cfg']=cfg
    request['number']=model.select(model.read_registry(cfg['registry']),request['project']).next_number
    request['facts']['number']=request['number'];save(out/'request.json',request)
    registry=Path(cfg['registry']);before=sha(registry)
    import pythoncom,win32com.client,win32process,win32api,win32con,win32event
    pythoncom.CoInitialize();app=None;book=None;handle=None;owned=False
    try:
        before_pids=set(win32process.EnumProcesses());app=win32com.client.DispatchEx('Excel.Application')
        _,pid=win32process.GetWindowThreadProcessId(app.Hwnd)
        if pid in before_pids:raise RuntimeError('FOREIGN_TEST_INSTANCE_NOT_PROVEN_NEW; not closing it')
        owned=True;handle=win32api.OpenProcess(win32con.SYNCHRONIZE|win32con.PROCESS_QUERY_INFORMATION,False,pid)
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.AutomationSecurity=3;app.AskToUpdateLinks=False
        book=app.Workbooks.Open(str(registry),UpdateLinks=0,ReadOnly=False,AddToMru=False)
        env=dict(os.environ);env['LOCALAPPDATA']=str(out/'local');env['PATH']=str(Path(os.environ['SystemRoot'])/'System32')
        proc=subprocess.Popen([accepted['exe'],'--task','check','--request',str(out/'request.json'),'--result',str(out/'result.json')],env=env)
        proc.wait(timeout=120)
        result=read(out/'result.json');assert proc.returncode!=0 and result['status']=='FAIL'
        assert 'занят' in str(result).lower() or 'доступ' in str(result).lower(),result
        assert app.Workbooks.Count==1 and not book.ReadOnly
        assert sha(registry)==before
        supervisor=read(result['office_supervisor'])
        assert not supervisor['instances'] and supervisor['owned_orphans']==0
        observation={'pid':pid,'still_open_after_app_refusal':True,'workbooks':1,'registry_sha256_unchanged':True}
    finally:
        if owned:
            if book is not None:book.Close(SaveChanges=False);book=None
            app.Quit();app=None
        pythoncom.CoUninitialize()
        if handle is not None:
            code=win32event.WaitForSingleObject(handle,30000);win32api.CloseHandle(handle)
            assert code==0,'Test-owned Excel did not exit gracefully; not force terminating'
    assert sha(registry)==before
    save(out/'RESULT.json',{'status':'PASS','scope':'Separately owned native Excel; disposable registry only',
         'foreign_excel_preserved_by_app':observation,'test_owned_excel_graceful_exit':True,'forced_terminations':0,
         'production_inputs_written':False})
    print('FOREIGN_EXCEL_LOCK_PASS')
if __name__=='__main__':main()
