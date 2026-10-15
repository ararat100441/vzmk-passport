"""Real concurrent/crash process; operates only on new bootstrapped TEST data."""
import os,sys,time,json
from pathlib import Path
from stage8_closure import core,storage as S
from stage8_closure.office import diagnostic_scope

command,path=sys.argv[1:3]
try:
    with diagnostic_scope(Path(path).parent/'worker_office'):
        if command=='recover':result=core.recover(path)
        else:
            plan=S.read(path)
            if command=='parallel':
                marker=Path(sys.argv[3]);end=time.monotonic()+30
                while not marker.exists():
                    if time.monotonic()>end:raise RuntimeError('BARRIER_TIMEOUT')
                    time.sleep(.01)
            def fault(point,target):
                if command=='crash' and point==sys.argv[3]:os._exit(73)
                if command=='save_event_crash' and point=='save_io':
                    # A real Excel BeforeSave COM callback proves Save is active.
                    import win32com.client
                    class Events:
                        def OnWorkbookBeforeSave(self,Wb,SaveAsUI,Cancel):
                            S.write(Path(path).with_name('SAVE_CALLBACK_ENTERED.json'),
                                    {'time':S.now(),'workbook':str(Wb.FullName),'phase':'INSIDE_NATIVE_SAVE'})
                            os._exit(73)
                    from stage8_closure.office import _apartment
                    record=_apartment.current_record
                    excel=win32com.client.GetObject(str(target)).Application
                    assert int(excel.Hwnd)
                    listener=win32com.client.WithEvents(excel,Events)
                    excel.EnableEvents=True
                    # Keep callback proxy alive during the following Save call.
                    fault.listener=listener
            result=core.close(plan,core.confirmation(plan),_fault=fault)
        print(json.dumps(result,ensure_ascii=False),flush=True)
except Exception as exc:
    print(json.dumps({'status':'STOP','error':str(exc)},ensure_ascii=False),flush=True)
    sys.exit(2)
