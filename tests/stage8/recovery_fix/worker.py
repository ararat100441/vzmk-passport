import os,sys,json,time
from pathlib import Path
from stage8_closure import core,storage as S,office
cmd,path=sys.argv[1:3];path=Path(path)
try:
    with office.diagnostic_scope(path.parent/'worker_office'):
        def fault(point,target):
            if len(sys.argv)>3 and point==sys.argv[3]:os._exit(73)
        if cmd=='crash_close':
            p=S.read(path);result=core.close(p,core.confirmation(p),_fault=fault)
        elif cmd=='recover':result=core.recover(path)
        elif cmd=='preview':result=core.preview(path,path/'inputs/CLOSE_REQUEST.json')
        elif cmd in ('restore','crash_restore','parallel_restore'):
            from stage8_closure import recovery
            if cmd=='parallel_restore':
                marker=Path(sys.argv[3]);end=time.monotonic()+25
                while not marker.exists():
                    if time.monotonic()>end:raise S.Stop('TEST_BARRIER_TIMEOUT')
                    time.sleep(.01)
            p=S.read(path);result=recovery.restore(p,recovery.confirmation(p),_fault=fault)
        elif cmd=='restore_preview':
            from stage8_closure import recovery
            result=recovery.preview(path)
        else:raise ValueError(cmd)
        print(json.dumps(result,ensure_ascii=False),flush=True)
except Exception as e:
    print(json.dumps({'status':'STOP','error':office._error(e)},ensure_ascii=False),flush=True);sys.exit(2)
