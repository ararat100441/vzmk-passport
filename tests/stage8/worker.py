"""Fresh-process close/recovery/crash worker; only bootstrapped TEST sandboxes."""
import json
import os
from pathlib import Path
import sys
import time
from stage8_closure import core
from stage8_closure.storage import read, sandbox

command,path=sys.argv[1:3]
try:
    if command=='recover':result=core.recover(path)
    else:
        plan=read(path);sandbox(plan['payload']['sandbox'])
        if command=='parallel':
            marker=Path(sys.argv[3])
            start=time.monotonic()
            while not marker.exists():
                if time.monotonic()-start>20:raise RuntimeError('BARRIER_TIMEOUT')
                time.sleep(.02)
        def fault(point,target):
            if command=='crash' and point==sys.argv[3]:os._exit(73)
        result=core.close(plan,core.confirmation(plan),_fault=fault)
    print(json.dumps(result,ensure_ascii=False))
except Exception as exc:
    print(json.dumps({'status':'STOP','error':str(exc)},ensure_ascii=False));sys.exit(2)
