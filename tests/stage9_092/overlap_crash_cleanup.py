"""A normal frozen job overlaps deliberate worker death and its owned cleanup."""
from common import *
import time,concurrent.futures
RUN=OUT/('overlap_crash_'+uuid4().hex[:8]);RUN.mkdir();work=SETUP['probe_workspace']
def call(mode):
    started=time.time();data,sup,path=task(RUN,'office-probe','PASS' if mode=='normal' else 'FAIL',workspace=work,probe_mode=mode)
    return {'mode':mode,'started':started,'ended':time.time(),'result':str(path),'supervisor':data['office_supervisor'],'force':any(r.get('forced_termination_attempted') for r in sup['instances'])}
with concurrent.futures.ThreadPoolExecutor(2) as pool:
    crash=pool.submit(call,'crash');time.sleep(5);normal=pool.submit(call,'normal');cases=[crash.result(),normal.result()]
assert not cases[1]['force']
save(RUN/'RESULT.json',{'status':'PASS','cases':cases,'exe_sha256':sha(EXE)})
save(OUT/'OVERLAP_CRASH_FINAL.json',{'status':'PASS','evidence':str(RUN/'RESULT.json')})
