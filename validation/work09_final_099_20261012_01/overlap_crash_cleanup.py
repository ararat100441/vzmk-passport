# Adapted independent run controller; original SHA 9b3c4c7f384b5fbd4b8f1a7037fab39cd7cf0dac83fa8ac2b6bce0d67a185e31
"""A normal frozen job overlaps deliberate worker death and its owned cleanup."""
from common import *
import time,concurrent.futures
RUN=OUT/('overlap_crash_'+uuid4().hex[:8]);RUN.mkdir();work=read(OUT/'LIFECYCLE_FINAL.json')['workspace']
def call(mode):
    started=time.time();data,sup,path=task(RUN,'office-probe','PASS' if mode=='normal' else 'FAIL',workspace=work,probe_mode=mode)
    return {'mode':mode,'started':started,'ended':time.time(),'result':str(path),'supervisor':data['office_supervisor'],'force':any(r.get('forced_termination_attempted') for r in sup['instances'])}
with concurrent.futures.ThreadPoolExecutor(2) as pool:
    crash=pool.submit(call,'crash');time.sleep(5);normal=pool.submit(call,'normal');cases=[crash.result(),normal.result()]
assert not cases[1]['force']
save(RUN/'RESULT.json',{'status':'PASS','cases':cases,'exe_sha256':sha(EXE)})
save(OUT/'OVERLAP_CRASH_FINAL.json',{'status':'PASS','evidence':str(RUN/'RESULT.json')})
