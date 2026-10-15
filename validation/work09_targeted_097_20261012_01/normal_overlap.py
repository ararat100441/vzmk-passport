# Adapted independent run controller; original SHA da7cdd95211fe52cbd310b84cd80636f2c7e26ecae0a6b2288c1179669bcf262
from common import *
import concurrent.futures,time
RUN=OUT/('normal_overlap_'+uuid4().hex[:8]);RUN.mkdir();cases=[]
def probe():
    started=time.time();data,sup,path=task(RUN,'office-probe',workspace=read(OUT/'LIFECYCLE_FINAL.json')['workspace'])
    assert all(r['quit_returned'] and not r.get('forced_termination_attempted') for r in sup['instances'])
    return {'result':str(path),'supervisor':data['office_supervisor'],'started':started,'ended':time.time()}
for group in range(3):
    with concurrent.futures.ThreadPoolExecutor(2) as pool:
        futures=[pool.submit(probe) for _ in range(2)];cases += [f.result() for f in futures]
save(RUN/'RESULT.json',{'status':'PASS','cases':cases,'exe_sha256':sha(EXE)})
save(OUT/'NORMAL_OVERLAP_FINAL.json',{'status':'PASS','evidence':str(RUN/'RESULT.json'),'cases':len(cases),'exe_sha256':sha(EXE)})
