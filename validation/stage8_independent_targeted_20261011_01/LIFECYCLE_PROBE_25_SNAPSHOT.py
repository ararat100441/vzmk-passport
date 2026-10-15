"""Fresh process, unwrapped native Excel inspect; remains alive after return."""
import gc,json,os,subprocess,sys,time
from pathlib import Path
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT))
from stage8_closure import native,office,storage as S
import win32api,win32con,win32process
def save(p,v):p.write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
OUT=RUN/'lifecycle_probe'
def alive(pid):
    try:
        h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,pid)
        try:return win32process.GetExitCodeProcess(h)==259
        finally:h.Close()
    except Exception:return False
if len(sys.argv)>1 and sys.argv[1]=='child':
    seed=Path(S.read(RUN/'SEED.json')['sandbox'])
    with office.diagnostic_scope(OUT/'office'):result=native.inspect(seed/'registry/УЗК_TEST.xlsm')
    result=None;gc.collect();save(OUT/'RETURNED.json',{'worker_pid':os.getpid(),'time':S.now()})
    deadline=time.monotonic()+55
    while not (OUT/'RELEASE').exists() and time.monotonic()<deadline:gc.collect();time.sleep(.1)
    sys.exit(0)
OUT.mkdir(exist_ok=False)
p=subprocess.Popen([sys.executable,'-B',str(__file__),'child'],env=os.environ|{'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1'},stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf8')
deadline=time.monotonic()+30
while not (OUT/'RETURNED.json').exists():
    assert p.poll() is None,'CHILD_FAILED'
    assert time.monotonic()<deadline,'CHILD_TIMEOUT'
    time.sleep(.1)
record=S.read(next((OUT/'office').glob('*/START.json')));end=S.read(next((OUT/'office').glob('*/END.json')))
assert record['owned'] and all(x['status']=='PASS' for x in end['cleanup'])
start=time.monotonic();deadline=start+25
while alive(record['pid']) and time.monotonic()<deadline:time.sleep(.1)
exited=not alive(record['pid']);client_alive=p.poll() is None
save(OUT/'BEFORE_CLIENT_EXIT.json',{'native_returned':True,'own_excel_exited':exited,'client_alive':client_alive,'seconds_after_return':round(time.monotonic()-start,2),'ownership':record,'cleanup':end})
(OUT/'RELEASE').write_text('GO');stdout,stderr=p.communicate(timeout=30)
deadline=time.monotonic()+25
while alive(record['pid']) and time.monotonic()<deadline:time.sleep(.1)
result={'name':'fresh_native_inspect_owned_Excel_exit_while_client_alive','status':'PASS' if exited and client_alive else 'FAIL','detail':{'native_returned':True,'own_excel_exited_before_client':exited,'client_alive_at_check':client_alive,'owned_pid':record['pid'],'exited_after_client':not alive(record['pid']),'worker_exit':p.returncode,'stdout':stdout,'stderr':stderr}}
save(OUT/'RESULT.json',result);print(json.dumps(result,ensure_ascii=False))
