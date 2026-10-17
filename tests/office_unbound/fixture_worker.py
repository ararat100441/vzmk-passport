"""TEST ONLY: request an owned Excel, leave it unbound, publish worker PASS."""
import argparse,os,sys,time
from pathlib import Path
from uuid import uuid4
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage8_closure import storage as S
from openpyxl import Workbook

p=argparse.ArgumentParser();p.add_argument('--result',type=Path,required=True)
p.add_argument('--office-worker',type=Path,required=True)
args,_=p.parse_known_args();base=args.office_worker.parent
probe=base/'UNBOUND_TEST_PROBE.xlsx';book=Workbook();book.active['A1']='ARTIFICIAL UNBOUND EXCEL TEST';book.save(probe);book.close()
request=base/('launch_'+uuid4().hex+'.json')
save(request,{'progid':'Excel.Application','probe':str(probe),'probe_sha256':S.sha(probe),'worker_pid':os.getpid()})
response=request.with_suffix('.response.json');deadline=time.monotonic()+40
while not response.exists() and time.monotonic()<deadline:time.sleep(.05)
reply=S.read(response)
if reply['status']!='PASS':raise RuntimeError(reply)
pid=reply['record']['pid']
import win32gui,win32process
def ready():
    found=[]
    def top(hwnd,_):
        if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid:return
        def child(hwnd,_):
            if win32gui.GetClassName(hwnd)=='EXCEL7' and win32process.GetWindowThreadProcessId(hwnd)[1]==pid:found.append(hwnd)
        win32gui.EnumChildWindows(hwnd,child,None)
    win32gui.EnumWindows(top,None)
    return bool(found)
deadline=time.monotonic()+30
while not ready() and time.monotonic()<deadline:time.sleep(.1)
if not ready():raise RuntimeError('TEST_EXCEL_NATIVE_WINDOW_NOT_READY')
save(base/'UNBOUND_TEST_WORKER.json',{'worker_pid':os.getpid(),'office_pid':pid,'bound':False,'status':'PASS'})
save(args.result,{'status':'PASS','test':'SUCCESSFUL_WORKER_WITH_UNBOUND_OWNED_EXCEL'})
