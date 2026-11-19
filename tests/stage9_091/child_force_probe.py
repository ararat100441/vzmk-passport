from common import *
import time,win32process,traceback
from openpyxl import Workbook
from stage9_app.ownership import process_image
RUN=OUT/('child_force_'+uuid4().hex[:8]);RUN.mkdir()
book=RUN/'ARTIFICIAL_BLANK.xlsx';b=Workbook();b.active['A1']='ARTIFICIAL STAGE9 TERMINATION TEST';b.save(book);b.close()
image=Path(os.environ['ProgramFiles'])/'Microsoft Office/root/Office16/EXCEL.EXE'
p=subprocess.Popen([str(image),'/x',str(book)],creationflags=subprocess.CREATE_NO_WINDOW)
r={'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'image':process_image(p._handle),'parent_pid':os.getpid(),'actual_child':True,'handle':int(p._handle)}
save(RUN/'CHILD_STARTED.json',r)
time.sleep(5)
try:
    assert p.poll() is None and process_image(p._handle)==r['image'] and win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat()==r['creation_time']
    save(RUN/'FORCE_BEGIN.json',r|{'monotonic':time.monotonic()})
    p.terminate();p.wait(timeout=10)
    save(RUN/'RESULT.json',r|{'exit_code':p.returncode,'still_running':False,'monotonic':time.monotonic()})
except Exception as exc:save(RUN/'FAILURE.json',{'error':repr(exc),'traceback':traceback.format_exc(),'child':r,'monotonic':time.monotonic()});raise
print(str(RUN),flush=True)
