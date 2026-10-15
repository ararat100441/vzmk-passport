from common import *
import time,win32process,win32event,traceback
from openpyxl import Workbook
from stage9_app.ownership import process_image,quit_owned_window
RUN=OUT/('direct_'+uuid4().hex[:8]);RUN.mkdir()
book=RUN/'ARTIFICIAL_BLANK.xlsx';b=Workbook();b.active['A1']='ARTIFICIAL STAGE9 CHILD PROCESS TEST';b.save(book);b.close()
image=Path(os.environ['ProgramFiles'])/'Microsoft Office/root/Office16/EXCEL.EXE'
p=subprocess.Popen([str(image),'/x',str(book)],creationflags=subprocess.CREATE_NO_WINDOW)
r={'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'image':process_image(p._handle),'hwnd':0,'parent_pid':os.getpid(),'actual_child':True}
save(RUN/'CHILD_STARTED.json',r)
try:
    time.sleep(5)
    result=quit_owned_window(r,p._handle,lambda phase,details:save(RUN/(phase+'_'+uuid4().hex[:8]+'.json'),{'phase':phase,'details':details}))
    save(RUN/'QUIT.json',result)
    try:p.wait(timeout=90)
    except subprocess.TimeoutExpired:
        save(RUN/'FORCE_BEGIN.json',r|{'valid_handle':int(p._handle),'image_rechecked':process_image(p._handle),'birth_rechecked':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat()})
        try:p.terminate();p.wait(timeout=10)
        except Exception as exc:save(RUN/'FORCE_FAILED.json',{'error':repr(exc)})
    save(RUN/'RESULT.json',r|{'exit_code':p.poll(),'still_running':p.poll() is None})
except Exception as exc:save(RUN/'FAILURE.json',{'error':repr(exc),'traceback':traceback.format_exc(),'child':r});raise
print(str(RUN),flush=True)
