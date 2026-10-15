"""Retained read-only baseline Office handles; no COM or termination API."""
from common import *
import win32api,win32process,win32event
from stage9_app.ownership import process_image
handles=[];baseline=[]
for pid in win32process.EnumProcesses():
 if not pid:continue
 h=None
 try:
  h=win32api.OpenProcess(0x1000|0x100000,False,pid)
  if win32event.WaitForSingleObject(h,0)==0:continue
  image=process_image(h)
  if Path(image).name.upper() not in ('EXCEL.EXE','WINWORD.EXE'):continue
  r={'pid':pid,'creation_time':win32process.GetProcessTimes(h)['CreationTime'].isoformat(),'image':image,'exit_observed':False};baseline.append(r);handles.append(h);h=None
 except Exception:pass
 finally:
  if h is not None:h.Close()
save(OUT/'OFFICE_BASELINE.json',{'access':'QUERY_LIMITED_INFORMATION|SYNCHRONIZE','instances':baseline,'termination_API':False})
print('Baseline Office',len(baseline),flush=True)
try:
 while not (OUT/'MONITOR_STOP.json').exists():
  for h,r in zip(handles,baseline):
   if not r['exit_observed'] and win32event.WaitForSingleObject(h,0)==0:
    r.update(exit_observed=True,exit_code=win32process.GetExitCodeProcess(h),observed_unix=time.time())
    if not (OUT/'FOREIGN_OFFICE_STOP.json').exists():save(OUT/'FOREIGN_OFFICE_STOP.json',{'status':'STOP','cause':'UNKNOWN','identity':r,'no_termination_API':True})
  time.sleep(.25)
 save(OUT/'OFFICE_BASELINE_FINAL.json',{'instances':baseline,'termination_API':False,'all_same_identities_live':all(not x['exit_observed'] for x in baseline)})
finally:
 for h in handles:h.Close()
