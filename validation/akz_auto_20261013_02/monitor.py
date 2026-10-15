"""Read-only retained handles for Office present before integration tests."""
import json,time,sys,subprocess,os
from pathlib import Path
import win32api,win32con,win32event,win32process
OUT=Path(__file__).resolve().parent
def save(p,d):
    with p.open('x',encoding='utf8') as f:json.dump(d,f,ensure_ascii=False,indent=2,default=str)
cmd=['powershell.exe','-NoProfile','-NonInteractive','-Command',"Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='EXCEL.EXE'\" | Select-Object Name,ProcessId,ParentProcessId,CreationDate,ExecutablePath | ConvertTo-Json -Depth 3"]
data=subprocess.check_output(cmd,encoding='utf8',creationflags=subprocess.CREATE_NO_WINDOW);rows=json.loads(data) if data.strip() else [];rows=[rows] if isinstance(rows,dict) else rows
items=[]
for r in rows:
    h=win32api.OpenProcess(0x1000|win32con.SYNCHRONIZE,False,r['ProcessId']);r['creation_time']=win32process.GetProcessTimes(h)['CreationTime'].isoformat();items.append((h,r))
save(OUT/'OFFICE_BASELINE.json',{'at':time.time(),'processes':rows,'read_only_handles':True,'terminate_access':False});print('FOREIGN BASELINE MONITOR READY',len(items),flush=True)
try:
    while not (OUT/'MONITOR_STOP.json').exists():
        for h,r in items:
            if win32event.WaitForSingleObject(h,0)==win32con.WAIT_OBJECT_0:
                save(OUT/'FOREIGN_OFFICE_STOP.json',{'reason':'BASELINE_OFFICE_EXIT','identity':r,'exit_code':win32process.GetExitCodeProcess(h),'cause':'UNKNOWN'});raise RuntimeError('STOP_FOREIGN_OFFICE_EXIT')
        time.sleep(.25)
finally:
    alive=[dict(r,alive=win32event.WaitForSingleObject(h,0)!=win32con.WAIT_OBJECT_0) for h,r in items]
    save(OUT/'OFFICE_BASELINE_FINAL.json',{'processes':alive,'foreign_stop':(OUT/'FOREIGN_OFFICE_STOP.json').exists()})
    for h,r in items:h.Close()
