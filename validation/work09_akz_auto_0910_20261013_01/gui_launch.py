from common import *
import win32api,win32process,win32event
mode=sys.argv[1];s=config();work=Path(read(OUT/'WORKSPACE.json')['workspace'])
if mode=='LAUNCH':
 assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists()
 path=Path(s['base'])/'VZMKPassport/settings.json';path.parent.mkdir(parents=True,exist_ok=True)
 if not path.exists():save(path,{'workspace':str(work),'version':'0.9.10'})
 before={str(p):sha(p) for p in work.rglob('*.xlsx') if 'Stage9_AKZ_Auto' in p.parts};save(OUT/'GUI_BEFORE_OUTPUT_HASHES.json',before)
 proc=subprocess.Popen([str(EXE)],env=environment(),stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
 birth=win32process.GetProcessTimes(proc._handle)['CreationTime'].isoformat()
 save(OUT/'GUI_LAUNCH.json',{'pid':proc.pid,'creation_time':birth,'image':str(EXE),'exe_sha256':sha(EXE),'workspace':str(work),'LOCALAPPDATA':s['base'],'PATH':environment()['PATH']})
 print('Own frozen GUI',proc.pid,birth,flush=True)
elif mode=='EXIT':
 r=read(OUT/'GUI_LAUNCH.json');h=win32api.OpenProcess(0x1000|0x100000,False,r['pid'])
 try:
  assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time']
  from stage9_app.ownership import process_image
  assert process_image(h).casefold()==r['image'].casefold()
  status=win32event.WaitForSingleObject(h,1200000)
  save(OUT/'GUI_EXIT.json',{'identity':r,'observed_exit':status==0,'exit_code':win32process.GetExitCodeProcess(h)})
 finally:h.Close()
