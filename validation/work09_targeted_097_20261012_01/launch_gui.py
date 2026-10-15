from common import *
import shutil,win32process
RUN=OUT/'gui';RUN.mkdir()
work,_,path=task(RUN,'import',source=SOURCE,parent=WORKSPACES)
vom=WORKSPACES/'TEST_VOM.xls';shutil.copyfile(VOM,vom);assert sha(vom)==sha(VOM)
cfg=Path(SETUP['base'])/'VZMKPassport/settings.json';save(cfg,{'workspace':work['workspace'],'version':'0.9.7'})
proc=subprocess.Popen([str(EXE)],env=ENV,creationflags=subprocess.CREATE_NO_WINDOW,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
save(OUT/'GUI_LAUNCH.json',{'exe':str(EXE),'exe_sha256':sha(EXE),'pid':proc.pid,'creation_time':win32process.GetProcessTimes(proc._handle)['CreationTime'].isoformat(),'workspace':work['workspace'],'vom':str(vom),'import_result':str(path),'LOCALAPPDATA':ENV['LOCALAPPDATA'],'PATH':ENV['PATH']})
print('GUI PID',proc.pid,'VOM',vom,flush=True)
