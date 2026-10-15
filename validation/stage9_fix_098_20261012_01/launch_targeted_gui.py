from common import *
import win32process
old=read(OUT/'OLD_GUI_LAUNCH.json')
cfg=Path(SETUP['base'])/'VZMKPassport/settings.json';cfg.parent.mkdir(parents=True);save(cfg,{'workspace':old['read_only_workspace'],'version':'0.9.9'})
p=subprocess.Popen([str(EXE)],env=ENV,creationflags=subprocess.CREATE_NO_WINDOW,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
save(OUT/'TARGETED_GUI_LAUNCH.json',{'exe':str(EXE),'exe_sha256':sha(EXE),'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'workspace':old['read_only_workspace'],'manifest':old['manifest'],'LOCALAPPDATA':ENV['LOCALAPPDATA'],'PATH':ENV['PATH']})
print('Targeted GUI',p.pid,flush=True)
