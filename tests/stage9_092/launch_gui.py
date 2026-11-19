from common import *
import win32process,time
p=subprocess.Popen([str(EXE)],env=ENV,creationflags=subprocess.CREATE_NO_WINDOW)
save(OUT/'GUI_LAUNCH_097.json',{'exe':str(EXE),'exe_sha256':sha(EXE),'pid':p.pid,
    'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),
    'path':ENV['PATH'],'python_in_path':None,'localappdata':ENV['LOCALAPPDATA']})
print('Launched own GUI',p.pid)
