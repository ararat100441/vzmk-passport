"""Launch isolated frozen GUI / retain read-only exit handle. No GUI controls."""
from prepare import *
mode=sys.argv[1]
binding=read(OUT/'BINDING_BEFORE.json');exe=Path(binding['build']['exe'])
assert sha(exe)==binding['build']['exe_sha256']
if mode=='EXISTING':
    pid=47440;h=win32api.OpenProcess(win32con.SYNCHRONIZE|0x1000,False,pid)
    birth=win32process.GetProcessTimes(h)['CreationTime'].isoformat()
    assert datetime.fromisoformat(birth).timestamp()==datetime.fromisoformat('2026-10-12T10:11:58.452000+00:00').timestamp()
    save(OUT/'GUI_EXISTING_EXIT_WATCH.json',{'pid':pid,'creation_time':birth,'handle_access':'QUERY_LIMITED_INFORMATION|SYNCHRONIZE','terminate_access':False,'exe_sha256':sha(exe)})
else:
    base=OUT/'isolated_gui';(base/'VZMKPassport').mkdir(parents=True);(base/'tmp').mkdir()
    save(base/'VZMKPassport/settings.json',{'workspace':binding['workspace'],'version':'0.9.9'})
    env=os.environ.copy();env['LOCALAPPDATA']=str(base);env['TEMP']=env['TMP']=str(base/'tmp');env['PATH']=os.environ['SystemRoot']+'\\System32;'+os.environ['SystemRoot'];env.pop('PYTHONPATH',None);env.pop('PYTHONHOME',None)
    child=subprocess.Popen([str(exe)],env=env,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,creationflags=subprocess.CREATE_NO_WINDOW)
    pid=child.pid;h=child._handle;birth=win32process.GetProcessTimes(h)['CreationTime'].isoformat()
    save(OUT/'GUI_FRESH_LAUNCH.json',{'pid':pid,'creation_time':birth,'exe':str(exe),'exe_sha256':sha(exe),'LOCALAPPDATA':str(base),'PATH':env['PATH'],'python_in_path':None,'original_handle_source':'CREATE_PROCESS','gui_controls':'Computer Use sky only'})
print(mode,'GUI WATCH READY',pid,flush=True)
deadline=time.monotonic()+300
while win32event.WaitForSingleObject(h,500)!=win32con.WAIT_OBJECT_0:
    if time.monotonic()>deadline:raise TimeoutError('GUI_EXIT_NOT_OBSERVED_NO_TERMINATE')
exit_code=win32process.GetExitCodeProcess(h)
save(OUT/('GUI_'+mode+'_EXIT.json'),{'pid':pid,'creation_time':birth,'exit_code':exit_code,'exit_observed':True,'forced':False})
if mode!='EXISTING':child.returncode=exit_code  # Prevent Popen.__del__ querying an already closed handle.
h.Close();print(mode,'GUI EXIT PASS',flush=True)
