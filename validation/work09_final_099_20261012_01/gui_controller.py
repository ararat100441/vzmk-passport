"""Launch/watch own GUI and inspect own Office read-only; no UI input APIs."""
from common import *
import win32api,win32process,win32event,gc,pythoncom,shutil
from stage9_app.ownership import process_image,native_application
mode=sys.argv[1]
if mode=='LAUNCH':
    work=read(OUT/'EXE_E2E_FINAL.json')['workspace']
    cfg=Path(SETUP['base'])/'VZMKPassport/settings.json';cfg.parent.mkdir(parents=True,exist_ok=True)
    save(cfg,{'workspace':work,'version':'0.9.9'})
    fixture=Path.home()/'Documents'/('!'+Path(SETUP['parent']).name);fixture.mkdir()
    vom=fixture/'TEST_VOM.xls';shutil.copyfile(VOM,vom);assert sha(vom)==sha(VOM)
    save(OUT/'GUI_INPUTS.json',{'fixture_folder':str(fixture),'vom':str(vom),'vom_sha256':sha(vom),'workspace':work,'control_manifests':read(OUT/'EXE_E2E_FINAL.json')['manifests']})
    p=subprocess.Popen([str(EXE)],env=ENV,creationflags=subprocess.CREATE_NO_WINDOW,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    save(OUT/'GUI_LAUNCH.json',{'pid':p.pid,'creation_time':win32process.GetProcessTimes(p._handle)['CreationTime'].isoformat(),'exe':str(EXE),'exe_sha256':sha(EXE),'workspace':work,'LOCALAPPDATA':ENV['LOCALAPPDATA'],'PATH':ENV['PATH'],'python_in_path':None})
    print('Own independent GUI launched',p.pid,flush=True)
elif mode=='SNAPSHOT':
    label=sys.argv[2];launch=read(OUT/'GUI_LAUNCH.json');base=Path(SETUP['base'])/'VZMKPassport/diagnostics'
    dirs=sorted(base.glob('viewer_*'),key=lambda p:p.stat().st_ctime_ns);folder=dirs[-1]
    begin=read(next(folder.glob('ownership_OPEN_BEGIN_*.json')))
    paths={str(Path(p).resolve()).casefold():sha(p) for p in begin['paths']}
    starts=[(p,read(p)) for p in folder.glob('ownership_*_START.json')];assert len(starts)==2
    items=[];pythoncom.CoInitialize()
    try:
        for path,r in starts:
            h=win32api.OpenProcess(0x1000|0x100000,False,r['pid'])
            try:
                assert r['launcher_pid']==launch['pid'] and r['actual_child'] and r['owned'] and r['handle_source']=='CREATE_PROCESS'
                assert process_image(h)==r['image'] and win32process.GetProcessTimes(h)['CreationTime'].isoformat()==r['creation_time']
                docs=[]
                with native_application(r,h) as app:
                    collection=app.Documents if r['progid']=='Word.Application' else app.Workbooks
                    assert collection.Count==3
                    for i in range(1,collection.Count+1):
                        doc=collection(i);p=str(Path(doc.FullName).resolve());assert doc.ReadOnly and paths[p.casefold()]==sha(p)
                        docs.append({'path':p,'readonly':bool(doc.ReadOnly),'sha256':sha(p)});doc=None
                    collection=None
                app=None;gc.collect();items.append({'identity':r,'start_file':str(path),'documents':docs})
            finally:h.Close()
    finally:app=collection=None;gc.collect();pythoncom.CoUninitialize()
    save(OUT/('GUI_SNAPSHOT_'+label+'.json'),{'status':'PASS','viewer_id':folder.name,'open_attempt_id':begin['open_attempt_id'],'instances':items,'physical_files':6,'no_write_close_quit_by_controller':True})
    print('Read-only snapshot',label,'PASS 3 Word / 3 Excel',flush=True)
elif mode=='WINDOWS':
    import win32gui
    launch=read(OUT/'GUI_LAUNCH.json');base=Path(SETUP['base'])/'VZMKPassport/diagnostics'
    own={}
    for p in base.glob('viewer_*/ownership_*_START.json'):
        r=read(p)
        if r['launcher_pid']==launch['pid']:own[r['pid']]=r
    result=[]
    def window(hwnd,_):
        pid=win32process.GetWindowThreadProcessId(hwnd)[1]
        if pid in own and win32gui.IsWindowVisible(hwnd):result.append({'hwnd':hwnd,'pid':pid,'title':win32gui.GetWindowText(hwnd),'identity':own[pid]})
    win32gui.EnumWindows(window,None);print(json.dumps(result,ensure_ascii=False),flush=True)
elif mode=='EXIT':
    launch=read(OUT/'GUI_LAUNCH.json');h=win32api.OpenProcess(0x1000|0x100000,False,launch['pid'])
    try:
        assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==launch['creation_time'] and process_image(h)==str(EXE)
        save(OUT/'GUI_EXIT_WATCH.json',{'pid':launch['pid'],'creation_time':launch['creation_time'],'termination_access':False})
        deadline=time.monotonic()+1200
        while win32event.WaitForSingleObject(h,500)!=0:
            if time.monotonic()>deadline:raise TimeoutError('GUI_EXIT_NOT_OBSERVED_NO_TERMINATE')
        save(OUT/'GUI_EXIT.json',{'pid':launch['pid'],'creation_time':launch['creation_time'],'exit_code':win32process.GetExitCodeProcess(h),'forced':False})
    finally:h.Close()
