"""Independent bounded rereview. All writable data is beneath RUN, product untouched."""
import copy, ctypes, gc, hashlib, io, json, os, shutil, subprocess, sys, time, traceback
from contextlib import contextmanager
from pathlib import Path
from uuid import uuid4
from unittest.mock import patch
RUN=Path(__file__).resolve().parent; ROOT=RUN.parents[1]
sys.path.insert(0,str(ROOT));sys.dont_write_bytecode=True
from stage8_closure import core, native, office, storage as S, model
from stage8_closure.fixtures import create
import pythoncom, win32api, win32con, win32file, win32process, win32com.client
ENV=os.environ|{'PYTHONPATH':str(ROOT),'PYTHONUTF8':'1','PYTHONDONTWRITEBYTECODE':'1'}
def save(p,v):Path(p).write_text(json.dumps(v,ensure_ascii=False,indent=2),encoding='utf8')
def reg(r):return r/'registry/УЗК_TEST.xlsm'
def req(r):return r/'inputs/CLOSE_REQUEST.json'
def balances(r):return {k:v['before'] for k,v in model.registry_rows(reg(r),'STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ',{'s8-a','s8-b'}).items()}
def states(r):return [e['state'] for e in S.history(r)['events']]
def plan(r):return core.preview(r,req(r))
def close(r,fault=None):
    p=plan(r);return core.close(p,core.confirmation(p),_fault=fault)
def stopped(fn,contains=None):
    try:fn()
    except S.Stop as e:
        if contains:assert contains in str(e),(contains,str(e))
        return office._error(e)
    raise AssertionError('EXPECTED_SAFE_STOP')
def proc(*args,timeout=65):
    r=subprocess.run([sys.executable,'-B',str(__file__),*map(str,args)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=timeout)
    return {'exit':r.returncode,'stdout':r.stdout,'stderr':r.stderr}
def alive(pid):
    try:
        h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,pid)
        try:return win32process.GetExitCodeProcess(h)==259
        finally:h.Close()
    except Exception:return False
def creation(pid):
    h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,pid)
    try:return win32process.GetProcessTimes(h)['CreationTime'].isoformat()
    finally:h.Close()
def wait_exit(pid):
    end=time.monotonic()+25
    while alive(pid) and time.monotonic()<end:gc.collect();time.sleep(.1)
    return not alive(pid)
def terminate_owned(record):
    assert record['owned'] is True and creation(record['pid'])==record['creation_time']
    assert all(Path(p).is_relative_to(RUN) for p in record['expected_workbooks'])
    h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.PROCESS_TERMINATE,False,record['pid'])
    try:
        assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==record['creation_time']
        win32api.TerminateProcess(h,91)
    finally:h.Close()
def crash_cleanup(root):
    """Review-only cleanup; never assumes finally runs after os._exit."""
    results=[]
    for path in root.glob('transactions/*/office/*/START.json'):
        q=S.read(path)
        if not q.get('owned') or not alive(q['pid']) or creation(q['pid'])!=q['creation_time']:continue
        expected=q['expected_workbooks'];assert expected and all(Path(p).is_relative_to(root/'transactions') for p in expected)
        rot=pythoncom.GetRunningObjectTable();ctx=pythoncom.CreateBindCtx(0);book=None;excel=None;actual=[]
        for moniker in rot.EnumRunning():
            if moniker.GetDisplayName(ctx,None) in expected:
                book=win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch));excel=book.Application
                assert win32process.GetWindowThreadProcessId(excel.Hwnd)[1]==q['pid'] and not excel.Visible
                actual=[str(Path(b.FullName).resolve()) for b in excel.Workbooks];assert set(actual).issubset(expected)
                excel.EnableEvents=False;book.Close(SaveChanges=False);excel.Quit();break
        if excel is None:
            end=S.read(path.parent/'END.json')
            assert all(any(v['step']==s and v['status']=='PASS' for v in end['cleanup']) for s in ('Workbooks.Close','Quit'))
        book=None;excel=None;rot=None;ctx=None;gc.collect();exited=wait_exit(q['pid'])
        if not exited:terminate_owned(q)
        results.append({'ownership':q,'actual_workbooks':actual,'Quit_exited':exited,'review_forced_own_only':not exited,'final_exited':not alive(q['pid'])})
    return results
def publish(mode,src,dst):
    try:
        if mode=='rename':os.rename(dst,str(dst)+'.external_old');os.replace(src,dst)
        elif mode=='replace':os.replace(src,dst)
        elif mode=='write':
            with open(dst,'r+b') as f:f.write(b'EXTERNAL_WRITE');f.flush();os.fsync(f.fileno())
        elif mode=='parent':os.rename(dst,str(dst)+'.external_old')
        else:
            name=str(Path(dst).resolve())
            class Info(ctypes.Structure):
                _fields_=[('Flags',ctypes.c_uint32),('Root',ctypes.c_void_p),('Length',ctypes.c_uint32),('Name',ctypes.c_wchar*len(name))]
            info=Info(3,None,len(name.encode('utf-16-le')),name)
            fn=ctypes.WinDLL('kernel32',use_last_error=True).SetFileInformationByHandle
            fn.argtypes=[ctypes.c_void_p,ctypes.c_int,ctypes.c_void_p,ctypes.c_uint32];fn.restype=ctypes.c_int
            h=win32file.CreateFile(str(src),win32con.GENERIC_READ|win32con.DELETE,win32con.FILE_SHARE_READ|win32con.FILE_SHARE_DELETE,None,win32con.OPEN_EXISTING,0,None)
            try:
                if not fn(int(h),22,ctypes.byref(info),ctypes.sizeof(info)):raise ctypes.WinError(ctypes.get_last_error())
            finally:h.Close()
        return {'status':'PUBLISHED','mode':mode}
    except Exception as e:return {'status':'BLOCKED','mode':mode,**office._error(e),'winerror':getattr(e,'winerror',None)}
if len(sys.argv)>1 and sys.argv[1] in ('publisher','worker'):
    try:
        if sys.argv[1]=='publisher':
            r=publish(*sys.argv[2:5]);print(json.dumps(r,ensure_ascii=False));sys.exit(0 if r['status']=='PUBLISHED' else 2)
        cmd,path=sys.argv[2:4];path=Path(path)
        with office.diagnostic_scope(RUN/'worker_office'):
            if cmd=='recover':r=core.recover(path)
            elif cmd=='preview':r=plan(path)
            else:
                p=S.read(path)
                if cmd=='parallel':
                    gate=Path(sys.argv[4]);end=time.monotonic()+25
                    while not gate.exists():
                        if time.monotonic()>end:raise TimeoutError('GATE')
                        time.sleep(.02)
                def fault(point,target):
                    if cmd=='crash' and point==sys.argv[4]:os._exit(73)
                r=core.close(p,core.confirmation(p),_fault=fault)
        print(json.dumps(r,ensure_ascii=False));sys.exit(0)
    except Exception as e:print(json.dumps({'status':'STOP','error':office._error(e)},ensure_ascii=False));sys.exit(2)

SEED=None;EXTERNAL=None;RESULTS=[];PARTIAL=None
def fixture():
    r=RUN/'sandboxes'/uuid4().hex
    for name in ('registry','journal','backups','transactions','inputs'):(r/name).mkdir(parents=True)
    shutil.copyfile(reg(SEED),reg(r));m=S.read(SEED/'SANDBOX.json');m.pop('seal');m['sandbox_id']=r.name;m['seal']=S.digest(m);save(r/'SANDBOX.json',m)
    for p in (SEED/'inputs').glob('*.json'):
        q=S.read(p)
        if p.name=='CLOSE_REQUEST.json':
            for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['path']=str(r/'inputs'/Path(e['path']).name)
        save(r/'inputs'/p.name,q)
    q=S.read(req(r))
    for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:e['sha256']=S.sha(e['path'])
    save(req(r),q)
    return r
def amend(r,name,fn):
    p=r/'inputs'/name;q=S.read(p);fn(q);save(p,q)
    if name!='CLOSE_REQUEST.json':
        q=S.read(req(r))
        for e in [q['shipment'],q['requirements'],q['control'],*q['documents']]:
            if e['path']==str(p):e['sha256']=S.sha(p)
        save(req(r),q)
def reseed(r):
    m=S.read(r/'SANDBOX.json');m.pop('seal');m['initial_registry_sha256']=S.sha(reg(r));m['seal']=S.digest(m);save(r/'SANDBOX.json',m)
def case(name,fn):
    print('START',name,flush=True);start=time.monotonic()
    try:item={'name':name,'status':'PASS','detail':fn()}
    except Exception as e:item={'name':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    gc.collect();item['seconds']=round(time.monotonic()-start,3);RESULTS.append(item)
    save(RUN/'cases'/(name+'.json'),item)
    save(RUN/'RESULTS.json',{'PASS':sum(x['status']=='PASS' for x in RESULTS),'FAIL':sum(x['status']=='FAIL' for x in RESULTS),'results':RESULTS})
    assert {p:S.sha(p) for p in PRODUCT}==PRODUCT,'PRODUCT_CHANGED'
    print(item['status'],name,item.get('error',''),flush=True)
def prepare_seed():
    global SEED,EXTERNAL
    f=create(RUN/'sandboxes');SEED=Path(f['sandbox']);save(RUN/'SEED.json',f)
    EXTERNAL=RUN/'EXTERNAL_VALID_4.xlsm';shutil.copyfile(reg(SEED),EXTERNAL)
    with office.excel_session(workbooks=[EXTERNAL]) as e:
        b=e.Workbooks.Open(Filename=str(EXTERNAL),UpdateLinks=0,ReadOnly=False,CorruptLoad=0)
        try:b.Worksheets('УЗК').Range('C24').Value2=4;b.Save()
        finally:b.Close(SaveChanges=False)
    save(RUN/'EXTERNAL_NATIVE_4.json',native.inspect(EXTERNAL,{'УЗК!C24':4,'УЗК!C25':10}))
def race(mode,point):
    r=fixture();p=plan(r);inc=r/'transactions/external4.xlsm';shutil.copyfile(EXTERNAL,inc);attempts=[]
    def fault(phase,target):
        if phase==point:
            a=proc('publisher',mode,inc,reg(r).parent if mode=='parent' else reg(r));attempts.append(a)
            assert a['exit']==2,a
    result=core.close(p,core.confirmation(p),_fault=fault)
    assert attempts and result['status']=='CLOSED_TEST' and balances(r)=={'s8-a':9,'s8-b':6}
    assert S.file_identity(reg(r))==p['payload']['registry_identity']
    before=S.sha(reg(r));events=states(r)
    assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST' and S.sha(reg(r))==before and states(r)==events
    return {'sandbox':str(r),'initial':12,'external_intended':4,'final':9,'attempts':attempts,'journal':events,'identity_preserved':True}
def prior_publish():
    r=fixture();p=plan(r);inc=r/'transactions/external4.xlsm';shutil.copyfile(EXTERNAL,inc)
    a=proc('publisher','native',inc,reg(r));assert a['exit']==0,a
    h=S.sha(reg(r));e=stopped(lambda:core.close(p,core.confirmation(p)))
    assert balances(r)['s8-a']==4 and not states(r) and S.sha(reg(r))==h
    return {'sandbox':str(r),'publisher':a,'refusal':e,'balance_preserved':4,'journal':[]}
def identity_only():
    r=fixture();p=plan(r);inc=r/'transactions/same.xlsm';shutil.copyfile(reg(r),inc);os.replace(inc,reg(r))
    assert S.sha(reg(r))==p['payload']['before_sha256']
    e=stopped(lambda:core.close(p,core.confirmation(p)),'FILE_ID_CHANGED');assert not states(r)
    return {'sandbox':str(r),'same_sha':True,'different_identity':True,'refusal':e}
def parallel():
    r=fixture();p=plan(r);f=r/'PLAN.json';save(f,p);gate=r/'GO'
    workers=[subprocess.Popen([sys.executable,'-B',str(__file__),'worker','parallel',str(f),str(gate)],env=ENV,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf8') for _ in range(2)]
    gate.write_text('GO');out=[]
    for w in workers:
        stdout,stderr=w.communicate(timeout=65);out.append({'exit':w.returncode,'stdout':stdout,'stderr':stderr})
    assert sum(w['exit']==0 and json.loads(w['stdout'])['status']=='CLOSED_TEST' for w in out)==1,out
    assert states(r).count('COMMITTED')==1 and balances(r)=={'s8-a':9,'s8-b':6}
    return {'sandbox':str(r),'workers':out,'journal':states(r),'balances':balances(r)}
def preexisting_writer(mapping=False):
    r=fixture();p=plan(r);f=reg(r).open('r+b')
    if mapping:
        import mmap
        mm=mmap.mmap(f.fileno(),0,access=mmap.ACCESS_WRITE);f.close()
    try:e=stopped(lambda:core.close(p,core.confirmation(p)))
    finally:mm.close() if mapping else f.close()
    assert not states(r) and balances(r)['s8-a']==12
    return {'sandbox':str(r),'refusal':e,'mapping':mapping}
def positive():
    import olefile
    from zipfile import ZipFile
    def streams(p):
        with ZipFile(p) as z,olefile.OleFileIO(io.BytesIO(z.read('xl/vbaProject.bin'))) as ole:
            return {'/'.join(s):hashlib.sha256(ole.openstream(s).read()).hexdigest() for s in ole.listdir()}
    r=fixture();p=plan(r);result=core.close(p,core.confirmation(p));op=result['operation'];backup=Path(op['backup']['path'])
    before=native.inspect(backup,{'УЗК!C24':12,'УЗК!C25':10});after=native.inspect(reg(r),{'УЗК!C24':9,'УЗК!C25':6})
    assert streams(backup)==streams(reg(r)) and len(streams(backup))==11
    assert before['sheets']==after['sheets']
    assert {k:v['formula'] for k,v in before['formulas'].items()}=={k:v['formula'] for k,v in after['formulas'].items()}
    assert after['formulas']['S8_FORMULAS!$A$1']['value']==23 and after['formulas']['S8_FORMULAS!$A$2']['value']==5
    with ZipFile(backup) as a,ZipFile(reg(r)) as b:
        parts={'same_part_names':set(a.namelist())==set(b.namelist()),'styles_exact':a.read('xl/styles.xml')==b.read('xl/styles.xml'),'relationships_exact':a.read('xl/_rels/workbook.xml.rels')==b.read('xl/_rels/workbook.xml.rels')}
    assert all(parts.values()),parts
    ignored={(native.sheet_part(reg(r)),'C24'),(native.sheet_part(reg(r)),'C25')}
    assert native.package(backup,ignored)==native.package(reg(r),ignored)
    h=S.sha(reg(r));j=states(r);assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
    alias=r/'inputs/RENAMED_REQUEST.json';q=S.read(req(r));q['shipment_id']='OTHER';save(alias,q)
    assert core.preview(r,alias)['payload']['status']=='ALREADY_CLOSED_TEST'
    f=r/'PLAN.json';save(f,p);repeat=proc('worker','close',f)
    assert repeat['exit']==0 and json.loads(repeat['stdout'])['status']=='ALREADY_CLOSED_TEST'
    assert h==S.sha(reg(r)) and j==states(r) and j.count('COMMITTED')==1
    info={'sandbox':str(r),'before':before,'after':after,'streams':streams(reg(r)),'package':parts,'backup':op['backup'],'fresh_process_repeat':repeat,'identity_preserved':S.file_identity(reg(r))==p['payload']['registry_identity']}
    save(RUN/'NATIVE_PRESERVATION.json',info);return info
def cleanup(which,primary=False):
    calls=[];owned={};real=win32com.client.DispatchEx
    class Books:
        def __init__(self,b):self.b=b
        def __getattr__(self,k):return getattr(self.b,k)
        def Close(self):
            calls.append('Workbooks.Close')
            if which=='Workbooks.Close':raise pythoncom.com_error(-2147418111,'REVIEW_CLOSE_REJECTION',None,None)
            self.b.Close()
    class Excel:
        def __init__(self,e):object.__setattr__(self,'e',e)
        def __getattr__(self,k):
            if k=='Workbooks':return Books(self.e.Workbooks)
            if k=='Quit':
                def quit():
                    calls.append('Quit')
                    if which=='Quit':raise pythoncom.com_error(-2147418111,'REVIEW_QUIT_REJECTION',None,None)
                    return self.e.Quit()
                return quit
            return getattr(self.e,k)
        def __setattr__(self,k,v):
            if k=='CutCopyMode':
                calls.append('CutCopyMode')
                if which=='CutCopyMode':raise pythoncom.com_error(-2147418111,'REVIEW_CUT_COPY_REJECTION',None,None)
            setattr(self.e,k,v)
    def dispatch(*a,**k):
        e=real(*a,**k);owned['real']=e;owned['pid']=win32process.GetWindowThreadProcessId(e.Hwnd)[1];return Excel(e)
    directory=RUN/'cleanup'/uuid4().hex;original=S.Stop('REVIEW_PRIMARY_EXCEPTION');observed=None
    with patch('win32com.client.DispatchEx',side_effect=dispatch):
        try:
            with office.excel_session(diagnostics=directory,workbooks=[reg(SEED)]) as e:
                b=e.Workbooks.Open(Filename=str(reg(SEED)),UpdateLinks=0,ReadOnly=True,CorruptLoad=0)
                b.Close(SaveChanges=False);b=None
                if primary:raise original
        except S.Stop as exc:
            assert exc is original if primary else 'EXCEL_CLEANUP_FAILED' in str(exc)
            observed=office._error(exc)
        else:raise AssertionError('CLEANUP_FAILURE_HIDDEN')
    report=S.read(next(directory.glob('*/END.json')))
    assert calls==['CutCopyMode','Workbooks.Close','Quit'] and any(x['step']==which and x['status']=='ERROR' for x in report['cleanup'])
    pid=owned['pid']
    if which=='Quit':owned['real'].Quit()
    e=None;b=None;owned.clear();original.__traceback__=None;original=None;gc.collect()
    assert wait_exit(pid),'OWN_EXCEL_REMAINED_AFTER_CLEANUP'
    return {'calls':calls,'pid':pid,'exited':True,'primary_preserved':primary,'review_only_Quit_after_injection':which=='Quit','observed':observed,'diagnostics':report}
def book_close_primary():
    directory=RUN/'book_close'/uuid4().hex;original=pythoncom.com_error(-2147417848,'REVIEW_SAVE_DISCONNECTED',None,None)
    class Book:
        def Close(self,**kw):raise pythoncom.com_error(-2147418111,'REVIEW_BOOK_CLOSE_REJECTION',None,None)
    def body():
        with office.excel_session(diagnostics=directory):
            try:raise original
            finally:native._close_book(Book())
    e=stopped(body);assert e['cause']['hresult_hex']=='0x80010108'
    report=S.read(next(directory.glob('*/END.json')));assert any(x['step']=='Workbook.Close' and x['status']=='ERROR' for x in report['cleanup'])
    original.__traceback__=None;return {'refusal':e,'diagnostics':report}
def unrelated_excel():
    directory=RUN/'unrelated';pid=None
    with office.excel_session(diagnostics=directory) as e:
        pid=win32process.GetWindowThreadProcessId(e.Hwnd)[1];b=e.Workbooks.Add();b.Worksheets(1).Range('A1').Value2='REVIEW_USER_SENTINEL'
        with office.excel_session(diagnostics=RUN/'own_other') as other:
            assert win32process.GetWindowThreadProcessId(other.Hwnd)[1]!=pid
        assert alive(pid) and b.Worksheets(1).Range('A1').Value2=='REVIEW_USER_SENTINEL'
        b.Close(SaveChanges=False);b=None
    e=None;gc.collect();assert wait_exit(pid)
    return {'sentinel_pid':pid,'unrelated_instance_preserved':True,'review_cleaned_after_check':True}
def disconnect(exact=True):
    r=fixture();p=plan(r);actual=native.excel_session;captured={}
    class Book:
        def __init__(self,b):self.b=b
        def __getattr__(self,k):return getattr(self.b,k)
        def Save(self):raise pythoncom.com_error(-2147417848,'REVIEW_RPC_E_DISCONNECTED_SAVE',None,None)
    class Books:
        def __init__(self,b):self.b=b
        def Open(self,**kw):return Book(self.b.Open(**kw))
    class Excel:
        def __init__(self,e):self.e=e
        def __getattr__(self,k):return Books(self.e.Workbooks) if k=='Workbooks' else getattr(self.e,k)
    @contextmanager
    def session(**kw):
        with actual(**kw) as e:
            captured['excel']=e;yield Excel(e) if exact else e
    def fault(point,target):
        if not exact and point=='save_io':captured['excel'].Quit()
    with patch('stage8_closure.native.excel_session',side_effect=session):e=stopped(lambda:core.close(p,core.confirmation(p),_fault=fault),'MICROSOFT_EXCEL_FAILURE')
    captured.clear();gc.collect()
    assert states(r)==['INTENT','BACKUP_VERIFIED'] and S.sha(reg(r))==p['payload']['before_sha256']
    report=S.read(next(r.glob('transactions/*/office/*/END.json')))
    assert report['original_exception']['cause']['hresult_hex']=='0x80010108' if exact else report['original_exception']['cause']['hresult'] is not None
    op=S.history(r)['active'];S.verify_backup(op['backup'],r)
    recovery=proc('worker','recover',r);assert recovery['exit']==0 and json.loads(recovery['stdout'])['status']=='ABORTED_UNCHANGED'
    assert balances(r)['s8-a']==12 and states(r).count('COMMITTED')==0
    return {'sandbox':str(r),'exact_injection':exact,'real_Excel_Quit_disconnect':not exact,'refusal':e,'cleanup':report,'backup':op['backup'],'fresh_recovery':recovery}
def crash(point):
    global PARTIAL
    r=fixture();p=plan(r);f=r/'PLAN.json';save(f,p);worker=proc('worker','crash',f,point);assert worker['exit']==73,worker
    op=S.history(r)['active'];assert 'COMMITTED' not in states(r);S.verify_backup(op['backup'],r)
    cleanup=crash_cleanup(r)
    before=states(r);h=S.sha(reg(r));backup=native.inspect(op['backup']['path'],{'УЗК!C24':12,'УЗК!C25':10})
    recovery=proc('worker','recover',r)
    if point=='during_publish':
        assert h not in (op['before_sha256'],op['after_sha256']) and recovery['exit']==2 and 'AMBIGUOUS_OPERATION' in recovery['stdout']
        assert proc('worker','preview',r)['exit']==2 and proc('worker','close',f)['exit']==2
        count=[]
        with patch('win32com.client.DispatchEx',side_effect=lambda *a,**k:count.append(1)):
            package_failure=stopped(lambda:native.inspect(reg(r)))
        assert not count and S.sha(reg(r))==h and states(r)==before
        PARTIAL=r
        save(RUN/'PARTIAL_STATE.json',{'sandbox':str(r),'plan':str(f),'sha256':h,'operation':op})
        extra={'partial_detected':True,'native_inspect_rejects_before_COM':package_failure,'native_opens':len(count),'further_operations_blocked':True}
    else:
        assert recovery['exit']==0,recovery
        answer=json.loads(recovery['stdout']);assert answer['registry_writes']==0
        assert answer['status']==('CLOSED_TEST' if point=='after_replace' else 'ABORTED_UNCHANGED')
        assert balances(r)['s8-a']==(9 if point=='after_replace' else 12)
        extra={'recovery_result':answer}
    return {'sandbox':str(r),'worker':worker,'review_only_crash_cleanup':cleanup,'journal_before':before,'journal_after':states(r),'registry_sha_after_crash':h,'backup_native':backup,'backup':op['backup'],'fresh_recovery':recovery,**extra}
def replace_restore_attempt():
    r=PARTIAL;assert r is not None;op=S.history(r)['active'];before_id=S.file_identity(reg(r));S.verify_backup(op['backup'],r)
    saved=r/'DAMAGED_XLSM_EVIDENCE.bin';shutil.copyfile(reg(r),saved)
    incoming=r/'transactions/RESTORE_VERIFIED_BACKUP.xlsm';shutil.copyfile(op['backup']['path'],incoming);os.replace(incoming,reg(r))
    assert S.sha(reg(r))==op['before_sha256'] and S.file_identity(reg(r))!=before_id
    recovery=proc('worker','recover',r);assert recovery['exit']==2 and 'AMBIGUOUS_REGISTRY_FILE_ID' in recovery['stdout']
    assert 'COMMITTED' not in states(r)
    return {'sandbox':str(r),'verified_backup_restored_byte_for_byte':True,'old_identity':before_id,'new_identity':S.file_identity(reg(r)),'fresh_recovery':recovery,'journal_still_pending':True}
def recovery_acceptance():
    assert PARTIAL is not None
    r=PARTIAL;result=proc('worker','recover',r)
    gap={'sandbox':str(r),'fresh_recovery':result,'instruction':str(ROOT/'STAGE8_DEV.md'),'instruction_lines':'71-85','documented_file_restore_steps':False,'restore_command':False,'automatic_restore_required':False,'ordinary_verified_backup_replace_still_pending':S.history(r)['active'] is not None}
    save(RUN/'RECOVERY_GAP.json',gap)
    assert result['exit']==0 and S.history(r)['active'] is None, 'No documented/available operator workflow for partial XLSM; ordinary verified-backup replacement leaves identity STOP. Automatic restoration is not required. See RECOVERY_GAP.json.'
    return result
def expert_fixture_only_restore():
    crash_evidence=crash('during_publish');r=PARTIAL;op=S.history(r)['active'];before_identity=S.file_identity(reg(r))
    damaged=r/'DAMAGED_XLSM_EVIDENCE.bin';shutil.copyfile(reg(r),damaged)
    # This is independent fixture manipulation, NOT an exposed product recovery
    # operation or a proposed change to the product. It proves data is recoverable.
    with S.lock(r),S.registry_guard(reg(r)) as guard,S.frozen([op['backup']['path'],op['backup']['metadata'],*(r/'journal').glob('*.json')]):
        S.verify_backup(op['backup'],r)
        guard.publish(op['backup']['path'],op['plan']['registry_identity'],S.sha(reg(r)),op['before_sha256'])
    assert S.file_identity(reg(r))==before_identity and S.sha(reg(r))==op['before_sha256']
    recovery=proc('worker','recover',r);assert recovery['exit']==0 and json.loads(recovery['stdout'])['status']=='ABORTED_UNCHANGED'
    assert not S.history(r)['closed'] and balances(r)['s8-a']==12
    result=close(r);assert result['status']=='CLOSED_TEST' and balances(r)['s8-a']==9 and states(r).count('COMMITTED')==1
    return {'sandbox':str(r),'review_only_internal_guard_copy':True,'available_operator_product_workflow':False,'crash':crash_evidence,'fresh_recovery_after_review_manipulation':recovery,'new_explicit_close_once':True,'final_balances':balances(r)}
def corrupt_backup():
    r=fixture();p=plan(r)
    def fault(point,target):
        if point=='after_save':raise S.Stop('REVIEW_PAUSE')
    stopped(lambda:core.close(p,core.confirmation(p),_fault=fault),'REVIEW_PAUSE');op=S.history(r)['active']
    with Path(op['backup']['path']).open('ab') as f:f.write(b'DAMAGED_BACKUP')
    e=stopped(lambda:core.recover(r),'BACKUP_DAMAGED');assert balances(r)['s8-a']==12 and 'COMMITTED' not in states(r)
    return {'sandbox':str(r),'refusal':e,'registry_unchanged':True}
def corrupt_journal():
    r=fixture();close(r);h=S.sha(reg(r));(r/'journal/00000000.json').write_text('{BROKEN',encoding='utf8')
    e=stopped(lambda:plan(r),'UNREADABLE_JSON');rec=stopped(lambda:core.recover(r),'UNREADABLE_JSON')
    assert S.sha(reg(r))==h and balances(r)['s8-a']==9
    return {'sandbox':str(r),'preview_refusal':e,'recovery_refusal':rec,'no_repeat':True}
def changed_sha():
    r=fixture();p=plan(r);(r/'inputs/VIK.json').write_bytes((r/'inputs/VIK.json').read_bytes()+b' ')
    e=stopped(lambda:core.close(p,core.confirmation(p)));assert not states(r) and S.sha(reg(r))==p['payload']['before_sha256']
    return {'sandbox':str(r),'refusal':e}
def insufficient():
    r=fixture()
    for name in ('REQUIREMENTS.json','CONTROL.json'):amend(r,name,lambda q:q['uzk'][0].update(quantity=13))
    amend(r,'SHIPMENT.json',lambda q:q['rows'][0].update(quantity=13))
    e=stopped(lambda:plan(r),'INSUFFICIENT_UZK_BALANCE');assert not states(r) and balances(r)['s8-a']==12
    return {'sandbox':str(r),'refusal':e}
def no_uzk():
    r=fixture();amend(r,'REQUIREMENTS.json',lambda q:q.update(uzk_required=False,uzk=[]));amend(r,'CONTROL.json',lambda q:q.update(uzk=[]))
    amend(r,'CLOSE_REQUEST.json',lambda q:q.update(documents=[d for d in q['documents'] if d['kind']!='UZK']))
    before=S.sha(reg(r));out=close(r);assert out['status']=='CLOSED_TEST' and S.sha(reg(r))==before
    assert not list((r/'backups').rglob('*.xlsm')) and not list((r/'transactions').rglob('stage.xlsm'))
    return {'sandbox':str(r),'registry_unchanged':True,'journal':states(r)}
def dev_forbidden():
    r=fixture();amend(r,'CLOSE_REQUEST.json',lambda q:q.update(passport='9/2026'))
    e=stopped(lambda:plan(r),'PASSPORT_9_2026');assert not states(r)
    return {'sandbox':str(r),'refusal':e,'registry_writes':0}
def torn_stage():
    r=fixture();p=plan(r)
    def fault(point,target):
        if point=='during_save':target.write_bytes(b'REVIEW_TORN_NATIVE_SAVE')
    e=stopped(lambda:core.close(p,core.confirmation(p),_fault=fault),'XLSM_PACKAGE_UNREADABLE');assert S.sha(reg(r))==p['payload']['before_sha256']
    rec=core.recover(r);assert rec['status']=='ABORTED_UNCHANGED'
    return {'sandbox':str(r),'refusal':e,'recovery':rec}
def ordinary_owned_processes():
    gc.collect();records=[]
    for f in RUN.rglob('START.json'):
        q=S.read(f)
        if q.get('owned') and alive(q['pid']) and creation(q['pid'])==q['creation_time']:
            if not wait_exit(q['pid']):records.append({'start':str(f),**q})
    save(RUN/'OWNED_PROCESS_CHECK.json',{'live':records})
    assert not records,records
    return {'own_live_processes':0}
if __name__=='__main__':
    (RUN/'cases').mkdir(exist_ok=True);(RUN/'sandboxes').mkdir(exist_ok=True)
    PRODUCT={p:S.sha(p) for p in S.read(RUN/'PRODUCT_BEFORE.json')}
    with office.diagnostic_scope(RUN/'office'):
        prepare_seed()
        case('external_rename_then_replace_before_publish',lambda:race('rename','before_replace'))
        case('external_native_atomic_replace_before_publish',lambda:race('native','before_replace'))
        case('concurrent_valid_xlsm_balance_resurrection',lambda:race('native','before_replace'))
        case('cleanup_com_failure_must_still_attempt_quit',lambda:cleanup('CutCopyMode'))
        case('write_during_publish_blocked',lambda:race('write','during_publish'))
        case('native_during_publish_blocked',lambda:race('native','during_publish'))
        case('parent_rename_before_COMMITTED_blocked',lambda:race('parent','after_replace'))
        case('external_publish_before_acquisition_preserves_4',prior_publish)
        case('same_SHA_changed_file_identity_STOP',identity_only)
        case('two_simultaneous_closes_once',parallel)
        case('preexisting_writer_STOP',preexisting_writer)
        case('preexisting_writable_mapping_STOP',lambda:preexisting_writer(True))
        case('cleanup_Close_still_calls_Quit',lambda:cleanup('Workbooks.Close'))
        case('cleanup_Quit_error_recorded',lambda:cleanup('Quit'))
        case('cleanup_primary_exception_preserved',lambda:cleanup('Workbooks.Close',True))
        case('Workbook_Close_preserves_Save_hresult',book_close_primary)
        case('unrelated_Excel_preserved',unrelated_excel)
        case('RPC_E_DISCONNECTED_exact_injection',disconnect)
        case('real_Excel_Quit_disconnect_at_Save',lambda:disconnect(False))
        case('normal_native_preservation_debit_and_repeats',positive)
        case('death_before_Save_safe_abort',lambda:crash('before_save'))
        case('death_during_publication_fail_closed_backup_intact',lambda:crash('during_publish'))
        case('death_after_publication_commit_without_replay',lambda:crash('after_replace'))
        case('operator_replace_backup_attempt',replace_restore_attempt)
        case('partial_XLSM_supported_operator_recovery',recovery_acceptance)
        case('expert_fixture_restore_proves_backup_data_recoverable',expert_fixture_only_restore)
        case('damaged_verified_backup_STOP',corrupt_backup)
        case('damaged_journal_STOP_no_replay',corrupt_journal)
        case('changed_input_SHA_STOP',changed_sha)
        case('insufficient_balance_STOP',insufficient)
        case('no_UZK_no_registry_write',no_uzk)
        case('DEV_9_2026_blocked',dev_forbidden)
        case('torn_native_Save_result_STOP',torn_stage)
        case('ordinary_owned_processes_exited',ordinary_owned_processes)
    print('DONE',sum(x['status']=='PASS' for x in RESULTS),sum(x['status']=='FAIL' for x in RESULTS),flush=True)
