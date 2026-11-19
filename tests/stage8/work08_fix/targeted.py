"""Bounded Work08 fix validation; fresh TEST files, deterministic seams + processes."""
import gc,json,os,shutil,subprocess,sys,time,traceback,mmap
from pathlib import Path
from contextlib import contextmanager
from unittest.mock import patch
import pythoncom,win32api,win32con,win32process,win32com.client,win32file
from stage8_closure import core,storage as S,native,model,office
from stage8_closure.fixtures import create

ROOT=S.ROOT;RUN=ROOT/'validation/stage8_fix_20261010_01'
OUT=RUN/sys.argv[1];OUT.mkdir(exist_ok=False);(OUT/'cases').mkdir()
ENV=os.environ|{'PYTHONPATH':str(ROOT),'PYTHONUTF8':'1'}
HERE=Path(__file__).parent;RESULTS=[]
ONLY=set(sys.argv[2:])
PRODUCT=core.code_bindings();S.write(OUT/'PRODUCT_BEFORE.json',PRODUCT)
DELIVERY_PATHS=S.read(RUN/'PRODUCT_CHECKPOINT.json')
S.write(OUT/'DELIVERY_BEFORE.json',{p:S.sha(p) for p in DELIVERY_PATHS})

def save(p,v):S.write(p,v)
def make(uzk=True):
    f=create(RUN/'sandboxes',uzk=uzk);root=Path(f['sandbox']);reg=root/'registry/УЗК_TEST.xlsm'
    return f,root,reg
def balance(reg):return {k:v['before'] for k,v in model.registry_rows(reg,'STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ',{'s8-a','s8-b'}).items()}
def states(root):return [e['state'] for e in S.history(root)['events']]
def stopped(fn,contains=None):
    try:fn()
    except Exception as e:
        assert isinstance(e,S.Stop),('NOT_A_SAFE_STOP',type(e).__name__,str(e))
        if contains:assert contains in str(e),(contains,str(e))
        return {'type':type(e).__name__,'error':str(e),'cause':str(e.__cause__),'notes':getattr(e,'__notes__',[])}
    raise AssertionError('STOP_REQUIRED')
def runprocess(args):
    p=subprocess.run([sys.executable,'-B',*map(str,args)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=45)
    return {'exit':p.returncode,'stdout':p.stdout,'stderr':p.stderr}
def publisher(mode,src,reg):return runprocess([HERE/'publisher.py',mode,src,reg])
def case(name,fn):
    if ONLY and name not in ONLY:return
    start=time.monotonic()
    try:
        info=fn() or {};result={'case':name,'status':'PASS','seconds':round(time.monotonic()-start,3),**info}
    except Exception as e:
        result={'case':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
    assert core.code_bindings()==PRODUCT,'PRODUCT_CHANGED_DURING_TESTS'
    RESULTS.append(result);save(OUT/'cases'/(name+'.json'),result)
    print(result['status'],name,result.get('error',''),flush=True)

def incoming4(reg,root):
    p=root/'transactions/EXTERNAL_4.xlsm';shutil.copyfile(reg,p)
    with office.excel_session(workbooks=[p]) as excel:
        b=excel.Workbooks.Open(Filename=str(p),UpdateLinks=0,ReadOnly=False,AddToMru=False,CorruptLoad=0)
        try:b.Worksheets('УЗК').Cells(24,3).Value2=4;b.Save()
        finally:b.Close(SaveChanges=False)
    return p,native.inspect(p,{'УЗК!C24':4})

def races(mode,phase):
    f,root,reg=make();p=core.preview(root,f['request']);inc,n=incoming4(reg,root)
    attempts=[]
    def fault(point,target):
        if point==phase:
            target=reg.parent if mode=='parent' else reg
            r=publisher('rename' if mode=='parent' else mode,inc,target);attempts.append(r)
            assert r['exit']==2,r
    r=core.close(p,core.confirmation(p),_fault=fault)
    assert attempts and r['status']=='CLOSED_TEST' and balance(reg)=={'s8-a':9,'s8-b':6}
    before=S.sha(reg);assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
    assert S.sha(reg)==before and states(root).count('COMMITTED')==1
    return {'sandbox':str(root),'initial':12,'external_intended':4,'final':9,
            'external_native':n,'attempts':attempts,'journal':states(root),
            'target_identity':S.file_identity(reg),'identity_preserved':S.file_identity(reg)==p['payload']['registry_identity'],
            'backup':S.history(root)['events'][1]['backup'],'final_sha':before}

def prior_replace(mode):
    f,root,reg=make();p=core.preview(root,f['request']);inc,n=incoming4(reg,root)
    r=publisher(mode,inc,reg);assert r['exit']==0,r
    after=S.sha(reg);error=stopped(lambda:core.close(p,core.confirmation(p)))
    assert balance(reg)['s8-a']==4 and S.sha(reg)==after and not states(root)
    return {'sandbox':str(root),'error':error,'external_process':r,'final_balance':4,'final_sha':after,'journal':states(root)}

def identity_only():
    f,root,reg=make();p=core.preview(root,f['request']);inc=root/'transactions/same.xlsm'
    shutil.copyfile(reg,inc);old_id=S.file_identity(reg);os.replace(inc,reg)
    assert S.sha(reg)==p['payload']['before_sha256'] and S.file_identity(reg)!=old_id
    e=stopped(lambda:core.close(p,core.confirmation(p)),'FILE_ID_CHANGED')
    assert not states(root) and balance(reg)['s8-a']==12
    return {'sandbox':str(root),'error':e,'old_id':old_id,'new_id':S.file_identity(reg),'same_sha':S.sha(reg)}

def changed_input():
    f,root,reg=make();p=core.preview(root,f['request']);doc=root/'inputs/VIK.json'
    with doc.open('ab') as o:o.write(b' ')
    e=stopped(lambda:core.close(p,core.confirmation(p)))
    assert not states(root) and S.sha(reg)==p['payload']['before_sha256']
    return {'sandbox':str(root),'error':e,'journal':states(root)}

def process_creation(pid):
    h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,pid)
    try:return win32process.GetProcessTimes(h)['CreationTime'].isoformat()
    finally:h.Close()
def alive(pid):
    try:
        h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,pid)
        try:return win32process.GetExitCodeProcess(h)==259
        finally:h.Close()
    except Exception:return False
def wait_exit(pid):
    end=time.monotonic()+25
    while alive(pid) and time.monotonic()<end:time.sleep(.05)
    return not alive(pid)

def matching_record(path):
    r=S.read(path)
    try:return alive(r['pid']) and process_creation(r['pid'])==r['creation_time']
    except Exception:return False

def owned_record(root):
    records=[p for p in root.glob('transactions/*/office/*/START.json') if matching_record(p)]
    assert records,records
    records.sort(key=lambda p:(p.parent/'END.json').exists())
    p=records[0];r=S.read(p);assert r['owned'] is True
    assert r['creation_time']==process_creation(r['pid'])
    assert all(Path(b).is_relative_to(root/'transactions') for b in r['expected_workbooks'])
    return p,r

def terminate_owned(record):
    assert record['owned'] and process_creation(record['pid'])==record['creation_time']
    h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION|win32con.PROCESS_TERMINATE,False,record['pid'])
    try:
        assert win32process.GetProcessTimes(h)['CreationTime'].isoformat()==record['creation_time']
        win32api.TerminateProcess(h,91)
    finally:h.Close()
    assert wait_exit(record['pid'])

def orphan_cleanup(root):
    p,r=owned_record(root)
    before={'ownership_log':str(p),**r};pid=r['pid'];actual=[];excel=None;book=None
    rot=pythoncom.GetRunningObjectTable();context=pythoncom.CreateBindCtx(0)
    for moniker in rot.EnumRunning():
        name=moniker.GetDisplayName(context,None)
        if name in r['expected_workbooks']:
            book=win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
            excel=book.Application
            assert win32process.GetWindowThreadProcessId(excel.Hwnd)[1]==pid and not excel.Visible
            actual=[str(Path(b.FullName).resolve()) for b in excel.Workbooks]
            assert actual and set(actual).issubset(r['expected_workbooks']),actual
            excel.EnableEvents=False;book.Close(SaveChanges=False);excel.Quit();break
    if excel is None:
        # Only a PID with our completed Close+Quit ledger, never an arbitrary Excel.
        end=S.read(p.parent/'END.json')
        assert all(any(s['step']==step and s['status']=='PASS' for s in end['cleanup'])
                   for step in ('Workbooks.Close','Quit'))
    book=None;excel=None;gc.collect();exited=wait_exit(pid)
    if not exited:terminate_owned(r)
    return {'record':before,'actual_workbooks':actual,'Quit_exited':exited,
            'forced_owned_only':not exited,'final_exited':not alive(pid)}

def crash(phase,event=False):
    f,root,reg=make();p=core.preview(root,f['request']);plan=root/'PLAN.json';save(plan,p)
    command='save_event_crash' if event else 'crash'
    r=runprocess([HERE/'worker.py',command,plan,phase]);assert r['exit']==73,r
    journal=states(root);assert 'COMMITTED' not in journal
    cleanup=[]
    while any(matching_record(p) for p in root.glob('transactions/*/office/*/START.json')):
        cleanup.append(orphan_cleanup(root))
    if event:assert (root/'SAVE_CALLBACK_ENTERED.json').exists()
    op=S.history(root)['active'];S.verify_backup(op['backup'],root)
    assert S.sha(op['backup']['path'])==p['payload']['before_sha256']
    current=S.sha(reg)
    if phase=='during_publish':
        assert current not in (op['before_sha256'],op['after_sha256'])
        recovered=stopped(lambda:core.recover(root),'AMBIGUOUS_OPERATION')
        stopped(lambda:core.close(p,core.confirmation(p)),'INCOMPLETE_OPERATION')
        stopped(lambda:core.preview(root,f['request']),'INCOMPLETE_OPERATION')
        assert states(root)==journal and S.sha(reg)==current
    else:
        recovered=core.recover(root)
        if phase=='after_replace':
            assert recovered['status']=='CLOSED_TEST' and recovered['registry_writes']==0
            assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
            assert balance(reg)['s8-a']==9 and states(root).count('COMMITTED')==1
        else:
            assert recovered['status']=='ABORTED_UNCHANGED' and balance(reg)['s8-a']==12
        assert S.sha(reg)==current
    return {'sandbox':str(root),'process':r,'journal_before_recovery':journal,
            'journal_after_recovery':states(root),'recovery':recovered,'sha_after_crash':current,
            'backup':op['backup'],'orphan_cleanup':cleanup,'native_save_callback':event}

def parallel():
    f,root,reg=make();p=core.preview(root,f['request']);plan=root/'PLAN.json';save(plan,p)
    gate=root/'GO';children=[subprocess.Popen([sys.executable,'-B',str(HERE/'worker.py'),'parallel',str(plan),str(gate)],
        env=ENV,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf8') for _ in range(2)]
    gate.write_text('GO');result=[]
    for c in children:
        out,err=c.communicate(timeout=45);result.append({'exit':c.returncode,'stdout':out,'stderr':err})
    assert sum(r['exit']==0 and 'CLOSED_TEST' in r['stdout'] for r in result)==1,result
    assert balance(reg)=={'s8-a':9,'s8-b':6} and states(root).count('COMMITTED')==1
    return {'sandbox':str(root),'workers':result,'journal':states(root),'final_sha':S.sha(reg)}

def disconnect():
    f,root,reg=make();p=core.preview(root,f['request']);ownership=[];handshake=[];captured={}
    realdispatch=win32com.client.DispatchEx
    def dispatch(*a,**k):
        excel=realdispatch(*a,**k);captured['excel']=excel
        return excel
    def fault(point,target):
        if point=='save_io':
            _,r=owned_record(root);ownership.append(r);terminate_owned(r)
            # A dead COM server can first return RPC_S_CALL_FAILED; force its
            # disconnection handshake before Save, recording that real HRESULT.
            try:captured['excel'].Visible
            except Exception as exc:handshake.append(office._error(exc))
    with patch('win32com.client.DispatchEx',side_effect=dispatch):
        e=stopped(lambda:core.close(p,core.confirmation(p),_fault=fault),'MICROSOFT_EXCEL_FAILURE')
    captured.clear();gc.collect()
    assert states(root)==['INTENT','BACKUP_VERIFIED'] and S.sha(reg)==p['payload']['before_sha256']
    diagnostic=list(root.glob('transactions/*/office/*/END.json'));assert len(diagnostic)==1
    report=S.read(diagnostic[0]);assert [s['step'] for s in report['cleanup']][-3:]==['CutCopyMode','Workbooks.Close','Quit']
    assert report['original_exception']['cause']['hresult_hex'] in ('0x80010108','0x800706ba','0x800706be')
    recovered=core.recover(root);assert recovered['status']=='ABORTED_UNCHANGED'
    return {'sandbox':str(root),'exception':e,'ownership':ownership,'disconnect_handshake':handshake,'cleanup_diagnostic':report,'recovery':recovered,
            'native_original':native.inspect(reg)}

def exact_disconnect_injection():
    f,root,reg=make();p=core.preview(root,f['request']);actual_session=native.excel_session
    class Book:
        def __init__(self,b):self.book=b
        def __getattr__(self,k):return getattr(self.book,k)
        def Save(self):raise pythoncom.com_error(-2147417848,'INJECTED_RPC_E_DISCONNECTED_SAVE',None,None)
    class Books:
        def __init__(self,b):self.books=b
        def Open(self,**kwargs):return Book(self.books.Open(**kwargs))
    class Excel:
        def __init__(self,e):self.excel=e
        def __getattr__(self,k):return Books(self.excel.Workbooks) if k=='Workbooks' else getattr(self.excel,k)
    @contextmanager
    def session(**kwargs):
        with actual_session(**kwargs) as e:yield Excel(e)
    with patch('stage8_closure.native.excel_session',side_effect=session):
        e=stopped(lambda:core.close(p,core.confirmation(p)),'MICROSOFT_EXCEL_FAILURE')
    assert '-2147417848' in e['error']
    assert states(root)==['INTENT','BACKUP_VERIFIED'] and S.sha(reg)==p['payload']['before_sha256']
    report=S.read(next(root.glob('transactions/*/office/*/END.json')))
    assert report['original_exception']['cause']['hresult_hex']=='0x80010108'
    assert [s['step'] for s in report['cleanup']]==['CutCopyMode','Workbooks.Close','Quit']
    assert all(s['status']=='PASS' for s in report['cleanup'])
    S.verify_backup(S.history(root)['active']['backup'],root)
    recovered=core.recover(root);assert recovered['status']=='ABORTED_UNCHANGED'
    return {'sandbox':str(root),'fault_injection':True,'real_server_death':False,'exception':e,
            'cleanup_diagnostic':report,'recovery':recovered,'native_original':native.inspect(reg)}

def cleanup(which,body_error=False):
    calls=[];owned={};realdispatch=win32com.client.DispatchEx
    class Books:
        def __init__(self,real):self.real=real
        def Close(self):
            calls.append('Workbooks.Close')
            if which=='Workbooks.Close':raise pythoncom.com_error(-2147418111,'INJECTED_BOOKS_CLOSE',None,None)
            return self.real.Close()
    class Proxy:
        def __init__(self,real):object.__setattr__(self,'real',real)
        def __getattr__(self,name):
            if name=='Workbooks':return Books(self.real.Workbooks)
            if name=='Quit':
                def quit():
                    calls.append('Quit')
                    if which=='Quit':raise pythoncom.com_error(-2147418111,'INJECTED_QUIT',None,None)
                    return self.real.Quit()
                return quit
            return getattr(self.real,name)
        def __setattr__(self,name,value):
            if name=='CutCopyMode':
                calls.append('CutCopyMode')
                if which=='CutCopyMode':raise pythoncom.com_error(-2147418111,'INJECTED_CUT_COPY',None,None)
            setattr(self.real,name,value)
    def dispatch(*a,**k):
        e=realdispatch(*a,**k);owned['excel']=e;owned['pid']=win32process.GetWindowThreadProcessId(e.Hwnd)[1]
        return Proxy(e)
    directory=OUT/'cleanup'/S.digest([which,body_error])[:12]
    original=S.Stop('ORIGINAL_TEST_EXCEPTION')
    try:
        with patch('win32com.client.DispatchEx',side_effect=dispatch):
            try:
                with office.excel_session(diagnostics=directory):
                    if body_error:raise original
            except S.Stop as e:
                if body_error:assert e is original
                else:assert 'EXCEL_CLEANUP_FAILED' in str(e)
            else:raise AssertionError('CLEANUP_ERROR_MUST_STOP')
        assert calls==['CutCopyMode','Workbooks.Close','Quit'],calls
        p=next(directory.glob('*/END.json'));report=S.read(p)
        assert any(s['step']==which and s['status']=='ERROR' for s in report['cleanup'])
        pid=owned['pid']
        if which=='Quit':owned['excel'].Quit()
        owned.pop('excel',None);gc.collect()
        assert wait_exit(pid)
        return {'calls':calls,'diagnostic':report,'original_preserved':body_error,'own_pid':owned['pid']}
    finally:
        if 'excel' in owned:
            owned.pop('excel').Quit();gc.collect();assert wait_exit(owned['pid'])

def native_close_original():
    class Book:
        def Close(self,**k):raise pythoncom.com_error(-2147418111,'BOOK_CLOSE_ORIGINAL',None,None)
    original=pythoncom.com_error(-2147417848,'SAVE_DISCONNECTED',None,None)
    directory=OUT/'book_original'
    try:
        with office.excel_session(diagnostics=directory):
            try:raise original
            finally:native._close_book(Book())
    except S.Stop as e:assert e.__cause__ is original
    else:raise AssertionError('STOP_REQUIRED')
    r=S.read(next(directory.glob('*/END.json')))
    assert r['original_exception']['cause']['hresult_hex']=='0x80010108'
    assert any(x['step']=='Workbook.Close' and x['status']=='ERROR' for x in r['cleanup'])
    return {'diagnostic':r,'original_hresult_preserved':True}

def preservation():
    f,root,reg=make();p=core.preview(root,f['request']);seed=native.inspect(reg);oldid=S.file_identity(reg)
    r=core.close(p,core.confirmation(p));closed=S.sha(reg)
    check=native.inspect(reg,{'УЗК!C24':9,'УЗК!C25':6,'УЗК!C26':8})
    assert check['vba_streams']==seed['vba_streams'] and len(check['vba_streams'])==11
    assert check['sheets']==seed['sheets'] and S.file_identity(reg)==oldid
    assert {k:v['formula'] for k,v in check['formulas'].items()}=={k:v['formula'] for k,v in seed['formulas'].items()}
    assert check['formulas']['S8_FORMULAS!$A$1']['value']==23
    assert S.sha(reg)==closed and core.recover(root)['status']=='NO_PENDING_OPERATION'
    op=r['operation'];S.verify_backup(op['backup'],root)
    proof=S.read(op['preservation']['path']);assert proof['all_package_parts_preserved']
    return {'sandbox':str(root),'input_sha':p['payload']['before_sha256'],'output_sha':closed,
            'native_before':seed,'native_after':check,'preservation':proof,'journal':states(root)}

def no_uzk():
    f,root,reg=make(False);p=core.preview(root,f['request']);before=S.sha(reg);identity=S.file_identity(reg)
    r=core.close(p,core.confirmation(p));assert r['status']=='CLOSED_TEST'
    assert S.sha(reg)==before and S.file_identity(reg)==identity and not list((root/'backups').iterdir())
    assert core.close(p,core.confirmation(p))['status']=='ALREADY_CLOSED_TEST'
    return {'sandbox':str(root),'registry_unchanged':True,'journal':states(root)}

def dev_block():
    f,root,reg=make();req=Path(f['request']);q=S.read(req);q['passport']='9/2026'
    req.write_text(json.dumps(q,ensure_ascii=False),encoding='utf8')
    e=stopped(lambda:core.preview(root,req),'PASSPORT_9_2026_DEV_TEST_FORBIDDEN')
    assert not states(root) and S.sha(reg)==S.read(root/'SANDBOX.json')['initial_registry_sha256']
    return {'sandbox':str(root),'error':e,'registry_writes':0}

def corrupt_backup():
    f,root,reg=make();p=core.preview(root,f['request']);before=S.sha(reg)
    def fault(point,target):
        if point=='after_backup':
            with target.open('ab') as o:o.write(b'CORRUPT')
    e=stopped(lambda:core.close(p,core.confirmation(p),_fault=fault),'BACKUP_DAMAGED')
    assert S.sha(reg)==before and states(root)==['INTENT']
    return {'sandbox':str(root),'error':e,'journal':states(root)}

def torn_stage():
    f,root,reg=make();p=core.preview(root,f['request'])
    def fault(point,target):
        if point=='during_save':target.write_bytes(b'TORN_NATIVE_SAVE_TEST')
    e=stopped(lambda:core.close(p,core.confirmation(p),_fault=fault),'XLSM_PACKAGE_UNREADABLE')
    assert S.sha(reg)==p['payload']['before_sha256'] and 'COMMITTED' not in states(root)
    r=core.recover(root);assert r['status']=='ABORTED_UNCHANGED'
    return {'sandbox':str(root),'error':e,'recovery':r}

def held_writer():
    f,root,reg=make();p=core.preview(root,f['request'])
    h=win32file.CreateFile(str(reg),win32con.GENERIC_READ|win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE,None,win32con.OPEN_EXISTING,0,None)
    try:e=stopped(lambda:core.close(p,core.confirmation(p)))
    finally:h.Close()
    assert not states(root) and balance(reg)['s8-a']==12
    return {'sandbox':str(root),'error':e,'journal':states(root)}

def held_mapping():
    f,root,reg=make();p=core.preview(root,f['request'])
    file=reg.open('r+b');mapping=mmap.mmap(file.fileno(),0);file.close()
    try:e=stopped(lambda:core.close(p,core.confirmation(p)),'GUARD_OR_IO_FAILURE')
    finally:mapping.close()
    assert not states(root) and balance(reg)['s8-a']==12
    return {'sandbox':str(root),'error':e,'journal':states(root)}

def other_excel_preserved():
    sentinel=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(sentinel.Hwnd)[1]
    sentinel.Visible=False;sentinel.DisplayAlerts=False
    b=sentinel.Workbooks.Add();b.Worksheets(1).Range('A1').Value2='UNRELATED TEST USER EXCEL'
    try:
        with office.excel_session(diagnostics=OUT/'other_excel') as own:
            other=win32process.GetWindowThreadProcessId(own.Hwnd)[1];assert other!=pid
        own=None;gc.collect()
        assert wait_exit(other) and alive(pid) and b.Worksheets(1).Range('A1').Value2=='UNRELATED TEST USER EXCEL'
        return {'unrelated_test_pid':pid,'own_pid':other,'unrelated_preserved':True}
    finally:b.Close(SaveChanges=False);sentinel.Quit();b=None;sentinel=None;gc.collect();assert wait_exit(pid)

with office.diagnostic_scope(OUT/'office'):
    for mode in ('rename','native','replace'):
        case('work08_'+mode+'_before_publish',lambda m=mode:races(m,'before_replace'))
    for phase in ('after_intent','during_publish','after_replace'):
        for mode in ('native','write','parent'):
            case(mode+'_'+phase,lambda m=mode,p=phase:races(m,p))
    case('prior_native_replace_preserves_4',lambda:prior_replace('native'))
    case('same_sha_changed_file_identity',identity_only)
    case('changed_input_sha',changed_input)
    case('preexisting_writer',held_writer)
    case('preexisting_writable_mapping',held_mapping)
    case('two_simultaneous_closes',parallel)
    for phase in ('before_save','during_save','before_replace','during_publish','after_replace'):
        case('crash_'+phase,lambda p=phase:crash(p))
    case('real_native_Save_callback_process_death',lambda:crash('save_io',event=True))
    for step in ('CutCopyMode','Workbooks.Close','Quit'):
        case('cleanup_'+step.replace('.','_'),lambda s=step:cleanup(s))
    case('cleanup_preserves_original',lambda:cleanup('CutCopyMode',True))
    case('book_close_preserves_Save_hresult',native_close_original)
    case('unrelated_Excel_preserved',other_excel_preserved)
    case('real_COM_server_death_Save',disconnect)
    case('RPC_E_DISCONNECTED_Save_injection',exact_disconnect_injection)
    case('native_normal_reopen_XLSM_preservation',preservation)
    case('no_UZK_no_registry_write',no_uzk)
    case('DEV_9_2026_forbidden',dev_block)
    case('damaged_backup_STOP',corrupt_backup)
    case('torn_native_save_STOP',torn_stage)

save(OUT/'PRODUCT_AFTER.json',core.code_bindings())
save(OUT/'DELIVERY_AFTER.json',{p:S.sha(p) for p in DELIVERY_PATHS})
save(OUT/'RESULTS.json',{'results':RESULTS,'PASS':sum(r['status']=='PASS' for r in RESULTS),
                        'FAIL':sum(r['status']=='FAIL' for r in RESULTS)})
print('TOTAL',len(RESULTS),'FAIL',sum(r['status']=='FAIL' for r in RESULTS),flush=True)
sys.exit(1 if any(r['status']=='FAIL' for r in RESULTS) else 0)
