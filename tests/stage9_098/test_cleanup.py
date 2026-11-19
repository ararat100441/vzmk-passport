"""Targeted source contracts. Real frozen/GUI Office checks are separate."""
import concurrent.futures,json,os,sys,tempfile,types,unittest
from contextlib import ExitStack,contextmanager
from pathlib import Path
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app import diagnostics as D
from stage9_app.ownership import owned_office
from stage9_app.office import Viewer
from stage9_app.gui import App

class CleanupTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='stage9_098_',dir=os.environ.get('VZMK_TEST_PARENT'))
        self.root=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()
    def read(self,p):return json.loads(Path(p).read_text(encoding='utf8'))
    def test_empty_guard_single_final_write_and_repeated_attempts(self):
        import win32com.client,pythoncom
        originals=(win32com.client.DispatchEx,pythoncom.CoInitialize,pythoncom.CoUninitialize)
        for i in range(3):
            p=self.root/f'guard{i}.json'
            with owned_office(p):pass
            self.assertEqual(self.read(p)['status'],'PASS')
        self.assertEqual(originals,(win32com.client.DispatchEx,pythoncom.CoInitialize,pythoncom.CoUninitialize))
    def test_existing_final_is_never_overwritten_or_successful(self):
        p=self.root/'guard.json';p.write_bytes(b'EXISTING EVIDENCE')
        with self.assertRaises(D.OfficeCleanupError):
            with owned_office(p):pass
        self.assertEqual(p.read_bytes(),b'EXISTING EVIDENCE')
    def test_write_failure_preserves_primary_and_restores_guard(self):
        primary=ValueError('PRIMARY_OPEN_FAILURE')
        with patch('stage9_app.workspace.save',side_effect=OSError('FINAL_DISK_ERROR')):
            try:
                with owned_office(self.root/'guard.json'):raise primary
            except ValueError as actual:self.assertIs(actual,primary)
            else:self.fail('Primary exception was hidden')
        self.assertIn('FINAL_DISK_ERROR',' '.join(primary.__notes__))
        with owned_office(self.root/'next.json'):pass
    def test_event_write_failure_is_not_success(self):
        with patch('stage9_app.diagnostics.save',side_effect=PermissionError('EVENT_DISK_DENIED')):
            with self.assertRaises(D.OfficeCleanupError):
                with owned_office(self.root/'guard.json'):pass
        self.assertEqual(self.read(self.root/'guard.json')['status'],'FAIL')
        self.assertIn('EVENT_DISK_DENIED',self.read(self.root/'guard.json')['diagnostic_errors'][0]['error'])
    def test_concurrent_same_process_guard_rejected_without_patching(self):
        with owned_office(self.root/'first.json'):
            def other():
                with owned_office(self.root/'second.json'):pass
            with concurrent.futures.ThreadPoolExecutor(1) as pool:
                with self.assertRaisesRegex(RuntimeError,'OFFICE_GUARD_ALREADY_ACTIVE'):pool.submit(other).result()
        self.assertFalse((self.root/'second.json').exists())
    def test_parallel_diagnostic_allocations_and_events(self):
        def allocate(i):
            p=D.viewer_diagnostic(self.root);return p,D.event(p,'OPEN_BEGIN',{'attempt':i})
        with concurrent.futures.ThreadPoolExecutor(12) as pool:rows=list(pool.map(allocate,range(120)))
        self.assertEqual(len(set(p.parent for p,_ in rows)),120)
        self.assertEqual(sorted(self.read(e)['attempt'] for _,e in rows),list(range(120)))
    def test_forced_name_collision_allocates_fresh_identity_and_preserves_old(self):
        old=self.root/'viewer_A';old.mkdir();(old/'ownership.json').write_bytes(b'OLD')
        with patch.object(D,'uuid4',side_effect=[types.SimpleNamespace(hex='A'),types.SimpleNamespace(hex='B')]):
            p=D.viewer_diagnostic(self.root)
        self.assertEqual(p.parent.name,'viewer_B');self.assertEqual((old/'ownership.json').read_bytes(),b'OLD')
        event=p.with_name('ownership_CLOSE_A.json');event.write_bytes(b'OLD EVENT')
        with patch.object(D,'uuid4',side_effect=[types.SimpleNamespace(hex='A'),types.SimpleNamespace(hex='B')]):
            new=D.event(p,'CLOSE',{'status':'PASS'})
        self.assertNotEqual(str(event),new);self.assertEqual(event.read_bytes(),b'OLD EVENT')
    def viewer(self,failed=()):
        events=[]
        class Collection:
            def __init__(self,name):self.name=name
            def Close(self,*args):
                events.append(self.name+'.Close')
                if self.name+'.Close' in failed:
                    import pywintypes
                    raise pywintypes.com_error(-2147418111,'INJECTED_CLOSE_REJECTED',None,None)
        class Application:
            def __init__(self,name):
                self.name=name;setattr(self,'Documents' if name=='Word' else 'Workbooks',Collection(name))
            def Quit(self):
                events.append(self.name+'.Quit')
                if self.name+'.Quit' in failed:raise RuntimeError('INJECTED_QUIT_ERROR')
        with patch('stage9_app.gui.config_path',return_value=self.root/'settings.json'):v=Viewer()
        v.word=Application('Word');v.excel=Application('Excel');v.open_attempt_id='SOURCE_OPEN'
        return v,events
    def test_word_close_failure_persisted_and_excel_cleanup_independent(self):
        v,events=self.viewer(('Word.Close',));result=v.close()
        self.assertEqual(events,['Word.Close','Word.Quit','Excel.Close','Excel.Quit'])
        self.assertEqual(result['status'],'FAIL')
        errors=self.read(v.diagnostic)['cleanup_errors'];self.assertEqual(errors[0]['step'],'Word.Close')
        self.assertEqual(errors[0]['hresult'],-2147418111)
        self.assertIn('INJECTED_CLOSE_REJECTED',errors[0]['traceback'])
        self.assertIs(v.close(),result)
    def test_excel_close_and_quit_failures_both_recorded(self):
        v,events=self.viewer(('Excel.Close','Excel.Quit'));result=v.close()
        self.assertEqual(events,['Word.Close','Word.Quit','Excel.Close','Excel.Quit'])
        self.assertEqual(result['status'],'FAIL')
        self.assertEqual([r['step'] for r in self.read(v.diagnostic)['cleanup_errors']][:2],['Excel.Close','Excel.Quit'])
    def test_primary_open_error_not_replaced_by_cleanup(self):
        v,events=self.viewer(('Word.Close',));primary=ValueError('PRIMARY')
        result=v.close(primary=primary)
        self.assertEqual(result['primary_error']['error'],'PRIMARY')
        self.assertEqual(self.read(v.diagnostic)['primary_error']['error'],'PRIMARY')
        self.assertIn('Word.Close',' '.join(primary.__notes__))
        self.assertEqual(events[-2:],['Excel.Close','Excel.Quit'])
    def test_close_result_write_failure_returns_failure(self):
        v,_=self.viewer()
        real=D.event
        def write(path,kind,data):
            if kind=='CLOSE_RESULT':raise OSError('RESULT_DISK_FAILED')
            return real(path,kind,data)
        with patch('stage9_app.office.event',side_effect=write):result=v.close()
        self.assertEqual(result['status'],'FAIL');self.assertEqual(result['cleanup_errors'][-1]['step'],'Diagnostic.CloseResult')
    def test_gui_surfaces_cleanup_failure_and_does_not_report_success(self):
        v,_=self.viewer(('Word.Close',));messages=[];destroyed=[]
        app=types.SimpleNamespace(office=[v],busy=False,window=types.SimpleNamespace(destroy=lambda:destroyed.append(True)),
                                  status=types.SimpleNamespace(set=lambda s:messages.append(s)))
        app.close_viewers=lambda:App.close_viewers(app)
        with patch('stage9_app.gui.messagebox.showerror',side_effect=lambda *a,**k:messages.extend(a)):
            App.shutdown(app)
        self.assertFalse(destroyed);self.assertTrue(any('Word.Close' in m for m in messages))
        App.shutdown(app);self.assertEqual(destroyed,[True])
    def test_gui_normal_shutdown_first_attempt(self):
        v,_=self.viewer();destroyed=[]
        app=types.SimpleNamespace(office=[v],busy=False,window=types.SimpleNamespace(destroy=lambda:destroyed.append(True)))
        app.close_viewers=lambda:App.close_viewers(app)
        App.shutdown(app);self.assertEqual(destroyed,[True]);self.assertEqual(v.close_result['status'],'PASS')
    @contextmanager
    def simulated_child(self,*,living=False,owned=True,force_log_error=False,handle_error=False):
        # Real guard/dispatch/finalizer, fake Win32 boundaries. No real process
        # is created or addressed, so refusal branches are safe to exercise.
        import win32con
        events=[];alive=[living]
        class Handle:
            def __int__(self):return 77
            def Close(self):
                events.append('Handle.Close')
                if handle_error:raise OSError('HANDLE_CLOSE_INJECTED')
        handle=Handle()
        record={'pid':123,'creation_time':'BIRTH','image':r'C:\Office\EXCEL.EXE','hwnd':101,
                'progid':'Excel.Application','worker_pid':os.getpid(),'launcher_pid':os.getpid(),
                'actual_child':True,'owned':owned,'handle_source':'CREATE_PROCESS','retained_handle':77}
        book=types.SimpleNamespace(Count=0)
        app=types.SimpleNamespace(Workbooks=book,UserControl=False,Hwnd=101,Quit=lambda:events.append('Quit'))
        @contextmanager
        def native(*args,**kwargs):yield app
        def terminate(*args):events.append('Terminate');alive[0]=False
        tick=[0]
        def monotonic():tick[0]+=20;return tick[0]
        real=D.event
        def write(path,kind,data):
            if force_log_error and kind=='VIEWER_EMERGENCY_FORCE_BEGIN':raise OSError('FORCE_LOG_DENIED')
            return real(path,kind,data)
        child=types.SimpleNamespace(_handle=handle,returncode=None)
        def poll_child():
            child.returncode=None if alive[0] else 1 if living else 0
            events.append('Popen.poll')
            return child.returncode
        child.poll=poll_child
        with ExitStack() as stack:
            patches=[patch('stage9_app.ownership.launch_child',return_value=(child,record)),
                patch('stage9_app.ownership.native_application',side_effect=native),
                patch('stage9_app.ownership.request_empty_excel_close'),
                patch('stage9_app.ownership.time.monotonic',side_effect=monotonic),
                patch('win32event.WaitForSingleObject',side_effect=lambda *args:258 if alive[0] else win32con.WAIT_OBJECT_0),
                patch('win32process.GetExitCodeProcess',side_effect=lambda *args:259 if alive[0] else 1 if living else 0),
                patch('win32process.GetProcessId',return_value=123),
                patch('win32process.GetProcessTimes',return_value={'CreationTime':types.SimpleNamespace(isoformat=lambda:'BIRTH')}),
                patch('stage9_app.ownership.process_image',return_value=record['image']),
                patch('win32api.TerminateProcess',side_effect=terminate),
                patch('pythoncom._GetInterfaceCount',return_value=0),patch.object(D,'event',side_effect=write)]
            for p in patches:stack.enter_context(p)
            yield events,record
    def simulate_dispatch(self):
        import win32com.client
        app=win32com.client.DispatchEx('Excel.Application');app.Quit();app=None
    def test_handle_closed_even_when_final_save_fails(self):
        with self.simulated_child() as (events,_):
            with patch('stage9_app.workspace.save',side_effect=OSError('SAVE_DENIED')):
                # START write fails too; guard must still close its retained handle.
                with self.assertRaises(OSError):
                    with owned_office(self.root/'guard.json'):self.simulate_dispatch()
            self.assertIn('Handle.Close',events)
    def test_handle_close_failure_durably_reported(self):
        with self.simulated_child(handle_error=True) as (events,_):
            with self.assertRaises(D.OfficeCleanupError):
                with owned_office(self.root/'guard.json'):self.simulate_dispatch()
        self.assertEqual(self.read(self.root/'guard.json')['cleanup_errors'][0]['step'],'ProcessHandle.Close')
    def test_force_requires_exact_ownership_and_durable_begin(self):
        for name,kwargs in [('unowned',{'owned':False}),('log_denied',{'force_log_error':True})]:
            with self.simulated_child(living=True,**kwargs) as (events,_):
                with self.assertRaises(D.OfficeCleanupError):
                    with owned_office(self.root/(name+'.json')):self.simulate_dispatch()
                self.assertNotIn('Terminate',events,(name,self.read(self.root/(name+'.json'))))
                # The refused live child remains owned by Popen, whose handle
                # must remain valid for its eventual exit observation.
                self.assertIn('Popen.poll',events);self.assertNotIn('Handle.Close',events)
    def test_proven_emergency_force_stays_failure_and_records_final_outcome_once(self):
        with self.simulated_child(living=True) as (events,_):
            with self.assertRaises(D.OfficeCleanupError):
                with owned_office(self.root/'guard.json'):self.simulate_dispatch()
            self.assertIn('Terminate',events);self.assertIn('Handle.Close',events)
        result=self.read(self.root/'guard.json')
        self.assertEqual(result['status'],'FAIL');self.assertEqual(result['owned_orphans'],0)
        self.assertTrue(result['instances'][0]['forced_termination_attempted'])

if __name__=='__main__':unittest.main(verbosity=2)
