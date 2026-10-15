"""Headless source tests only. GUI callbacks use fake widgets/applications."""
from common import *
import win32api,win32process,pywintypes,types
from stage9_app.ownership import process_image
from unittest.mock import patch
from contextlib import contextmanager
baseline=[]
for pid in win32process.EnumProcesses():
    try:
        h=win32api.OpenProcess(0x100400,False,pid)
        try:
            path=process_image(h)
            if Path(path).name.upper() in ('WINWORD.EXE','EXCEL.EXE',EXE.name.upper()):
                baseline.append({'pid':pid,'image':path,'creation_time':win32process.GetProcessTimes(h)['CreationTime'].isoformat()})
        finally:h.Close()
    except pywintypes.error:pass
save(OUT/'BASELINE_PROCESSES.json',baseline)
parent=Path(SETUP['source_test_parent']);assert parent.resolve().is_relative_to(Path(SETUP['base']).resolve())
env=os.environ.copy();env['VZMK_TEST_PARENT']=str(parent)
log=OUT/'SOURCE_TARGETED_17.log'
with log.open('xb') as stream:
    proc=subprocess.run([sys.executable,'-B',str(ROOT/'tests/stage9_098/test_cleanup.py')],env=env,stdout=stream,stderr=subprocess.STDOUT,timeout=180)
assert proc.returncode==0,log
text=log.read_text(encoding='utf8');assert 'Ran 17 tests' in text and '\nOK' in text
save(OUT/'SOURCE_TARGETED_17.json',{'status':'PASS','source_only':True,'tests':17,'test_file':str(ROOT/'tests/stage9_098/test_cleanup.py'),'test_file_sha256':sha(ROOT/'tests/stage9_098/test_cleanup.py'),'compiled_binding':str(OUT/'COMPILED_SOURCE_BEFORE.json'),'log':str(log),'real_gui_created_or_controlled':False})
from stage9_app.office import Viewer
from stage9_app.workspace import save as diagnostic_save
folder=parent/'viewer_repeat';folder.mkdir();calls=[]
class Collection:
    def __init__(self,name):self.name=name
    def Open(self,**kw):assert kw['ReadOnly'] is True;calls.append((self.name+'.Open',kw['Filename']))
    def Close(self,*args):calls.append((self.name+'.Close',None))
class Application:
    def __init__(self,name):
        self.name=name;self.Options=types.SimpleNamespace()
        setattr(self,'Documents' if name=='Word' else 'Workbooks',Collection(name))
    def Quit(self):calls.append((self.name+'.Quit',None))
@contextmanager
def fake_process_guard(path,**kw):
    yield []
    diagnostic_save(path,{'status':'PASS','source_mock_guard':True,'instances':[]})
paths=[folder/'TEST.docx',folder/'TEST.xlsx',folder/'TEST.docx'];viewers=[]
with patch('stage9_app.ownership.owned_office',side_effect=fake_process_guard),patch('stage9_app.gui.config_path',return_value=folder/'settings.json'),patch('win32com.client.DispatchEx',side_effect=lambda kind:Application('Word' if kind=='Word.Application' else 'Excel')):
    def open_word(app,path,readonly):assert readonly is True;calls.append(('Word.Open',str(path)))
    with patch('passport_app.vik.open_word',side_effect=open_word):
        for i in range(3):
            v=Viewer();v.open(paths);result=v.close();assert result['status']=='PASS' and v.close() is result
            assert read(v.diagnostic)['source_mock_guard'];viewers.append({'viewer_id':v.viewer_id,'diagnostic':str(v.diagnostic),'close_result':result})
assert len({v['viewer_id'] for v in viewers})==3
assert [step for step,_ in calls].count('Word.Open')==3 and [step for step,_ in calls].count('Excel.Open')==3
assert all([step for step,_ in calls].count(step)==3 for step in ('Word.Close','Word.Quit','Excel.Close','Excel.Quit'))
save(OUT/'VIEWER_REPEAT_SOURCE.json',{'status':'PASS','source_only':True,'native_Office':False,'GUI':False,'cycles':3,'unique_diagnostics':3,'duplicate_input_deduplicated':True,'viewers':viewers,'calls':calls,'limitation':'Real frozen GUI viewer reopen/switch/shutdown is verified separately in this independent GUI run'})
print('Source 17/17 + three mock viewer reopen cycles PASS; no real GUI used')
