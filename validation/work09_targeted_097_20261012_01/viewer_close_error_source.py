"""Source-only: exercise the actual Viewer.close -> App.shutdown call chain."""
from common import *
from stage9_app.office import Viewer
from stage9_app.gui import App
import types
events=[]
class Collection:
    def __init__(self,name):self.name=name
    def Close(self,*a):
        events.append(self.name+'.Close')
        if self.name=='Word':raise RuntimeError('INDEPENDENT_WORD_COLLECTION_CLOSE_ERROR')
class Application:
    def __init__(self,name):
        self.name=name;setattr(self,'Documents' if name=='Word' else 'Workbooks',Collection(name))
    def Quit(self):events.append(self.name+'.Quit')
class Guard:
    def __exit__(self,*a):events.append('guard.__exit__');return None
def make_viewer():
    v=Viewer.__new__(Viewer);v.word=Application('Word');v.excel=Application('Excel');v.guard=Guard();v.closed=False;return v
v=make_viewer();returned=v.close();direct_events=list(events)
assert returned==['INDEPENDENT_WORD_COLLECTION_CLOSE_ERROR'] and direct_events==['Word.Close','Word.Quit','Excel.Close','Excel.Quit','guard.__exit__']
events.clear();v=make_viewer();w=types.SimpleNamespace(destroy=lambda:events.append('window.destroy'))
app=types.SimpleNamespace(busy=False,office=[v],window=w)
result=App.shutdown(app)
assert events==direct_events+['window.destroy'] and result is None and v.closed
save(OUT/'VIEWER_CLOSE_ERROR_SOURCE.json',{'status':'REPRODUCED','source_only':True,'frozen_fault_injection':False,'direct_close_returned_errors':returned,'actual_shutdown_events':events,'shutdown_return':result,'persisted_cleanup_error_attribute':getattr(v,'cleanup_errors',None),'source_sha256':{'office.py':sha(ROOT/'stage9_app/office.py'),'gui.py':sha(ROOT/'stage9_app/gui.py')},'finding':'Viewer.close returns Word.Close diagnostics; App.shutdown ignores them and destroys the window. No durable diagnostic records the swallowed error.'})
print('SOURCE ONLY: Viewer Word.Close error cleanup continues; diagnostic is discarded by actual shutdown call chain',flush=True)
