from common import *
import importlib.util,types
from unittest.mock import patch
spec=importlib.util.spec_from_file_location('review_entry',ROOT/'stage9.py');entry=importlib.util.module_from_spec(spec);spec.loader.exec_module(entry)
events=[]
class Doc:
    @property
    def Content(self):raise RuntimeError('INDEPENDENT_PRIMARY_CONTENT')
    def Close(self,*a):events.append('Word.Close');raise RuntimeError('INDEPENDENT_SECONDARY_CLOSE')
class Book:
    def Worksheets(self,*a):return types.SimpleNamespace(Cells=lambda *a:types.SimpleNamespace(Value2=None))
    def Close(self,*a):events.append('Excel.Close')
class Word:
    Documents=types.SimpleNamespace(Add=lambda **kw:Doc())
    def Quit(self):events.append('Word.Quit')
class Excel:
    Workbooks=types.SimpleNamespace(Add=lambda:Book())
    def Quit(self):events.append('Excel.Quit')
result=OUT/'MINOR1_SOURCE_RESULT.json'
with patch('win32com.client.DispatchEx',side_effect=lambda p:Word() if p=='Word.Application' else Excel()),patch.object(sys,'argv',['stage9.py','--task','office-probe','--workspace',read(OUT/'LIFECYCLE_FINAL.json')['workspace'],'--result',str(result)]):
    code=entry.main()
r=read(result)
assert code==1 and r['error']=='INDEPENDENT_PRIMARY_CONTENT' and events==['Word.Close','Word.Quit','Excel.Close','Excel.Quit']
assert r['cleanup_errors']==[{'step':'Word.Close','error':"RuntimeError('INDEPENDENT_SECONDARY_CLOSE')"}]
assert 'INDEPENDENT_PRIMARY_CONTENT' in r['traceback']
save(OUT/'MINOR1_SOURCE_ASSERTIONS.json',{'status':'PASS','source_only':True,'events':events,'primary':r['error'],'secondary':r['cleanup_errors'],'source_sha256':sha(ROOT/'stage9.py')})
print('Primary/secondary source regression PASS',flush=True)
