"""Targeted source tests, explicitly separate from frozen EXE tests."""
import sys,importlib.util,types,json
from pathlib import Path
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage8_closure import storage as S
from stage9_app.workspace import save
OUT=ROOT/'validation/stage9_092_20261012_01';setup=S.read(OUT/'SETUP.json')
old=len(sys.argv)>1 and sys.argv[1]=='old';source=OUT/'source_before/stage9.py' if old else ROOT/'stage9.py'
spec=importlib.util.spec_from_file_location('entry_under_test',source);entry=importlib.util.module_from_spec(spec);spec.loader.exec_module(entry)
events=[]
class Document:
    @property
    def Content(self):raise RuntimeError('PRIMARY_CONTENT_FAILURE')
    def Close(self,*a,**kw):events.append('word.Close');raise RuntimeError('SECONDARY_CLOSE_FAILURE')
class Cell:Value2=None
class Sheet:
    def Cells(self,*a):return Cell()
class Book:
    def Worksheets(self,*a):return Sheet()
    def Close(self,*a,**kw):events.append('excel.Close')
class Word:
    Documents=types.SimpleNamespace(Add=lambda **kw:Document())
    def Quit(self,*a,**kw):events.append('word.Quit')
class Excel:
    Workbooks=types.SimpleNamespace(Add=lambda:Book())
    def Quit(self,*a,**kw):events.append('excel.Quit')
def dispatch(progid):return Excel() if progid=='Excel.Application' else Word()
target=OUT/('MINOR1_OLD.json' if old else 'MINOR1_FIXED_094.json')
with patch('win32com.client.DispatchEx',dispatch),patch.object(sys,'argv',['stage9.py','--task','office-probe','--workspace',setup['probe_workspace'],'--result',str(target)]):
    code=entry.main()
result=S.read(target);assert code==1
if old:
    assert result['error']=='SECONDARY_CLOSE_FAILURE' and events==['word.Close','word.Quit']
else:
    assert result['error']=='PRIMARY_CONTENT_FAILURE' and events==['word.Close','word.Quit','excel.Close','excel.Quit'],(result,events)
    assert result['cleanup_errors'][0]['step']=='Word.Close' and 'SECONDARY_CLOSE_FAILURE' in result['cleanup_errors'][0]['error']
save(OUT/('MINOR1_OLD_ASSERTIONS.json' if old else 'MINOR1_FIXED_094_ASSERTIONS.json'),{'status':'PASS','source_test_only':True,'source_sha256':S.sha(source),'events':events,'primary':result['error']})
print('minor1 reproduction' if old else 'minor1 regression','PASS',flush=True)
if old:
    old_e2e=S.read(ROOT/'validation/stage9_091_20261012_01/EXE_E2E_FINAL.json');m=S.read(old_e2e['manifests'][0])
    assert len(m['records'])==7 and len({str(Path(r['path']).resolve()).casefold() for r in m['records']})==6
    vik=[r for r in m['records'] if r['type']=='VIK'];assert len(vik)==2 and vik[0]['path']==vik[1]['path'] and vik[0]['number']!=vik[1]['number']
    save(OUT/'MINOR2_OLD.json',{'status':'REPRODUCED','manifest':old_e2e['manifests'][0],'document_records':7,'physical_files':6,'word_records':4,'word_files':3,'vik_records':vik,'prior_wrong_word_count':4})
