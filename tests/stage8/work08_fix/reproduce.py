"""Before-fix reproductions on fresh fixtures; never edits prior Work08 evidence."""
import json,os,shutil,subprocess,sys,time
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import patch
import win32file,win32con,win32process,win32api,pythoncom
from stage8_closure import core,storage as S,native
from stage8_closure.fixtures import create
from stage8_closure.office import excel_session

ROOT=S.ROOT;RUN=ROOT/'validation/stage8_fix_20261010_01';ENV=os.environ|{'PYTHONPATH':str(ROOT),'PYTHONUTF8':'1'}
OUT=RUN/'before_reproduction';OUT.mkdir(exist_ok=False)
PUBLISHER=ROOT/'tests/stage8/work08_fix/publisher.py'
def save(p,x):S.write(p,x)
def balances(reg):return {k:v['before'] for k,v in __import__('stage8_closure.model',fromlist=['registry_rows']).registry_rows(reg,'STAGE8 ИСКУССТВЕННЫЙ ПРОЕКТ',{'s8-a','s8-b'}).items()}
def publish(mode,src,dst):
    p=subprocess.run([sys.executable,'-B',str(PUBLISHER),mode,str(src),str(dst)],env=ENV,capture_output=True,text=True,encoding='utf8',timeout=20)
    return {'exit':p.returncode,'stdout':p.stdout,'stderr':p.stderr}

for mode in ('rename','native','numeric'):
    f=create(RUN/'sandboxes');root=Path(f['sandbox']);reg=root/'registry/УЗК_TEST.xlsm'
    incoming=root/'transactions/EXTERNAL_4.xlsm';shutil.copyfile(reg,incoming)
    with excel_session() as excel:
        b=excel.Workbooks.Open(Filename=str(incoming),UpdateLinks=0,ReadOnly=False,AddToMru=False,CorruptLoad=0)
        try:b.Worksheets('УЗК').Cells(24,3).Value2=4;b.Save()
        finally:b.Close(SaveChanges=False)
    validated=native.inspect(incoming,{'УЗК!C24':4});attempt=[];p=core.preview(root,f['request'])
    def fault(point,target):
        if point=='before_replace':
            r=publish('rename' if mode=='rename' else 'native',incoming,reg)
            r['observed']=balances(reg);attempt.append(r)
    result=core.close(p,core.confirmation(p),_fault=fault)
    assert attempt[0]['exit']==0 and attempt[0]['observed']['s8-a']==4
    assert result['status']=='CLOSED_TEST' and balances(reg)['s8-a']==9
    save(OUT/(mode+'.json'),{'defect_confirmed':True,'product':core.code_bindings(),'sandbox':str(root),
        'input_sha':p['payload']['before_sha256'],'external_native':validated,'external_process':attempt,
        'final_sha':S.sha(reg),'final_balances':balances(reg),'states':[e['state'] for e in S.history(root)['events']]})
    print('BEFORE',mode,'LOST_4_TO_9_COMMITTED_CONFIRMED',flush=True)

owned={};calls=[]
class Proxy:
    def __init__(self,real):object.__setattr__(self,'real',real)
    def __getattr__(self,k):
        if k=='Quit':
            def quit():calls.append('Quit');return self.real.Quit()
            return quit
        return getattr(self.real,k)
    def __setattr__(self,k,v):
        if k=='CutCopyMode':calls.append('CutCopyMode');raise pythoncom.com_error(-2147418111,'INJECTED_CLEANUP',None,None)
        setattr(self.real,k,v)
import win32com.client
dispatch=win32com.client.DispatchEx
def capture(*a,**k):
    e=dispatch(*a,**k);owned['real']=e;owned['pid']=win32process.GetWindowThreadProcessId(e.Hwnd)[1];return Proxy(e)
try:
    with patch('win32com.client.DispatchEx',side_effect=capture):
        with excel_session():pass
    h=win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION,False,owned['pid'])
    try:alive=win32process.GetExitCodeProcess(h)==259
    finally:h.Close()
    assert alive and calls==['CutCopyMode']
    save(OUT/'cleanup.json',{'defect_confirmed':True,'calls':calls,'owned_pid':owned['pid'],'alive':alive,'error_swallowed':True})
    print('BEFORE cleanup QUIT_SKIPPED_CONFIRMED',flush=True)
finally:owned['real'].Quit()

# OS capability probe, independent of the application's current publication code.
folder=OUT/'win32_probe';folder.mkdir();results=[]
for mode in ('rename','native','replace','write'):
    dst=folder/(mode+'_target');src=folder/(mode+'_incoming');dst.write_bytes(b'12');src.write_bytes(b'4')
    h=win32file.CreateFile(str(dst),win32con.GENERIC_READ|win32con.GENERIC_WRITE,
                         win32con.FILE_SHARE_READ,None,win32con.OPEN_EXISTING,0,None)
    try:
        r=publish(mode,src,dst);assert r['exit']!=0,(mode,r)
        assert dst.read_bytes()==b'12';results.append({'mode':mode,**r})
    finally:h.Close()
save(OUT/'NO_DELETE_HANDLE_PROBE.json',{'status':'PASS','results':results})
print('NO_DELETE Windows/NTFS probe PASS',flush=True)
