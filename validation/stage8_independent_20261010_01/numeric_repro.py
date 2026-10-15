"""Corroborate concurrency defect with real business data in valid native-saved XLSM."""
import independent as T
import shutil,subprocess,sys,json
from pathlib import Path
from stage8_closure import core,storage as S
from stage8_closure.office import excel_session
T.SEED=Path(S.read(T.RUN/'SEED.json')['sandbox'])
root=T.fixture();incoming=root/'transactions/VALID_CONCURRENT_VERSION.xlsm';shutil.copyfile(T.reg(root),incoming)
with excel_session() as excel:
    b=excel.Workbooks.Open(Filename=str(incoming),ReadOnly=False,UpdateLinks=0,CorruptLoad=0)
    try:b.Worksheets('УЗК').Cells(24,3).Value2=4;b.Save()
    finally:b.Close(SaveChanges=False)
native_before=T.native.inspect(incoming,{'УЗК!C24':4,'УЗК!C25':10});publisher=[]
def fault(point,target):
    if point!='before_replace':return
    code='import sys;from stage8_closure.storage import replace;replace(sys.argv[1],sys.argv[2]);print("VALID_NATIVE_XLSM_PUBLISHED")'
    r=subprocess.run([sys.executable,'-B','-c',code,str(incoming),str(T.reg(root))],env=T.ENV,capture_output=True,text=True,timeout=15)
    publisher.append({'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr,'balances_after_external_publish':T.balances(root),'sha256':S.sha(T.reg(root))})
p=T.plan(root);result=core.close(p,core.confirmation(p),_fault=fault)
detail={'sandbox':str(root),'preview_balance':12,'real_external_saved_balance':4,'external_version_native':native_before,
        'publisher':publisher,'product_returned':result['status'],'final_balances':T.balances(root),'journal':T.states(root),
        'expected':'STOP on competing replacement; never restore stale 12-minus-3=9 over balance 4',
        'actual':'Competing balance 4 was overwritten with stale balance 9; COMMITTED; external data lost'}
assert publisher[0]['exit_code']==0 and publisher[0]['balances_after_external_publish']['s8-a']==4
assert T.balances(root)['s8-a']==9 and result['status']=='CLOSED_TEST'
T.save(T.RUN/'NUMERIC_CONCURRENT_REPRO.json',detail)
T.save(T.RUN/'cases/concurrent_valid_xlsm_balance_resurrection.json',{'name':'concurrent_valid_xlsm_balance_resurrection','status':'FAIL','defect_confirmed':True,'detail':detail})
print(json.dumps({k:v for k,v in detail.items() if k!='external_version_native'},ensure_ascii=False))
