"""Reuse bounded independent Work08 scenarios on fresh data and current product."""
import gc,json,runpy,shutil,sys,time,traceback
from pathlib import Path
from stage8_closure import core,storage as S,office
from support import RUN,save

OUT=RUN/sys.argv[1];OUT.mkdir(exist_ok=False);(OUT/'cases').mkdir()
source=S.ROOT/'validation/stage8_independent_targeted_20261011_01/review_tests.py'
copy=RUN/'PRIOR_REVIEW_HARNESS.py';shutil.copyfile(source,copy)
saved_argv=sys.argv;sys.argv=[str(copy)]
module=runpy.run_path(str(copy),run_name='work08_recovery_regression');sys.argv=saved_argv
g=module['race'].__globals__
g.update(SEED=Path(S.read(RUN/'SEED.json')['sandbox']),EXTERNAL=RUN/'EXTERNAL_VALID_4.xlsm')
shutil.copyfile(S.ROOT/'validation/stage8_independent_targeted_20261011_01/EXTERNAL_VALID_4.xlsm',g['EXTERNAL'])
product={str(p):S.sha(p) for p in [*sorted((S.ROOT/'stage8_closure').glob('*.*')),S.ROOT/'stage8.py',S.ROOT/'STAGE8_DEV.md',S.ROOT/'Запустить Этап 8 DEV.cmd'] if p.is_file()}
S.write(OUT/'PRODUCT_BEFORE.json',product)
S.write(OUT/'REUSED_HARNESS_BINDING.json',{'source':str(source),'source_sha256':S.sha(source),'test_copy':str(copy),'copy_sha256':S.sha(copy),'selected_only':True,'old_outputs_not_written':True})
cases=[('previous_external_rename_replace',lambda:g['race']('rename','before_replace')),
       ('previous_external_native_replace',lambda:g['race']('native','before_replace')),
       ('previous_valid_external_4_stale_9',lambda:g['race']('native','during_publish')),
       ('previous_direct_write_critical',lambda:g['race']('write','during_publish')),
       ('previous_parent_rename',lambda:g['race']('parent','after_replace')),
       ('previous_external_4_before_acquisition',g['prior_publish']),
       ('previous_two_simultaneous_closes',g['parallel']),
       ('previous_existing_writer',g['preexisting_writer']),
       ('previous_writable_mapping',lambda:g['preexisting_writer'](True)),
       ('previous_cleanup_first_error_Quit',lambda:g['cleanup']('CutCopyMode')),
       ('previous_cleanup_Close_Quit',lambda:g['cleanup']('Workbooks.Close')),
       ('previous_cleanup_Quit_error',lambda:g['cleanup']('Quit')),
       ('previous_cleanup_primary_exception',lambda:g['cleanup']('Workbooks.Close',True)),
       ('previous_Workbook_Close_primary_hresult',g['book_close_primary']),
       ('previous_unrelated_Excel',g['unrelated_excel']),
       ('previous_RPC_E_DISCONNECTED_exact',g['disconnect']),
       ('previous_real_Excel_disconnect',lambda:g['disconnect'](False)),
       ('DEV_9_2026_still_BLOCKED',g['dev_forbidden'])]
results=[]
with office.diagnostic_scope(RUN/'office'):
    for name,fn in cases:
        start=time.monotonic()
        try:
            detail=fn();gc.collect()
            pid=detail.get('pid',detail.get('sentinel_pid')) if isinstance(detail,dict) else None
            if pid:assert g['wait_exit'](pid),'OWNED_EXCEL_NOT_EXITED'
            item={'case':name,'status':'PASS','detail':detail}
        except Exception as e:item={'case':name,'status':'FAIL','error':str(e),'traceback':traceback.format_exc()}
        item['seconds']=round(time.monotonic()-start,3)
        assert all(S.sha(p)==h for p,h in product.items()),'PRODUCT_CHANGED'
        S.write(OUT/'cases'/(name+'.json'),item);results.append(item);save(OUT/'RESULTS.json',results)
        print(item['status'],name,item.get('error',''),flush=True)
S.write(OUT/'PRODUCT_AFTER.json',{p:S.sha(p) for p in product})
print('DONE',sum(r['status']=='PASS' for r in results),sum(r['status']=='FAIL' for r in results),flush=True)
