from common import *
base=Path(read(ROOT/'validation/stage9_092_20261012_01/SETUP.json')['base'])/'VZMKPassport/diagnostics'
cases=[]
for job in ('job_4a7655cd24734d8f8bd8faa51b51956d','job_dc4c6a0fd1714a50961d06090eb589a6','job_c7d78a9e8f08474a82ea7af0b2075d6b'):
    folder=base/job;sup=read(folder/'SUPERVISOR.json');assert sup['worker_returncode']==0 and sup['owned_orphans']==0 and not sup['ownership_errors']
    records=[r for r in sup['instances'] if r.get('normal_window_close_requested')];assert len(records)==1
    r=records[0];assert r['quit_returned'] and r['workbooks_before_quit']==0 and r['exit_code']==0 and not r.get('forced_termination_attempted')
    event=read(next(folder.glob('office_EMPTY_EXCEL_WM_CLOSE_BEGIN_*.json')));d=event['data']
    assert d['handle_validated'] and d['com_interfaces']==0 and d['message']=='WM_CLOSE' and d['pid']==r['pid'] and d['creation_time']==r['creation_time'] and d['handle_source']=='CREATE_PROCESS' and d['actual_child'] and d['owned']
    created=read(folder/('created_'+str(r['pid'])+'.json'));assert all(created[k]==r[k] for k in ('pid','image','creation_time','launcher_pid','retained_handle'))
    cases.append({'job':str(folder),'supervisor_sha256':sha(folder/'SUPERVISOR.json'),'identity':r,'WM_CLOSE_proof':event})
save(OUT/'HISTORICAL_WM_CLOSE_READBACK.json',{'status':'PASS','evidence_readback_only':True,'new_independent_runtime':False,'cases':cases,'all_references_previously_SHA_verified':True})
print('Three frozen 0.9.7 WM_CLOSE cases readback PASS; not independent runtime',flush=True)
