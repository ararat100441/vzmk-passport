"""Read only existing traces/configuration; never open the working registry."""
from common import *
from stage8_closure import storage as S
work=ROOT/'validation/work09_independent_20261012_01';manifest=read(work/'EVIDENCE_MANIFEST.json')
all_jobs={Path(p).parent for p in manifest['external_files'] if Path(p).name=='SUPERVISOR.json'}
audit=read(work/'FINAL_PROCESS_AND_SOURCE_AUDIT_COMPLETE.json')
jobs={Path(r['path']).parent for r in audit['reports']}
assert all(p in all_jobs for p in jobs)
observed=[]
for job in sorted(jobs):
    result=job/'worker_result.json';data=read(result) if result.exists() else {}
    observed.append({'job':str(job),'supervisor':str(job/'SUPERVISOR.json'),'worker_result_keys':list(data),'worker_error':data.get('error'),'workspace':data.get('workspace'),'manifest':data.get('manifest')})
cfg=read(ROOT/'settings.json')
from passport_app.config import Settings
settings=Settings.load(ROOT/'settings.json')
production=ROOT.parent.parent/'ВЗМК.xlsx'
assert settings.registry_path==production.resolve()
from stage9_app.core import settings as dev_settings
own=Path(SETUP['probe_workspace']);bound=dev_settings(own)
assert bound.registry_path==own/'DEV/ВЗМК_DEV.xlsx'
new_jobs=[]
for request in (Path(SETUP['base'])/'VZMKPassport/diagnostics').glob('job_*/JOB_REQUEST.json'):
    data=read(request);args=data['arguments'];assert not any(str(production).casefold() in v.casefold() for v in args)
    new_jobs.append({'request':str(request),'arguments':args})
result={'status':'INVESTIGATED_WRITER_UNKNOWN','working_registry':str(production),'historical_expected_sha256':'99195989df116c879ae6de0693fb81c44951f7bf565410590baeebb8d08e7cd2','work_observed_sha256':'462d0a600048e7424646ee2a6fdcf739ca64610b93a62406ad158c153f2c22d5','writer':'UNKNOWN','working_registry_opened_for_this_investigation':False,'working_registry_modified_or_restored':False,'production_sha_used_as_gate':False,
 'configured_source_registry':str(settings.registry_path),'bootstrap_registry_actual_expression':"source/'DEV/ВЗМК_DEV.xlsx'",'bootstrap_uses_configured_registry':False,'bound_registry':str(bound.registry_path),
 'write_path_code_evidence':{'stage9_app/workspace.py':'bootstrap copies under new UUID only; bind sets accepted storage roots to own DEV','stage9_app/core.py':'settings always own DEV registry; existing generation uses run/registry.xlsx; new Stage 1 uses own root/DEV','stage9.py':'Stage 8 sandbox operations pass accepted closure/recovery guards; tests use artificial 8001/2099'},
 'historical_work_jobs':observed,'all_handoff_and_work_job_count':len(all_jobs),'new_explicit_job_requests':new_jobs,'native_office_file_io_audit_available':False,'historical_product_write_proven':False,'external_writer_proven':False,'limitation':'Arguments and code path resolution do not identify the historical file writer; no file-level Office open/save trace or readable Windows file-audit record.'}
save(OUT/('REGISTRY_INVESTIGATION_'+(sys.argv[1] if len(sys.argv)>1 else 'REFINED')+'.json'),result)
print('Registry writer UNKNOWN; production file not opened; inspected',len(jobs),'historical Work jobs and',len(new_jobs),'new requests')
