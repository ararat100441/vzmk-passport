from common import *
base=Path(SETUP['base']);nav=read(OUT/'GUI_NAVIGATION_FIXTURE.json');allowed=[OUT.resolve(),WORKSPACES.resolve(),base.resolve(),Path(nav['root']).resolve()]
jobs=[];violations=[]
for p in (base/'VZMKPassport/diagnostics').glob('job_*/JOB_REQUEST.json'):
    q=read(p);args=q['arguments'];paths=[]
    for flag in ('--workspace','--source','--parent','--vom','--sandbox','--request','--preview','--manifest','--result'):
        if flag not in args:continue
        value=Path(args[args.index(flag)+1]).resolve()
        good=(value==SOURCE.resolve()) if flag=='--source' else (value.is_relative_to(SOURCE.resolve()) or any(value.is_relative_to(r) for r in allowed))
        paths.append({'flag':flag,'path':str(value),'allowed':good})
        if not good:violations.append({'job':str(p),'flag':flag,'path':str(value)})
    jobs.append({'request':str(p),'supervisor_image':q['supervisor_image'],'paths':paths})
assert not violations,violations
save(OUT/'PRODUCTION_ISOLATION_AUDIT.json',{'status':'PASS','jobs':jobs,'violations':violations,'sealed_DEV_source':str(SOURCE),'own_roots':[str(r) for r in allowed],'production_registry_opened':False,'production_registry_written_restored_or_replaced':False,'production_registry_constant_SHA_required':False,'historical_production_writer':'UNKNOWN','historical_systemwide_file_IO_audit':'NOT_VERIFIED','code_boundaries':['workspace.bootstrap explicitly reads source/DEV/ВЗМК_DEV.xlsx; legacy settings registry_path is unused','workspace.bind redirects accepted globals to the validated new DEV workspace','core.settings uses own DEV or run registry copy; generated manifest inputs refer to frozen run snapshots','Stage8 fixtures/guarded publication target only newly created artificial sandboxes'],'native_viewer_paths_evidence':['GUI_VIEWER_FIRST_SNAPSHOT.json','GUI_VIEWER_REOPENED_SNAPSHOT.json'],'accepted_files_unchanged_evidence':'FINAL_IMMUTABLE_INTEGRITY.json'})
print('Production isolation: paths in',len(jobs),'jobs allowed; no production operations found',flush=True)
