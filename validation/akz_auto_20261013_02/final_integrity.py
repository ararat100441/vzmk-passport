from common import *
import difflib
accepted=read(OUT/'ACCEPTED_BEFORE.json');assert all(sha(p)==h for p,h in accepted.items())
seed=read(OUT/'SEED_BEFORE.json');assert all(sha(p)==h for p,h in seed['files'].items())
templates=read(ROOT/'validation/akz_auto_20261013_01/CURRENT_PROJECT_TEMPLATES.json')['files'];assert len(templates)==8 and all(sha(p['path'])==p['sha256'] for p in templates)
source=read(OUT/'SOURCE_AT_BUILD.json');assert all(sha(p)==h for p,h in source.items())
build=read(OUT/'BUILD_MANIFEST.json');current={str(p.relative_to(EXE.parent)):sha(p) for p in EXE.parent.rglob('*') if p.is_file()};assert current==build['files']
old=read(ROOT/'validation/work09_final_099_20261012_01/BUILD_SHA256_AFTER.json');oldexe=Path(old['exe']);oldcurrent={str(p.relative_to(oldexe.parent)):sha(p) for p in oldexe.parent.rglob('*') if p.is_file()};assert oldcurrent==old['files']
meta=read(WORKSPACE/'STAGE9_WORKSPACE.json');fixture='DEV/Stage6_Fixtures/assignments.xlsx'
immutable={str(WORKSPACE/p):r['sha256'] for p,r in meta['inputs'].items() if p!=fixture};assert all(sha(p)==h for p,h in immutable.items())
fixture1=read(OUT/'TEST_ASSIGNMENT_FIXTURE.json');fixture2=read(OUT/'TEST_ASSIGNMENT_FIXTURE_R2.json')
assert fixture1['before_sha256']==meta['inputs'][fixture]['sha256'] and fixture1['after_sha256']==fixture2['before_sha256'] and sha(WORKSPACE/fixture)==fixture2['after_sha256']
save(OUT/'SOURCE_INTEGRITY_FINAL.json',{'status':'PASS','accepted_stage1_8_files':accepted,'accepted_file_count':len(accepted),'immutable_seed_files':len(seed['files']),'immutable_workspace_inputs':immutable,'authorized_artificial_assignment_fixture':{'path':str(WORKSPACE/fixture),'original_sha256':meta['inputs'][fixture]['sha256'],'current_sha256':sha(WORKSPACE/fixture),'receipt_chain':[str(OUT/'TEST_ASSIGNMENT_FIXTURE.json'),str(OUT/'TEST_ASSIGNMENT_FIXTURE_R2.json')],'production_catalog_changed':False},'real_first_passport_template_pairs':templates,'build_source_still_equal':True,'new_onedir_files':len(current),'new_onedir_unchanged':True,'preserved_099_exe_sha256':sha(oldexe),'preserved_099_onedir_files':len(oldcurrent),'preserved_099_onedir_unchanged':True,'real_registries_or_VOM_written':False,'scope_note':'Real template and sealed DEV seed hashes are verified. No claim of OS-wide write tracing.'})
changed=[];patch=[]
paths=[ROOT/'stage9.py',*sorted((ROOT/'stage9_app').glob('*.py')),ROOT/'stage9_0910.spec',ROOT/'tests/stage9_akz_auto/test_auto.py']
for p in paths:
    rel=p.relative_to(ROOT);before=oldexe.parent/'_internal'/rel
    previous=before.read_text(encoding='utf8') if before.is_file() else ''
    actual=p.read_text(encoding='utf8')
    if previous==actual:continue
    dest=OUT/'source_before'/rel;dest.parent.mkdir(parents=True,exist_ok=True)
    if before.is_file():dest.write_bytes(before.read_bytes())
    changed.append({'path':str(p),'relative':str(rel),'before_sha256':sha(before) if before.is_file() else None,'after_sha256':sha(p),'change':'MODIFIED' if before.is_file() else 'ADDED'})
    patch+=difflib.unified_diff(previous.splitlines(True),actual.splitlines(True),fromfile='0.9.9/'+rel.as_posix(),tofile='0.9.10/'+rel.as_posix())
save(OUT/'CHANGED_FILES.json',changed);(OUT/'EXACT_CODE_DIFF.patch').write_text(''.join(patch),encoding='utf8')
final=read(OUT/'OFFICE_BASELINE_FINAL.json');assert not final['foreign_stop'] and all(p['alive'] for p in final['processes'])
cmd=['powershell.exe','-NoProfile','-NonInteractive','-Command',"Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='EXCEL.EXE'\" | Select-Object Name,ProcessId,ParentProcessId,CreationDate,ExecutablePath | ConvertTo-Json -Depth 3"]
p=subprocess.run(cmd,capture_output=True,encoding='utf8',timeout=30,creationflags=subprocess.CREATE_NO_WINDOW);assert p.returncode==0
snapshot=json.loads(p.stdout) if p.stdout.strip() else [];snapshot=[snapshot] if isinstance(snapshot,dict) else snapshot
save(OUT/'OFFICE_PROCESS_SNAPSHOT_FINAL.json',{'processes':snapshot,'baseline_original_handles':final,'previous_Word_PID_3132_cause':'UNKNOWN_HISTORICAL_NOT_ATTRIBUTED_TO_THIS_RUN'})
jobs=[];seen=set()
for path in sorted(OUT.glob('*.json')):
    d=read(path)
    if not isinstance(d,dict) or 'office_supervisor' not in d:continue
    if d['office_supervisor'] in seen:continue
    seen.add(d['office_supervisor'])
    supervisor=read(d['office_supervisor']);assert supervisor['owned_orphans']==0 and not supervisor['ownership_errors']
    instances=supervisor['instances'];frozen=path.name.startswith('EXE_')
    if frozen:assert all(not i.get('forced_termination_attempted') and i['exit_code']==0 for i in instances)
    jobs.append({'result':str(path),'runtime':'FROZEN_EXE' if frozen else 'SOURCE_OR_BOOTSTRAP','status':d['status'],'supervisor':d['office_supervisor'],'office_instances':len(instances),'owned_orphans_final':0,'ownership_errors':[],'forced_instances':[i for i in instances if i.get('forced_termination_attempted')]})
save(OUT/'OFFICE_VALIDATION_FINAL.json',{'status':'PASS_FROZEN_EXE_SCOPE','jobs':jobs,'frozen_task_launches':sum(j['runtime']=='FROZEN_EXE' for j in jobs),'frozen_office_instances':sum(j['office_instances'] for j in jobs if j['runtime']=='FROZEN_EXE'),'frozen_forced_terminations':0,'all_final_owned_orphans':0,'source_initial_forced_terminations':sum(len(j['forced_instances']) for j in jobs if j['runtime']!='FROZEN_EXE'),'foreign_baseline_preserved':True,'historical_source_failure':'Initial RPC_E_CALL_REJECTED while reading DEV VOM; owned Excel cleanup required supervisor fallback. Exact root cause UNKNOWN; not reproduced in subsequent source/frozen runs.'})
print('Integrity PASS:',len(accepted),'accepted files;',len(seed['files']),'seed files;',len(changed),'changed/added files;',len(current),'new onedir;',len(oldcurrent),'preserved 099 onedir',flush=True)
