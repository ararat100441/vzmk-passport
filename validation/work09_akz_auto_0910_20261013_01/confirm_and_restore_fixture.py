from common import *
import copy
work=Path(read(OUT/'WORKSPACE.json')['workspace']);manifest=read(OUT/'FROZEN_AUTO_FINAL.json')['full_manifest']
cfg=read(OUT/'GUI_SAVED_SETTINGS.json');changed=copy.deepcopy(cfg);changed['climate']['air_c'][0]=16.0
incoming=OUT/'LEGITIMATE_FUTURE_DATE_SETTINGS.json';save(incoming,changed)
d,p=task('LEGITIMATE_RANGE_CHANGE','akz-settings-save',workspace=work,akz_settings=incoming)
result_path=OUT/'OLD_MANIFEST_AFTER_LEGITIMATE_CHANGE.json';cmd=[str(EXE),'--task','verify','--workspace',str(work),'--manifest',str(manifest),'--result',str(result_path)]
save(result_path.with_suffix('.command.json'),{'cmd':cmd,'spec_expected':'PASS','exe_sha256':EXPECTED,'started':time.time()})
proc=subprocess.run(cmd,env=environment(),capture_output=True,timeout=120)
save(result_path.with_suffix('.process.json'),{'returncode':proc.returncode,'stderr':proc.stderr.decode('utf8','replace'),'stdout':proc.stdout.decode('utf8','replace'),'ended':time.time()})
r=read(result_path);assert r['status']=='FAIL' and 'settings.json' in r['error']
restore,rp=task('RESTORE_OWN_TEST_SETTINGS','akz-settings-save',workspace=work,akz_settings=OUT/'DEFAULT_SETTINGS.json')
verified,vp=task('OLD_MANIFEST_AFTER_RESTORE','verify',workspace=work,manifest=manifest)
assert all(sha(p)==h for p,h in read(OUT/'REPEAT_AND_DAILY_CACHE.json')['day_files_unchanged'].items())
assert all(sha(p)==h for p,h in read(OUT/'GUI_BEFORE_OUTPUT_HASHES.json').items())
save(OUT/'MANIFEST_MUTABLE_SETTINGS_CONFIRMED.json',{'status':'REPRODUCED_PRODUCT_FAILURE','spec_expected':'PASS','actual_after_legitimate_change':'FAIL','legitimate_change':'air lower15->16 for future dates; historical documents must remain verifiable','result':str(result_path),'error':r['error'],'restore_scope':'OWN_TEST_PREFERENCES_ONLY, not production/product','restore_result':str(rp),'old_manifest_after_exact_settings_restore':'PASS','verify_after_restore':str(vp),'all_preexisting_AKZ_XLSX_SHA_unchanged':True,'all_day_record_SHA_unchanged':True})
print('Independent manifest defect confirmed after legitimate change; own TEST preferences restored, old verify PASS',flush=True)
