"""Capture specification failures; these are not expected-negative PASS cases."""
from common import *
work=Path(read(OUT/'WORKSPACE.json')['workspace']);cfg_before=read(OUT/'DEFAULT_SETTINGS.json');cfg_after=read(work/'preferences/akz_auto/settings.json')
assert cfg_before==cfg_after
if not (OUT/'GUI_SAVED_SETTINGS.json').exists():save(OUT/'GUI_SAVED_SETTINGS.json',cfg_after)
else:assert read(OUT/'GUI_SAVED_SETTINGS.json')==cfg_after
manifest=Path(read(OUT/'FROZEN_AUTO_FINAL.json')['full_manifest']);m=read(manifest)
settings=work/'preferences/akz_auto/settings.json';before=m['inputs'][str(settings)];after=sha(settings)
assert before!=after and digest(cfg_before['climate'])!=digest(cfg_after['climate'])
assert all(sha(p)==h for p,h in read(OUT/'REPEAT_AND_DAILY_CACHE.json')['day_files_unchanged'].items())
p=OUT/'OLD_MANIFEST_AFTER_GUI_SAVE.json';cmd=[str(EXE),'--task','verify','--workspace',str(work),'--manifest',str(manifest),'--result',str(p)]
assert not (OUT/'FOREIGN_OFFICE_STOP.json').exists()
if not p.exists():
 save(p.with_suffix('.command.json'),{'cmd':cmd,'spec_expected':'PASS','exe_sha256':sha(EXE),'PATH':environment()['PATH'],'started':time.time()})
 proc=subprocess.run(cmd,env=environment(),capture_output=True,timeout=120)
 save(p.with_suffix('.process.json'),{'returncode':proc.returncode,'stdout':proc.stdout.decode('utf8','replace'),'stderr':proc.stderr.decode('utf8','replace'),'ended':time.time()})
actual=read(p);sup=read(actual['office_supervisor']);assert not sup['instances'] and sup['owned_orphans']==0
assert actual['status']=='FAIL' and 'settings.json' in actual['error'] and 'Изменён источник' in actual['error']
save(OUT/'SETTINGS_REPRESENTATION_AND_MANIFEST_DEFECT.json',{'status':'REPRODUCED_PRODUCT_FAILURE','spec_expected':'PASS','actual':'FAIL','numeric_settings_equal':True,'before_settings_sha256':before,'after_settings_sha256':after,'before_climate_digest':digest(cfg_before['climate']),'after_climate_digest':digest(cfg_after['climate']),'stored_day_climate_digests':{p.name:read(p)['settings_digest'] for p in (work/'preferences/akz_auto/days').glob('*.json')},'old_manifest':str(manifest),'manifest_seal_valid':m['seal']==digest({k:v for k,v in m.items() if k!='seal'}),'all_published_document_SHA_unchanged':all(sha(r['path'])==r['sha256'] for r in m['records']),'result':str(p),'error':actual['error'],'source_of_change':'actual frozen GUI Save unchanged values; GUI_02/GUI_04 screenshots','GUI_generation':'PENDING_CAPTURE','no_product_fixes':True})
print('Confirmed: numeric-equal GUI save changes climate digest and invalidates previously completed manifest',flush=True)
