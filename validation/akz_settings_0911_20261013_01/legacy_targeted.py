from common import *
from stage9_app import akz_settings as K
import copy
w=bootstrap('LEGACY');task('OLD_LEGACY','akz-settings-init',workspace=w);path=w/'preferences/akz_auto/settings.json';cfg=read(path)
for key in ('air_c','surface_c','relative_humidity'):cfg['climate'][key]=[float(v) for v in cfg['climate'][key]]
for n,pairs in cfg['thickness'].items():cfg['thickness'][n]=[[float(v) for v in pair] for pair in pairs]
save(OUT/'LEGACY_FLOAT_SETTINGS.json',cfg);task('OLD_LEGACY','akz-settings-save',workspace=w,akz_settings=OUT/'LEGACY_FLOAT_SETTINGS.json')
check,_=task('OLD_LEGACY','check',workspace=w,vom=SETUP['vom'],variants='0,1,2,3')
generated,_=task('OLD_LEGACY','generate',workspace=w,preview=check['preview'],confirm_test=True);manifest=generated['manifest'];m=read(manifest);assert 'akz_settings_snapshot' not in m
before={manifest:sha(manifest),**{r['path']:sha(r['path']) for r in m['records']},**{str(p):sha(p) for p in Path(manifest).parent.rglob('PAIR.json')}}
days={str(p):sha(p) for p in path.parent.glob('days/*.json')};assert all('settings_digest_schema' not in read(p) for p in days)
task('EXE_LEGACY','akz-settings-save',workspace=w,akz_settings=OUT/'LEGACY_FLOAT_SETTINGS.json')
task('EXE_LEGACY_NUMERIC_REUSE','akz-auto-generate',workspace=w,preview=check['preview'],confirm_test=True);assert all(sha(p)==h for p,h in days.items())
other=copy.deepcopy(cfg);other['climate']['air_c']=[16,26];save(OUT/'LEGACY_FUTURE_SETTINGS.json',other)
task('EXE_LEGACY','akz-settings-save',workspace=w,akz_settings=OUT/'LEGACY_FUTURE_SETTINGS.json')
verified,p=task('EXE_LEGACY_HISTORY','verify',workspace=w,manifest=manifest);proof=verified['settings_validation'];assert proof['status']=='LEGACY_SETTINGS_RECOVERED_EXACT_SHA'
source=Path(proof['source']);hold=source.parent.parent/('evidence_hold_'+uuid4().hex+'.json');source.rename(hold)
try:
    denied,dp=task('EXE_LEGACY_UNPROVEN','verify','FAIL',workspace=w,manifest=manifest);assert 'LEGACY_AKZ_SETTINGS_UNPROVEN' in denied['error']
finally:hold.rename(source)
task('EXE_LEGACY_RESTART','verify',workspace=w,manifest=manifest)
assert all(sha(p)==h for p,h in before.items());assert all(sha(p)==h for p,h in days.items())
save(OUT/'LEGACY_FINAL.json',{'status':'PASS','legacy_created_with':'PRESERVED_FROZEN_0.9.10','verified_with':'FROZEN_0.9.11','manifest':manifest,'legacy_recovery':proof,'unproven_result':str(dp),'old_documents_manifests_receipts':before,'day_files_unchanged':days,'history_file_exactly_restored':str(source),'no_old_manifests_or_receipts_rewritten':True})
print('Legacy exact-SHA recovery, unknown-evidence STOP, float-day reuse PASS',flush=True)
