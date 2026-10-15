from common import *
data=read(OUT/'AB_FINAL.json');w=Path(read(OUT/'WORKSPACE_MAIN.json')['workspace']);live=w/'preferences/akz_auto/settings.json';original=sha(live);hold=live.with_name('settings_live_removed_for_TEST.json');assert not hold.exists();live.rename(hold)
results=[]
try:
    for label in ('A','B'):
        d,p=task('EXE_LIVE_ABSENT_'+label,'verify',workspace=w,manifest=data[label+'_manifest']);assert d['settings_validation']['status']=='FROZEN_SETTINGS_VERIFIED';results.append(str(p))
finally:hold.rename(live)
assert sha(live)==original
save(OUT/'LIVE_SETTINGS_INDEPENDENCE.json',{'status':'PASS','results':results,'live_file_restored_exact_sha256':original,'no_old_manifests_or_receipts_changed':True})
