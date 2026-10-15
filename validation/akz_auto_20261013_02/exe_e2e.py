from common import *
cfg=read(WORKSPACE/'preferences/akz_auto/settings.json')
cases=[]
def run(label,name,expected='PASS',error=None,**args):
    d,p,s=task('EXE_'+label,name,expected,**args)
    if error:assert error in d.get('error',''),d
    cases.append({'case':label,'task':name,'expected_status':expected,'actual_status':d['status'],'gate':'PASS','result':str(p),'error':d.get('error'),'office_instances':len(s['instances'])})
    save(OUT/('EXE_CHECKPOINT_'+label+'.json'),cases[-1])
    return d
valid=OUT/'SETTINGS_VALID.json';save(valid,cfg)
run('SETTINGS','akz-settings-save',akz_settings=valid)
for label,change,error in [
 ('INVERTED_RANGE',lambda c:c['climate'].update(air_c=[25,15]),'RANGE_LIMITS_INVALID'),
 ('OFFICIAL_CONFIG',lambda c:c.update(official_issue=True),'OFFICIAL_OR_PROJECT_STOP'),
 ('UNAPPROVED_FORMULA',lambda c:c['climate'].update(formula='OTHER'),'FORMULA_NOT_APPROVED'),
 ('LAYER_COUNT',lambda c:c['thickness'].update({'1':[[150,180],[150,180]]}),'LAYER_COUNT_STOP'),
 ('TEMPLATE_PINS',lambda c:c['template_bindings'].update({'1':{}}),'SETTINGS_TEMPLATE_SHA_STOP'),
 ('BELOW_NORM',lambda c:c['thickness'].update({'0':[[79,100]]}),'BELOW_TEMPLATE_NORM')]:
    import copy
    bad=copy.deepcopy(cfg);change(bad);incoming=OUT/('BAD_'+label+'.json');save(incoming,bad)
    before=sha(WORKSPACE/'preferences/akz_auto/settings.json')
    run(label,'akz-settings-save','FAIL',error=error,akz_settings=incoming)
    assert sha(WORKSPACE/'preferences/akz_auto/settings.json')==before
run('WRONG_PROJECT','akz-auto-check','FAIL',error='Для проекта',project='OTHER',vom=SETUP['vom'],variants='0,1,2,3')
checked=run('CHECK_ALL','akz-auto-check',vom=SETUP['vom'],variants='0,1,2,3')
preview=checked['preview']
run('NO_CONFIRMATION','akz-auto-generate','FAIL',error='Подтвердите TEST',preview=preview)
first=run('GENERATE_ALL','akz-auto-generate',preview=preview,confirm_test=True)
save(OUT/'EXE_PAIRS_FIRST.json',first['pairs'])
days=WORKSPACE/'preferences/akz_auto/days';day_hashes={str(p):sha(p) for p in days.glob('*.json')}
second=run('REPEAT_ALL','akz-auto-generate',preview=preview,confirm_test=True)
save(OUT/'EXE_PAIRS_REPEAT.json',second['pairs'])
assert {str(p):sha(p) for p in days.glob('*.json')}==day_hashes
assert first['pairs']['path']!=second['pairs']['path']
for day in day_hashes:
    records=[g['proof']['days'][Path(day).stem] for pairs in (first['pairs'],second['pairs']) for g in pairs['groups'].values() if Path(day).stem in g['proof']['days']]
    assert all(x==records[0] for x in records),day
save(OUT/'EXE_DAILY_CLIMATE_REUSE.json',{'status':'PASS','date_count':len(day_hashes),'files':day_hashes,'cycles':2,'same_date_equal_across_acts_layers_and_processes':True})
# Damage one artificial date used by variant 2, after variants 0 and 1 finish.
target=days/'2026-10-03.json';original=target.read_bytes();published={p.name for p in (WORKSPACE/'DEV/Stage9_AKZ_Auto').glob('s9_*')}
try:
    damaged=read(target);damaged['values']['surface_c']=-99;target.write_text(json.dumps(damaged),encoding='utf8')
    run('DAMAGED_DAY_ATOMIC_STOP','akz-auto-generate','FAIL',error='DAY_RECORD_DAMAGED',preview=preview,confirm_test=True)
    assert {p.name for p in (WORKSPACE/'DEV/Stage9_AKZ_Auto').glob('s9_*')}==published
finally:target.write_bytes(original)
assert {str(p):sha(p) for p in days.glob('*.json')}==day_hashes
changed=copy.deepcopy(cfg);changed['climate']['air_c']=[16,25];incoming=OUT/'SETTINGS_CHANGED.json';save(incoming,changed)
run('CHANGE_SETTINGS','akz-settings-save',akz_settings=incoming)
run('STALE_PREVIEW_STOP','akz-auto-generate','FAIL',error='SETTINGS_CHANGED_AFTER_PREVIEW',preview=preview,confirm_test=True)
new=run('CHANGED_SETTINGS_CHECK','akz-auto-check',vom=SETUP['vom'],variants='0,1,2,3')
run('DAY_RANGE_CONFLICT','akz-auto-generate','FAIL',error='DAY_SETTINGS_CONFLICT',preview=new['preview'],confirm_test=True)
run('RESTORE_SETTINGS','akz-settings-save',akz_settings=valid)
assert {str(p):sha(p) for p in days.glob('*.json')}==day_hashes
normal=run('OFFICE_NORMAL','office-probe')
run('OFFICE_ERROR','office-probe','FAIL',error='STAGE9_TEST_OFFICE_ERROR',probe_mode='error')
run('WORD_CLOSE_ERROR','office-probe','FAIL',error='STAGE9_TEST_OFFICE_ERROR',probe_mode='cleanup-error')
full=run('FULL_STAGE9_GENERATE','generate',preview=preview,confirm_test=True)
save(OUT/'EXE_FULL_GENERATE.json',full)
run('FULL_MANIFEST_VERIFY','verify',manifest=full['manifest'])
run('FULL_NATIVE_VERIFY','native-verify',manifest=full['manifest'])
save(OUT/'EXE_TESTS_FINAL.json',{'status':'PASS','runtime':'FROZEN_EXE','exe_sha256':sha(EXE),'cases':cases,'GUI':'NOT_RUN_NO_NEW_PERMISSION','python_in_PATH':False,'full_manifest':full['manifest']})
print('FROZEN EXE E2E FINAL PASS',len(cases),'cases',flush=True)
