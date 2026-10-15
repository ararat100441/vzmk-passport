from common import *
import copy,traceback
work=Path(read(OUT/'WORKSPACE.json')['workspace']);s=config();cfg=read(OUT/'DEFAULT_SETTINGS.json');cases=[]
def run(label,name,expected='PASS',error=None,**kw):
 d,p=task(label,name,expected,workspace=work,**kw)
 if error:assert error in d.get('error',''),d
 cases.append({'label':label,'task':name,'expected':expected,'actual':d['status'],'result':str(p),'sha256':sha(p)})
 save(OUT/('CHECKPOINT_'+label+'.json'),cases[-1]);return d
try:
 # Test incoming configuration failures without changing persisted settings.
 badcases=[('INVERTED',lambda c:c['climate'].update(air_c=[25,15]),'RANGE_LIMITS_INVALID'),('NONFINITE',lambda c:c['climate'].update(air_c=[float('nan'),25]),'NUMBER_INVALID'),('PRECISION',lambda c:c['climate'].update(air_c=[15.01,25]),'PRECISION_INVALID'),('OFFICIAL',lambda c:c.update(official_issue=True),'OFFICIAL_OR_PROJECT_STOP'),('FORMULA',lambda c:c['climate'].update(formula='OTHER'),'FORMULA_NOT_APPROVED'),('LAYERS',lambda c:c['thickness'].update({'2':[[100,120]]}),'LAYER_COUNT_STOP'),('PINS',lambda c:c['template_bindings'].update({'3':{}}),'TEMPLATE_SHA_STOP'),('BELOW_NORM',lambda c:c['thickness'].update({'0':[[79,100]]}),'BELOW_TEMPLATE_NORM'),('SUM_BELOW180',lambda c:c['thickness'].update({'2':[[100,120],[70,100]]}),'CUMULATIVE_180_STOP')]
 for label,change,error in badcases:
  bad=copy.deepcopy(cfg);change(bad);p=OUT/('BAD_'+label+'.json');save(p,bad);before=sha(work/'preferences/akz_auto/settings.json')
  run(label,'akz-settings-save','FAIL',error=error,akz_settings=p);assert sha(work/'preferences/akz_auto/settings.json')==before
 run('WRONG_PROJECT','akz-auto-check','FAIL',error='Для проекта',project='OTHER',vom=s['vom'],variants='0,1,2,3')
 checked=run('CHECK_ALL','akz-auto-check',vom=s['vom'],variants='0,1,2,3');preview=checked['preview']
 assert checked['data']['akz_auto'] is True
 run('NO_CONFIRM','akz-auto-generate','FAIL',error='Подтвердите TEST',preview=preview)
 first=run('FIRST','akz-auto-generate',preview=preview,confirm_test=True);save(OUT/'PAIRS_FIRST.json',first['pairs'])
 first_files={str(p):sha(p) for p in Path(first['pairs']['path']).rglob('*') if p.is_file()}
 days=work/'preferences/akz_auto/days';dayhash={str(p):sha(p) for p in days.glob('*.json')}
 second=run('REPEAT','akz-auto-generate',preview=preview,confirm_test=True);save(OUT/'PAIRS_REPEAT.json',second['pairs'])
 assert first['pairs']['path']!=second['pairs']['path'] and all(sha(p)==h for p,h in first_files.items())
 assert dayhash=={str(p):sha(p) for p in days.glob('*.json')}
 for p in dayhash:
  records=[g['proof']['days'][Path(p).stem] for pairs in (first['pairs'],second['pairs']) for g in pairs['groups'].values() if Path(p).stem in g['proof']['days']]
  assert all(r==records[0] for r in records)
 save(OUT/'REPEAT_AND_DAILY_CACHE.json',{'status':'PASS','first_files_unchanged':first_files,'day_files_unchanged':dayhash,'separate_processes':True,'distinct_output_roots':True,'climate_equal_across_all_matching_dates':True})
 target=days/'2026-10-03.json';original=target.read_bytes();published={p.name for p in (work/'DEV/Stage9_AKZ_Auto').glob('s9_*')}
 try:
  bad=read(target);bad['values']['surface_c']=-99;target.write_text(json.dumps(bad),encoding='utf8')
  run('DAY_TAMPER','akz-auto-generate','FAIL',error='DAY_RECORD_DAMAGED',preview=preview,confirm_test=True)
  assert {p.name for p in (work/'DEV/Stage9_AKZ_Auto').glob('s9_*')}==published
 finally:target.write_bytes(original)
 assert all(sha(p)==h for p,h in dayhash.items())
 changed=copy.deepcopy(cfg);changed['climate']['air_c']=[16,25];p=OUT/'CHANGED_SETTINGS.json';save(p,changed)
 run('CHANGE','akz-settings-save',akz_settings=p)
 run('STALE','akz-auto-generate','FAIL',error='SETTINGS_CHANGED_AFTER_PREVIEW',preview=preview,confirm_test=True)
 new=run('CHANGED_CHECK','akz-auto-check',vom=s['vom'],variants='0,1,2,3')
 run('DAY_CONFLICT','akz-auto-generate','FAIL',error='DAY_SETTINGS_CONFLICT',preview=new['preview'],confirm_test=True)
 p=OUT/'DISABLED_SETTINGS.json';save(p,cfg|{'enabled':False});run('DISABLE','akz-settings-save',akz_settings=p)
 run('DISABLED_CHECK','akz-auto-check','FAIL',error='AKZ_AUTO_DISABLED',vom=s['vom'],variants='0,1,2,3')
 run('RESTORE_DEFAULTS','akz-settings-save',akz_settings=OUT/'DEFAULT_SETTINGS.json')
 assert all(sha(p)==h for p,h in dayhash.items())
 import win32file,win32con
 lock=work/'preferences/akz_auto/.lock';handle=win32file.CreateFile(str(lock),win32con.GENERIC_READ|win32con.GENERIC_WRITE,0,None,win32con.OPEN_ALWAYS,0,None)
 try:run('SETTINGS_LOCK','akz-settings-save','FAIL',error='LOCKED',akz_settings=OUT/'DEFAULT_SETTINGS.json')
 finally:handle.Close()
 full=run('FULL_GENERATE','generate',preview=preview,confirm_test=True);save(OUT/'FULL_BUNDLE.json',full)
 verified=run('FULL_VERIFY','verify',manifest=full['manifest']);assert verified['physical_files']==12 and verified['document_records']==13
 native=run('FULL_NATIVE','native-verify',manifest=full['manifest']);assert native['word_documents']==3 and native['excel_documents']==9
 assert native['rows']==104 and native['connections_m1_equal']
 save(OUT/'FROZEN_AUTO_FINAL.json',{'status':'PASS','exe_sha256':EXPECTED,'cases':cases,'GUI':'SEPARATE_PENDING','full_manifest':full['manifest'],'workspace':str(work)})
 print('Frozen AKZ series PASS',len(cases),flush=True)
except Exception as exc:
 save(OUT/'FROZEN_AUTO_FAILURE.json',{'status':'FAIL','error':repr(exc),'traceback':traceback.format_exc(),'cases':cases});raise
