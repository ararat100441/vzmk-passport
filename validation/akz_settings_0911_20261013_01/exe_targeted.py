from common import *
from stage9_app import akz_settings as K
import copy,win32api,win32con
work=Path(read(OUT/'WORKSPACE_MAIN.json')['workspace']);cases=[]
def call(label,name,expected='PASS',workspace=work,error=None,**args):
    d,p=task('EXE_'+label,name,expected,workspace=workspace,**args)
    if error:assert error in d.get('error',''),d
    cases.append({'case':label,'task':name,'result':str(p),'expected':expected,'actual':d['status']})
    save(OUT/('CHECKPOINT_'+label+'.json'),cases[-1]);return d
def pinned(manifest):
    m=read(manifest);return {str(manifest):sha(manifest),**{r['path']:sha(r['path']) for r in m['records']},**{p:sha(p) for p in m['inputs'] if Path(p).name=='PAIR.json'},**({m['akz_settings_snapshot']['path']:sha(m['akz_settings_snapshot']['path'])} if 'akz_settings_snapshot' in m else {})}
def unchanged(files):assert all(sha(p)==h for p,h in files.items())
settings=work/'preferences/akz_auto/settings.json';cfg=read(settings);save(OUT/'SETTINGS_A.json',cfg)
check=call('A_CHECK','check',vom=SETUP['vom'],variants='0,1,2,3');a=call('A_GENERATE','generate',preview=check['preview'],confirm_test=True)['manifest']
save(OUT/'BUNDLE_A.json',{'manifest':a});a_files=pinned(a);m=read(a);snap=m['akz_settings_snapshot'];assert str(settings) not in m['inputs'];assert m['inputs'][snap['path']]==snap['sha256']
days={str(p):sha(p) for p in settings.parent.glob('days/*.json')};before=sha(settings)
numeric=copy.deepcopy(cfg)
for key in ('air_c','surface_c','relative_humidity'):numeric['climate'][key]=[float(v) for v in numeric['climate'][key]]
for n,pairs in numeric['thickness'].items():numeric['thickness'][n]=[[float(v) for v in p] for p in pairs]
save(OUT/'SETTINGS_FLOAT_EQUIVALENT.json',numeric)
call('SAVE_EQUIVALENT','akz-settings-save',akz_settings=OUT/'SETTINGS_FLOAT_EQUIVALENT.json');assert sha(settings)==before
repeat=call('A_REPEAT','akz-auto-generate',preview=check['preview'],confirm_test=True);assert all(sha(p)==h for p,h in days.items());unchanged(a_files)
call('A_VERIFY_EQUIVALENT','verify',manifest=a);call('A_NATIVE','native-verify',manifest=a)
changed=copy.deepcopy(cfg);changed['climate']['air_c']=[16,26];save(OUT/'SETTINGS_B.json',changed)
call('SAVE_B','akz-settings-save',akz_settings=OUT/'SETTINGS_B.json');assert K.digest(cfg)!=K.digest(changed)
call('STALE_A_PREVIEW','generate','FAIL',preview=check['preview'],confirm_test=True,error='SETTINGS_CHANGED_AFTER_PREVIEW')
same=call('SAME_DATE_CHECK','check',vom=SETUP['vom'],variants='0,1,2,3')
call('REAL_RANGE_SAME_DATE_STOP','akz-auto-generate','FAIL',preview=same['preview'],confirm_test=True,error='DAY_SETTINGS_CONFLICT');assert all(sha(p)==h for p,h in days.items())
# Accepted Stage 1 creates a new TEST passport on today's date in this isolated
# workspace. No registry/template/VOM/catalog content is hand-edited.
bcheck=call('B_NEW_CHECK','check',vom=SETUP['vom'],variants='0,1,2,3',passport='10/2026',new=True)
b=call('B_GENERATE','generate',preview=bcheck['preview'],confirm_test=True)['manifest'];save(OUT/'BUNDLE_B.json',{'manifest':b});b_files=pinned(b)
assert read(read(b)['akz_settings_snapshot']['path'])['climate']['air_c']==[16,26]
assert read(snap['path'])['climate']['air_c']==[15,25]
for label,manifest in [('A_REOPEN',a),('B_REOPEN',b),('A_AFTER_RESTART',a),('B_AFTER_RESTART',b)]:
    proof=call(label,'verify',manifest=manifest)['settings_validation'];assert proof['status']=='FROZEN_SETTINGS_VERIFIED'
call('B_NATIVE','native-verify',manifest=b);unchanged(a_files);unchanged(b_files)
# Fault injection only in this disposable TEST snapshot; old receipts/manifests
# remain byte-identical. Restore exact bytes, with explicit before/after proof.
path=Path(snap['path']);raw=path.read_bytes();attributes=win32api.GetFileAttributes(str(path))
try:
    win32api.SetFileAttributes(str(path),win32con.FILE_ATTRIBUTE_NORMAL);path.write_bytes(raw+b' ')
    call('FROZEN_TAMPER_STOP','verify','FAIL',manifest=a,error='FROZEN_SETTINGS_SHA_CONFLICT')
finally:path.write_bytes(raw);win32api.SetFileAttributes(str(path),attributes)
call('A_AFTER_TEST_RESTORE','verify',manifest=a);unchanged(a_files);unchanged(b_files)
save(OUT/'AB_FINAL.json',{'status':'PASS','runtime':'FROZEN_EXE_0.9.11','cases':cases,'A_manifest':a,'B_manifest':b,'A_files':a_files,'B_files':b_files,'A_day_files':days,'new_B_dates':{str(p):read(p)['day'] for p in settings.parent.glob('days/*.json') if str(p) not in days},'GUI':'NOT_RUN','restart_method':'Every CLI command starts a fresh EXE process; GUI relaunch NOT_RUN','old_manifests_receipts_changed':False})
print('Mandatory A-L frozen CLI scenario PASS',len(cases),'checks',flush=True)
