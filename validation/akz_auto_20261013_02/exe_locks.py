from common import *
import win32file,win32con
report=read(OUT/'EXE_STAGE8_FINAL.json');fixture=read([s['result'] for s in report['steps'] if s['task']=='fixture'][-1]);sb=Path(fixture['sandbox']);reg=sb/'registry/УЗК_TEST.xlsm';cfg=WORKSPACE/'preferences/akz_auto/settings.json'
initial={'registry':sha(reg),'settings':sha(cfg)};cases=[]
for label,target,name,args in [
 ('SANDBOX',sb/'.stage8.lock','close-preview',{'sandbox':sb,'request':fixture['request']}),
 ('XLSM',reg,'close-preview',{'sandbox':sb,'request':fixture['request']}),
 ('AKZ_SETTINGS',cfg.parent/'.lock','akz-settings-save',{'akz_settings':OUT/'SETTINGS_VALID.json'})]:
    handle=win32file.CreateFile(str(target),win32con.GENERIC_READ|win32con.GENERIC_WRITE,0,None,win32con.OPEN_ALWAYS if label!='XLSM' else win32con.OPEN_EXISTING,0,None)
    try:d,p,s=task('EXE_LOCK_'+label,name,'FAIL',**args)
    finally:handle.Close()
    assert ('LOCKED' in d['error'] or 'PARALLEL_OPERATION' in d['error']
            or label=='XLSM' and 'REGISTRY_GUARD_OR_IO_FAILURE' in d['error'] and '(32,' in d['error']),d
    assert sha(reg)==initial['registry'] and sha(cfg)==initial['settings']
    cases.append({'case':label,'result':str(p),'actual_status':'FAIL','safe_STOP':'PASS','error':d['error'],'no_registry_or_settings_write':True})
save(OUT/'EXE_LOCKS_FINAL.json',{'status':'PASS','cases':cases,'before_sha256':initial,'after_sha256':{'registry':sha(reg),'settings':sha(cfg)}})
print('Frozen locks: 3/3 safe STOP PASS',flush=True)
