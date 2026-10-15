"""Isolated state, rebuild, integrity and fault-injection review."""
import sys, os, json, hashlib, shutil, socket, importlib, ctypes
from pathlib import Path
from copy import deepcopy
from unittest.mock import patch
sys.stdout.reconfigure(encoding='utf-8')
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
events=[]
def audit(event,args):
    if event in ('socket.connect','socket.getaddrinfo','socket.sendto'):
        events.append(['network',event]);raise RuntimeError('Review denied network')
    if event in ('open','os.rename','os.remove','os.mkdir','shutil.copyfile','os.rmdir'):
        for arg in args:
            if not isinstance(arg,(str,bytes)):continue
            p=Path(os.fsdecode(arg)).resolve()
            if p in (ROOT.parent.parent/'ВЗМК.xlsx',ROOT.parent.parent/'Профили.xlsx'):
                events.append(['production',event]);raise RuntimeError('Review denied production')
        if event=='open' and isinstance(args[0],(str,bytes)):
            mode,flags=args[1:3]; writing=(isinstance(mode,str) and any(c in mode for c in 'wax+')) or (isinstance(flags,int) and bool(flags&(os.O_WRONLY|os.O_RDWR|os.O_CREAT|os.O_TRUNC)))
            if writing and not Path(os.fsdecode(args[0])).resolve().is_relative_to(OUT):raise RuntimeError('Write outside review denied')
        if event in ('os.rename','os.remove','os.mkdir','os.rmdir'):
            for arg in args:
                if isinstance(arg,(str,bytes)) and not Path(os.fsdecode(arg)).resolve().is_relative_to(OUT):raise RuntimeError('Mutation outside review denied')
sys.addaudithook(audit)
from profiles_stage4 import storage as s, offline as o, model as m
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def save(name,data):(OUT/name).write_text(json.dumps(data,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
reviewed=read(OUT/'reviewed_hashes_before.json')
assert all(sha(p)==h for p,h in reviewed.items())
FILES=['Профили.xlsx',s.STATE,s.CATALOG,s.COMMIT]
REAL=s.DEFAULT_WORKBOOK.parent
def clone(name):
    folder=OUT/name/'DEV/Справочники';folder.mkdir(parents=True,exist_ok=True)
    for f in FILES:shutil.copy2(REAL/f,folder/f)
    return folder/'Профили.xlsx'
def hashes(target):return {f:sha(target.parent/f) for f in FILES}
results=[]
def result(name,status='PASS',**details):
    results.append({'name':name,'status':status,**details});save('fixture_results.json',results)
    print(json.dumps(results[-1],ensure_ascii=False),flush=True)
def rejected(fn):
    try:fn()
    except (ValueError,OSError,KeyError,TypeError) as e:return {'type':type(e).__name__,'message':str(e)}
    raise AssertionError('Operation unexpectedly accepted')

# Real offline renderer/build: preserve all user-maintained state.
target=clone('state_survival')
s.add_manual('Ручной обзор Ω 1','Двутавр','ТЕСТ НД REVIEW',{'h':100,'b':50,'s':4,'t':6},'Independent isolated fixture',workbook=target)
r=s.lookup('Ручной обзор Ω 1','ТЕСТ НД REVIEW',target)
assert r['profile']['status']=='РУЧНОЙ' and r['profile']['uzk_rule']==m.NOT_CONFIGURED
result('manual_add_no_automatic_uzk')
s.set_preference('20Ш1','ГОСТ 26020-83','Independent fixture preference',workbook=target)
cat=s.load_catalog(target);c=next(c for c in cat.conflicts if c['key']=='L125×80×7')
candidate=next(x for x in c['members'] if x['source']=='PTOflow')
s.resolve_conflict(c['id'],candidate['record_id'],'Independent fixture explicit resolution',workbook=target)
before=read(target.parent/s.STATE);build_result=s.build(target);after=read(target.parent/s.STATE)
assert all(before[k]==after[k] for k in ('manual_rows','preferences','resolutions'))
assert after['revision']==before['revision']+1
assert s.lookup('20Ш1',workbook=target)['status']=='FOUND_PREFERRED'
assert s.lookup('L125×80×7','ГОСТ 8510-86',target)['status']=='FOUND'
assert s.lookup('Ручной обзор Ω 1',workbook=target)['profile']['uzk_rule']==m.NOT_CONFIGURED
backup=Path(build_result['backup']);assert read(backup/s.STATE)==before
assert all((backup/f).is_file() for f in FILES)
result('real_offline_rebuild_preserves_manual_preference_resolution',network=build_result['network_attempts'],backup=str(backup))
before_hash=hashes(target)
rejected(lambda:s.add_manual('Ручной обзор Ω 1','Двутавр','ТЕСТ НД REVIEW',{'h':100,'b':50,'s':4,'t':6},workbook=target))
assert hashes(target)==before_hash
result('duplicate_manual_identity_rejected_without_change')

# State with an invalid preference is a blocked ambiguous lookup; conflict preference cannot resolve geometry.
cat=s.load_catalog(target)
bad=deepcopy(cat.user_state);bad['preferences']=[{'key':'35Б1','standard':'НЕСУЩЕСТВУЮЩИЙ НД','basis':'fixture','status':'ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ','date':'2026-10-07'}]
assert m.Catalog(cat.raw_records,cat.sources,bad).lookup('35Б1')['status']=='AMBIGUOUS'
bad['preferences'] += [{'key':'35Б1','standard':'ГОСТ 26020-83','basis':'fixture','status':'ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ','date':'2026-10-07'}]
assert m.Catalog(cat.raw_records,cat.sources,bad).lookup('35Б1')['status']=='AMBIGUOUS'
confcat=s.load_catalog();bad=deepcopy(confcat.user_state)
bad['preferences'].append({'key':'L125×80×7','standard':'ГОСТ 8510-86','basis':'fixture','status':'ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ','date':'2026-10-07'})
assert m.Catalog(confcat.raw_records,confcat.sources,bad).lookup('L125×80×7')['status']=='CONFLICT'
result('invalid_duplicate_conflicting_preference_safe')

# Real conflicting manual identity collision, committed only inside fixture.
target_collision=clone('manual_collision')
s.add_manual('35Б1','Двутавр','ГОСТ Р 57837-2017',{'h':347,'b':174,'s':6,'t':9},'Collision fixture',workbook=target_collision)
r=s.lookup('35Б1','ГОСТ Р 57837-2017',target_collision)
assert r['status']=='CONFLICT' and len(r['profile']['provenance'])==2
result('manual_generated_geometry_collision_blocked')

# Integrity changes existing before build must block both lookup and rebuild, retaining exact changed bytes.
for f in ('Профили.xlsx',s.CATALOG,s.STATE):
    t=clone('external_before_'+str(FILES.index(f)));p=t.parent/f
    p.write_bytes(p.read_bytes()+b'\nREVIEW_EXTERNAL_CHANGE')
    before_hash=hashes(t)
    e1=rejected(lambda:s.load_catalog(t));e2=rejected(lambda:s.build(t))
    assert hashes(t)==before_hash
    result('external_change_before_build_'+f,lookup=e1,rebuild=e2,preserved=True)
for f in ('Профили.xlsx',s.CATALOG,s.STATE,s.COMMIT):
    t=clone('missing_'+str(FILES.index(f)));(t.parent/f).unlink()
    surviving={x:sha(t.parent/x) for x in FILES if (t.parent/x).exists()}
    e=rejected(lambda:s.build(t));assert surviving=={x:sha(t.parent/x) for x in surviving}
    result('missing_file_'+f,rejected=e)
t=clone('malformed_manifest');(t.parent/s.COMMIT).write_text('{broken',encoding='utf-8');before_hash=hashes(t)
e=rejected(lambda:s.build(t));assert hashes(t)==before_hash;result('malformed_json_no_overwrite',rejected=e)
for x in ('NaN','Infinity','-Infinity','-1',True,'bad'):
    t=target;before_hash=hashes(t)
    rejected(lambda:s.add_manual('Invalid value','Круглая труба','REVIEW',{'D':100,'s':x},workbook=t))
    assert hashes(t)==before_hash
result('invalid_numbers_no_state_mutation')
for path in (ROOT/'Профили.xlsx',ROOT.parent.parent/'Профили.xlsx',ROOT/'validation'/'..'/'outside'/'DEV'/'Справочники'/'Профили.xlsx'):
    rejected(lambda:s.guard_target(path))
result('destination_path_traversal_guard')

# Transaction fault injection uses a byte-identical previously verified workbook as renderer output.
# Payload equality is checked first. These cases test transaction/guard logic, not rendering again.
original_payload=s.workbook_payload(s.load_catalog())
def copy_verified_renderer(payload,destination,previews=None):
    assert read(payload)==original_payload
    shutil.copy2(s.DEFAULT_WORKBOOK,destination)
    return {'network_attempts':0}
for point in ('before_commit','after_first_replace'):
    t=clone('fault_'+point);before_hash=hashes(t)
    with patch.object(s,'render_workbook',copy_verified_renderer):e=rejected(lambda:s.build(t,fail_at=point))
    assert hashes(t)==before_hash and s.lookup('35Б1',workbook=t)['status']=='FOUND_PREFERRED'
    assert not (t.parent/'.profiles.lock').exists()
    result('rollback_'+point,rejected=e,preserved=True)
t=clone('keyboard_interrupt');before_hash=hashes(t);orig_replace=s.os.replace;calls=0
def interrupted(src,dst):
    global calls
    calls+=1
    if calls==2:raise KeyboardInterrupt('Review injection after first replacement')
    return orig_replace(src,dst)
with patch.object(s,'render_workbook',copy_verified_renderer),patch.object(s.os,'replace',interrupted):
    try:s.build(t)
    except KeyboardInterrupt:pass
    else:raise AssertionError('Interrupt injection missing')
assert hashes(t)==before_hash and not (t.parent/'.profiles.lock').exists()
result('keyboard_interrupt_rolls_back_complete_generation')

# Actual Windows file lock permits reads but denies replace, as an open Excel workbook can do.
t=clone('locked_workbook');before_hash=hashes(t)
kernel=ctypes.WinDLL('kernel32',use_last_error=True)
kernel.CreateFileW.argtypes=[ctypes.c_wchar_p,ctypes.c_uint32,ctypes.c_uint32,ctypes.c_void_p,ctypes.c_uint32,ctypes.c_uint32,ctypes.c_void_p]
kernel.CreateFileW.restype=ctypes.c_void_p
kernel.CloseHandle.argtypes=[ctypes.c_void_p]
handle=kernel.CreateFileW(str(t),0x80000000,1,None,3,0x80,None)
assert handle not in (None,ctypes.c_void_p(-1).value)
try:
    with patch.object(s,'render_workbook',copy_verified_renderer):e=rejected(lambda:s.build(t))
finally:kernel.CloseHandle(handle)
assert hashes(t)==before_hash and s.lookup('35Б1',workbook=t)['status']=='FOUND_PREFERRED'
result('locked_workbook_no_overwrite',rejected=e,preserved=True)

# External edits arriving during the renderer window: previous hash check has already passed.
for f in ('Профили.xlsx',s.CATALOG,s.STATE):
    t=clone('external_during_'+str(FILES.index(f)));modified={}
    def racing_renderer(payload,destination,previews=None):
        out=copy_verified_renderer(payload,destination,previews)
        p=t.parent/f
        p.write_bytes(p.read_bytes()+b'\nREVIEW_CHANGE_DURING_BUILD')
        modified['hash']=sha(p)
        return out
    with patch.object(s,'render_workbook',racing_renderer):out=s.build(t)
    overwritten=sha(t.parent/f)!=modified['hash']
    assert overwritten
    result('external_change_during_build_'+f,'FINDING',changed_hash=modified['hash'],final_hash=sha(t.parent/f),
      silently_overwritten=overwritten,backup_retained_changed_bytes=sha(Path(out['backup'])/f)==modified['hash'])

# Missing cache and malformed source row are fail-closed; sources are never mutated.
original_sha=o.sha
def missing_cache(p):
    if str(p).endswith('09_gost_26020-83.html'):raise FileNotFoundError('Review simulated missing cache')
    return original_sha(p)
with patch.object(o,'sha',missing_cache):e=rejected(lambda:o.load_sources())
result('missing_source_cache_blocks',rejected=e)
original_read=o.read_json
def malformed_source(p):
    d=original_read(p)
    if Path(p).name=='parsed_source_records.json':d[0]['parsed']['t']='NaN'
    return d
with patch.object(o,'read_json',malformed_source):e=rejected(lambda:o.load_sources())
result('malformed_source_row_blocks',rejected=e)
assert all(sha(p)==h for p,h in reviewed.items())
assert not events
save('reviewed_hashes_after_fixtures.json',{p:sha(p) for p in reviewed})
save('fixture_summary.json',{'checks':len(results),'pass':sum(r['status']=='PASS' for r in results),
 'finding_reproductions':sum(r['status']=='FINDING' for r in results),'network_production_events':events,'reviewed_originals_unchanged':True})
print('COMPLETE',flush=True)
