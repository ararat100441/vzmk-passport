"""Bounded integration validation. Every mutation targets an isolated DEV fixture."""
from pathlib import Path
from copy import deepcopy
from unittest.mock import patch
import importlib,json,shutil,subprocess,sys,traceback,uuid
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT)); sys.path.insert(0,str(OUT)); sys.stdout.reconfigure(encoding='utf-8')
from profiles_stage4 import storage as s
from profiles_stage4.model import *
from profiles_stage4.offline import read_json,sha,no_network,load_sources
from profiles_stage4.windows_generation import ProtectedFile
from native_excel import run as native
from xlsx_tables import inspect

FILES=['Профили.xlsx',s.STATE,s.CATALOG,s.COMMIT]
RUN=uuid.uuid4().hex[:8]; tests=[]; network=[]; production=[]
def audit(event,args):
    if event in ('socket.connect','socket.getaddrinfo','socket.sendto'):
        network.append(event); raise AssertionError('Network forbidden')
    if event in ('open','os.rename','os.remove','os.mkdir'):
        for arg in args:
            if isinstance(arg,(str,bytes)) and any(p in str(arg).replace('\\','/') for p in ['/Работа/ВЗМК.xlsx','/Работа/Профили.xlsx']):
                production.append(str(arg)); raise AssertionError('Production forbidden')
sys.addaudithook(audit)
def hashes():
    paths=sorted(p for p in (ROOT/'profiles_stage4').glob('*') if p.is_file())
    paths += [ROOT/'stage4.py',ROOT/'STAGE4_README.md',ROOT/'Запустить Этап 4.cmd']+sorted(OUT.glob('*.py'))
    return {str(p.relative_to(ROOT)):sha(p) for p in paths}
def case(name,fn):
    try:
        detail=fn(); tests.append({'name':name,'status':'PASS','detail':detail}); print('PASS '+name,flush=True)
    except BaseException:
        tests.append({'name':name,'status':'FAIL','traceback':traceback.format_exc()}); s.dump(OUT/'targeted_tests.json',{'status':'FAIL','tests':tests}); raise
    s.dump(OUT/'targeted_tests.json',{'status':'RUNNING','tests':tests})
def protected():
    expected=read_json(OUT/'protected_before.json')
    for p,h in expected.items(): assert sha(p)==h,p
    for p,h in read_json(OUT/'generation_before.json').items(): assert sha(p)==h,p
    return {'protected_files':len(expected),'real_DEV_unchanged':True}
def fixture(name,seed=None):
    folder=OUT/'fixtures'/RUN/name/'DEV/Справочники'; folder.mkdir(parents=True)
    for n in FILES: shutil.copy2((seed or OUT/'baseline_generation')/n,folder/n)
    return folder/FILES[0]
def rejected(fn):
    try: fn()
    except (ValueError,OSError) as exc: return str(exc)
    raise AssertionError('Expected rejection')
def generation(t): return {n:sha(t.parent/n) for n in FILES}
def pairs(result): return [[x['L'],x['S']] for x in calculate_uzk(result)['rows']]

before=hashes(); s.dump(OUT/'product_hashes_before.json',before)
case('01 baseline integrity / accepted hash / protected evidence',protected)
old_data=read_json(OUT/'baseline_generation'/s.CATALOG)
baseline=Catalog(old_data['records'],old_data['sources'],old_data['user_state'])
state=read_json(OUT/'proposed_user_state.json')
catalog=Catalog(old_data['records'],old_data['sources'],state)
payload=s.workbook_payload(catalog)

def manual(family,key,nd,geometry):
    r=catalog.lookup(key,nd); assert r['status']=='FOUND'
    p=r['profile']; assert p['family']==family and p['geometry']==geometry and p['status']=='РУЧНОЙ'
    assert p['provenance'][0]['verification']['kind']=='KMD'
    assert 'ПОДТВЕРЖДЕНО ПО КМД' in p['notes']
    assert p['provenance'][0]['raw']['profile']==key
    return {'key':p['key'],'geometry':p['geometry'],'status':p['status'],'verification':p['provenance'][0]['verification']}
case('02 manual verified bent-channel add',lambda:manual('Гнутый швеллер','Гн.[140X60X4','ГОСТ 8278-83',{'h':140,'b':60,'t':4}))
case('03 manual verified closed-square add',lambda:manual('Квадратный профиль','Гн.[]120х3','ГОСТ 30245-2003',{'h':120,'b':120,'t':3}))
def raw_bent():
    r=catalog.lookup('Гн.[140X60X4'); assert r['status']=='FOUND' and r['profile']['uzk_rule']==BENT_RULE
    return r
case('04 exact raw KMD bent lookup',raw_bent)
def variants():
    values=[]
    for x in ['x','X','х','Х','×']:
        for text,expected in [(' Гн. [ 140'+x+'60'+x+'4 ',[400,4]),('Гн.[] 120 '+x+' 3',[480,3])]:
            r=catalog.lookup(text); assert r['status']=='FOUND' and pairs(r)==[expected]
            values.append({'raw':text,'key':r['key']})
    return values
case('05 x/X/х/Х/× and whitespace normalization',variants)
def open_closed():
    a=normalize_profile('Гн.[140×60×4'); b=normalize_profile('Гн.[]140×60×4')
    assert a!=b and b==normalize_profile('140х60х4')
    assert catalog.lookup('Гн.[140×60×4')['profile']['family']=='Гнутый швеллер'
    closed=catalog.lookup('Гн.[]140×60×4')
    assert closed['status']!='FOUND' or closed['profile']['family']=='Прямоугольный профиль'
    assert normalize_profile('Гн.[]120×3')==normalize_profile('120×120×3')
    return {'open_key':a,'closed_key':b,'square_expansion':'120×120×3'}
case('06 open vs closed identity collision prevention',open_closed)
def alias():
    r=catalog.lookup('[16П'); assert r['status']=='FOUND' and r['key']=='16П' and r['profile']['standard']=='ГОСТ 8240-97'
    assert catalog.lookup('[16П','ГОСТ 9999-99')['status']=='NOT_FOUND'
    return r
case('07 safe [16П alias / exact standard remains strict',alias)
def unsafe_alias():
    for text in ['[35Б1','[140×60×4','[Гн.[140×60×4','[[16П','[16П EXTRA']:
        assert catalog.lookup(text)['status']!='FOUND',text
    fake=deepcopy(catalog.raw_records[0]); fake.update(key='16П',display='16П',standard='ГОСТ TEST',family='Двутавр')
    assert Catalog(catalog.raw_records+[fake],catalog.sources,state).lookup('[16П')['status']=='AMBIGUOUS'
    return {'arbitrary_brackets_not_stripped':True,'multiple_target_identities':'AMBIGUOUS'}
case('08 alias arbitrary values / collisions rejected',unsafe_alias)
case('09 approved 16П single-item formula',lambda: {'UZK':pairs(alias())} if pairs(alias())==[[160,5],[128,8.4]] else (_ for _ in ()).throw(AssertionError()))
case('10 approved bent formula from geometry',lambda: {'UZK':pairs(raw_bent())} if pairs(raw_bent())==[[400,4]] else (_ for _ in ()).throw(AssertionError()))
def square():
    r=catalog.lookup('Гн.[]120х3'); assert r['status']=='FOUND' and pairs(r)==[[480,3]]
    return {'UZK':pairs(r),'rule':r['profile']['uzk_rule']}
case('11 approved closed-square formula',square)
def rectangular():
    r=catalog.lookup('80×40×4','ГОСТ 30245-2003'); assert r['status']=='FOUND' and pairs(r)==[[240,4]]
    return {'identity':[r['key'],r['profile']['standard']],'geometry':r['profile']['geometry'],'UZK':pairs(r)}
case('12 rectangular closed-profile formula',rectangular)
def same_s():
    r=deepcopy(catalog.lookup('16П','ГОСТ 8240-97')); r['profile']['geometry']['s']=8.4
    assert pairs(r)==[[288,8.4]]
    assert pairs(alias())==[[160,5],[128,8.4]]
    return {'synthetic_equal_S_within_one_profile':pairs(r),'separate_calls_not_accumulated':True}
case('13 same thickness groups within one item only',same_s)
def frequencies(values):
    result={str(x):frequency_angle(x) for x in values}
    for x in values: assert result[str(x)]=={'frequency_mhz':5 if x<12 else 2.5,'angle_degrees':70 if x<12 else 65}
    return result
case('14 frequencies 3/4/5/8.4',lambda:frequencies([3,4,5,8.4]))
case('15 threshold 11.5 / 12',lambda:frequencies([11.5,12]))
def assignment_limits():
    rows=[r for r in catalog.rows if r['uzk_rule']==CLOSED_RULE]
    assert len(rows)==493 and all(r['standard']=='ГОСТ 30245-2003' for r in rows)
    assert sum(r['uzk_rule']==CHANNEL_RULE for r in catalog.rows)==1
    assert sum(r['uzk_rule']==BENT_RULE for r in catalog.rows)==1
    assert catalog.lookup('20П','ГОСТ 8240-97')['profile']['uzk_rule']==NOT_CONFIGURED
    new_state=s.prepare_manual_state(catalog,'124×124×3','Квадратный профиль','ГОСТ 30245-2003',{'h':124,'b':124,'t':3})
    other=Catalog(catalog.raw_records,catalog.sources,new_state)
    assert other.lookup('124×124×3')['profile']['uzk_rule']==NOT_CONFIGURED
    no_approval=deepcopy(state); no_approval.pop('uzk_approvals')
    assert Catalog(catalog.raw_records,catalog.sources,no_approval).lookup('Гн.[140X60X4')['profile']['uzk_rule']==NOT_CONFIGURED
    return {'closed_assigned':len(rows),'unconfirmed_manual_no_rule':True,'other_channel_no_rule':True,'approval_required':True}
case('assignment scopes / unconfirmed manual safety',assignment_limits)
def new_conflict():
    row=deepcopy(next(r for r in catalog.raw_records if r['key']=='80×40×4' and r['standard']=='ГОСТ 30245-2003'))
    row['record_id']='TEST-CONFLICT'; row['geometry']['t']=3
    other=Catalog(catalog.raw_records+[row],catalog.sources,state)
    r=other.lookup('80×40×4','ГОСТ 30245-2003')
    assert r['status']=='CONFLICT' and r['profile']['uzk_rule']==NOT_CONFIGURED and not r['profile']['geometry']
    rejected(lambda:calculate_uzk(r))
    return {'new_rule_cannot_override_geometry_conflict':True}
case('approved scope cannot enable conflicting identity',new_conflict)
def m1():
    assert catalog.conflicts==baseline.conflicts
    for c in catalog.conflicts:
        r=catalog.lookup(c['key'],c['standard'])
        assert r['status']==('CONFLICT' if c['classification']=='BLOCKING' else 'FOUND')
        assert r['profile']['unresolved_fields']==c['fields']
    assert catalog.counts()['blocking_disagreements']==8 and catalog.counts()['non_blocking_disagreements']==7
    old={(r['key'],r['standard']):r for r in baseline.rows}
    for r in catalog.rows:
        if (r['key'],r['standard']) in old:
            a=deepcopy(r); b=deepcopy(old[(r['key'],r['standard'])]); a.pop('uzk_rule'); b.pop('uzk_rule'); assert a==b
    return {'blocking':8,'non_blocking':7,'existing_geometry_status_provenance_unchanged':2849}
case('19 M1 smoke / all old canonical data unchanged',m1)
def golden():
    result={}
    for nd,expected in [('ГОСТ Р 57837-2017',[[328,6],[348,9]]),('ГОСТ 26020-83',[[329,6.2],[310,8.5]])]:
        r=catalog.lookup('35Б1',nd); assert r['status']=='FOUND' and pairs(r)==expected
        result[nd]={'geometry':r['profile']['geometry'],'UZK':pairs(r)}
    assert catalog.lookup('35Б1')['status']=='FOUND_PREFERRED'
    return result
case('21 35Б1 standard-aware golden / preference',golden)
def beam20(): assert catalog.lookup('20Ш1')['status']=='AMBIGUOUS'; return 'AMBIGUOUS'
case('22 20Ш1 remains AMBIGUOUS',beam20)
def ingest():
    with no_network() as attempts: records,sources=load_sources()
    assert not attempts and records==old_data['records'] and sources==old_data['sources']
    return {'raw':len(records),'source_data_unchanged':True,'network_attempts':0}
case('offline Source A/B exact ingestion',ingest)

isolated=fixture('extended')
def generated():
    previous=generation(isolated)
    with no_network() as attempts: result=s.build(isolated,user_state=state,previews=OUT/'previews')
    assert not attempts and result['network_attempts']==0
    assert {n:sha(Path(result['backup'])/n) for n in FILES}==previous
    current=s.load_catalog(isolated); assert current.counts()==catalog.counts()
    assert current.user_state['manual_rows']==state['manual_rows'] and current.user_state['uzk_approvals']==state['uzk_approvals']
    assert s.verify_workbook(isolated,payload)['formulas']==0
    structure=inspect(isolated,OUT/'xlsx_xml'); assert structure['valid'],structure
    s.dump(OUT/'xlsx_structure.json',structure); s.dump(OUT/'isolated_build.json',result)
    return result
case('isolated batch build / safe backup / numeric reopen / 5 tables',generated)
def rebuild():
    result=s.build(isolated)
    current=s.load_catalog(isolated)
    assert current.user_state['manual_rows']==state['manual_rows']
    assert current.user_state['uzk_approvals']==state['uzk_approvals']
    assert current.user_state['preferences']==state['preferences']
    assert current.lookup('[16П')['status']=='FOUND'
    assert pairs(current.lookup('Гн.[140X60X4'))==[[400,4]] and pairs(current.lookup('Гн.[]120х3'))==[[480,3]]
    assert current.lookup('35Б1')['status']=='FOUND_PREFERRED'
    s.dump(OUT/'isolated_rebuild.json',result)
    return {'manual_survives':2,'approved_scopes_survive':3,'alias_survives':True,'35_preference_survives':True,'network_attempts':result['network_attempts']}
case('16/17/18 rebuild preserves manual / rules / alias / preference',rebuild)
case('23 native Excel isolated normal Open with alerts / no repair',lambda:native(isolated,'isolated_'+RUN))

template=OUT/'verified_template.xlsx'; shutil.copy2(isolated,template)
def copied_render(input,destination,previews=None):
    assert read_json(input)==payload
    shutil.copy2(template,destination); return {'network_attempts':0}
def rollback():
    results=[]
    for fail_at in ['before_commit','after_first_replace']:
        t=fixture('rollback_'+fail_at,isolated.parent); previous=generation(t)
        with patch.object(s,'render_workbook',copied_render): message=rejected(lambda:s.build(t,fail_at=fail_at))
        assert generation(t)==previous and s.load_catalog(t).user_state==s.load_catalog(isolated).user_state
        results.append({'failure':fail_at,'message':message,'previous_generation_preserved':True})
    return results
case('20 M2 safe write failure / rollback unchanged',rollback)
def external_race(name):
    t=fixture('external_'+str(FILES.index(name)),isolated.parent); previous=generation(t); edited={}
    def render(input,destination,previews=None):
        result=copied_render(input,destination); p=t.parent/name
        if name.endswith('.xlsx'):
            with open(p,'ab') as f:f.write(b'EXTERNAL_EDIT')
        else:
            value=read_json(p); value['external_probe']='EXTERNAL_EDIT'; s.dump(p,value)
        edited['hash']=sha(p); return result
    with patch.object(s,'render_workbook',render): message=rejected(lambda:s.build(t))
    final=generation(t); assert final[name]==edited['hash']!=previous[name]
    assert all(final[n]==previous[n] for n in FILES if n!=name) and not (t.parent/'backups').exists()
    return {'file':name,'STOP':message,'external_bytes_preserved':True,'backup_not_created':True}
for name in FILES: case('20 M2 external race '+name,lambda n=name:external_race(n))
def boundary():
    t=fixture('commit_boundary',isolated.parent); original=ProtectedFile.rename; attempts=[]
    def rename(handle,destination):
        if not attempts:
            p=subprocess.run([sys.executable,'-I','-B',str(OUT/'external_writer.py'),str(t.parent)],capture_output=True,text=True,encoding='utf-8',timeout=20)
            assert p.returncode==0,p.stderr
            attempts.extend(json.loads(p.stdout)); assert len(attempts)==12 and all(x['blocked'] for x in attempts)
        return original(handle,destination)
    with patch.object(s,'render_workbook',copied_render),patch.object(ProtectedFile,'rename',rename): result=s.build(t)
    expected=deepcopy(s.load_catalog(isolated).user_state); expected['revision']+=1
    assert s.load_catalog(t).user_state==expected
    s.dump(OUT/'commit_boundary.json',attempts)
    return {'independent_writer_attempts_blocked':12,'normal_commit':'PASS'}
case('20 M2 protected commit boundary independent process',boundary)
def manual_api():
    t=fixture('manual_api',isolated.parent)
    result=s.add_manual('Гн.[150×50×3','Гнутый швеллер','ГОСТ 8278-83',{'h':150,'b':50,'t':3},'UNCONFIRMED',workbook=t)
    r=s.load_catalog(t).lookup('Гн.[150×50×3'); assert r['status']=='FOUND' and r['profile']['uzk_rule']==NOT_CONFIGURED
    return {'manual_add_API':'PASS','no_unapproved_rule':True,'network_attempts':result['network_attempts']}
case('manual-add API unconfirmed bent remains without rule',manual_api)
def smoke():
    for module in ['main','stage2','stage3','passport_app.catalog']: importlib.import_module(module)
    checkpoint=read_json(ROOT/'DEV/DEV_REGISTRY_CHECKPOINT.json'); assert checkpoint
    return {'imports':['main','stage2','stage3','passport_app.catalog'],'checkpoint_readable':True,'full_regression':'NOT_RUN'}
case('Stage 1–3 imports / checkpoint readable',smoke)
case('24 DEV-only / no-network / protected evidence final',protected)
after=hashes(); s.dump(OUT/'product_hashes_after.json',after); assert before==after
assert not network and not production
s.dump(OUT/'targeted_tests.json',{'status':'PASS','tests':tests,'passed':len(tests),'failed':0,'product_hashes_before_after_equal':True,'network_events':network,'production_events':production,'real_DEV_rebuilds':0,'isolated_workbook':str(isolated)})
print(json.dumps({'status':'PASS','passed':len(tests),'counts':catalog.counts(),'real_DEV_rebuilds':0},ensure_ascii=False),flush=True)
