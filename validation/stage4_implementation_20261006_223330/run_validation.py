"""Targeted Stage 4 integration tests, bound to exact tested product hashes."""
import importlib
import json
from pathlib import Path
import socket
import sys
from copy import deepcopy
from unittest.mock import patch
import uuid

sys.stdout.reconfigure(encoding='utf-8')
OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
sys.path.insert(0, str(ROOT))
from profiles_stage4.model import Catalog, ProfileError, RULE, NOT_CONFIGURED, calculate_uzk, frequency_angle, normalize_profile, number
from profiles_stage4.offline import load_sources, ingest_a, ingest_b, sha, read_json, no_network, A_REL, B_REL
from profiles_stage4.storage import (build, load_catalog, lookup, add_manual, set_preference, resolve_conflict,
    confirm_manual_uzk, verify_workbook, workbook_payload, guard_target, dump, DEFAULT_WORKBOOK, STATE, CATALOG, COMMIT, HEADERS, CORE_COLUMNS)

def hashes():
    products = sorted([p for p in (ROOT/'profiles_stage4').glob('*') if p.is_file()] +
        [ROOT/'stage4.py', ROOT/'Запустить Этап 4.cmd', ROOT/'STAGE4_README.md', OUT/'run_validation.py', OUT/'finalize.py'])
    return {str(p.relative_to(ROOT)): sha(p) for p in products}
def require(value, message='Assertion failed'):
    if not value:
        raise AssertionError(message)
def raises(fn):
    try:
        fn()
    except (ValueError, OSError):
        return True
    return False

tests = []
network_events = []
production_events = []
def audit(event, args):
    if event in ('socket.connect', 'socket.getaddrinfo', 'socket.sendto'):
        network_events.append(event)
        raise ProfileError('Test audit hook: network denied')
    if event in ('open', 'os.rename', 'os.remove', 'os.mkdir'):
        for arg in args:
            if isinstance(arg, (str, bytes)):
                name = str(arg).replace('\\', '/')
                if '/Работа/ВЗМК.xlsx' in name or '/Работа/Профили.xlsx' in name:
                    production_events.append((event, name))
                    raise ProfileError('Test audit hook: production denied')
sys.addaudithook(audit)

before = hashes()
dump(OUT/'product_before.json', before)
protected = read_json(OUT/'protected_before.json')
require(all(sha(p) == h for p,h in protected.items()), 'Protected baseline changed BEFORE tests')

def test(i, title, fn):
    try:
        fn()
        tests.append({'id': i, 'name': title, 'status': 'PASS'})
    except Exception as exc:
        tests.append({'id': i, 'name': title, 'status': 'FAIL', 'error': str(exc)})
        print('FAIL', i, title, str(exc), flush=True)

def geometry(result):
    return result['profile']['geometry']
def core_tuple(result):
    return tuple(geometry(result)[f] for f in ('h','b','s','t'))

try:
    with no_network() as attempts:
        a, sa = ingest_a()
        b, sb = ingest_b()
        records, sources = a+b, sa+sb
        cat = Catalog(records, sources)
        fixture = OUT / ('tests_'+uuid.uuid4().hex[:8]) / 'DEV/Справочники/Профили.xlsx'
        first_build = build(fixture)
    test(1, 'Offline KM-KMD XLS origins / accepted datasets', lambda: require(len(a)>0 and set(r['standard'] for r in a)==set(s['standard'] for s in sa if s['type']!='ZIP')))
    test(2, 'Offline PTOflow cached HTML numeric group rows', lambda: require(len(b)>0 and len(sb)==9 and all(r['source']=='PTOflow' for r in b)))
    test(3, 'No network during ingestion/build/renderer', lambda: require(not attempts and not network_events and first_build['network_attempts']==0))
    local = load_catalog(fixture)
    payload = workbook_payload(local)
    test(4, 'Exact required sheet names/order', lambda: require(list(payload['sheets']) == ['Профили','Источники','Правила УЗК','Предпочтения','Конфликты']))
    test(5, 'Exact core columns/order', lambda: require(payload['sheets']['Профили'][0][:15]==CORE_COLUMNS))
    test(6, 'Excel numeric dimensions, complete reopen comparison', lambda: require(verify_workbook(fixture,payload)['numeric_geometry']))
    old=cat.lookup('35Б1','ГОСТ 26020-83')
    modern=cat.lookup('35Б1','ГОСТ Р 57837-2017')
    preferred=cat.lookup('35Б1')
    test(7, 'Decimals 6.2 / 8.5 and decimal comma', lambda: require(geometry(old)['s']==number('6,2')==6.2 and geometry(old)['t']==8.5))
    rounded=[r for r in b if r['display_rounding']]
    def exact_numeric():
        require(len(rounded)>0)
        row=next(r for r in b if r['standard']=='ГОСТ 26020-83' and r['key']=='12Б1')
        require(row['geometry']['h']==117.6 and number(row['raw']['h']['display'])==118)
        require(geometry(cat.lookup('12Б1','ГОСТ 26020-83'))['h']==117.6)
    test(8, 'PTOflow sort numeric, never rounded display', exact_numeric)
    test(9, 'Identity is key + ND; variants remain independent', lambda: require(len(cat.by_key['35Б1'])==3 and core_tuple(old)!=core_tuple(modern)))
    def duplicate():
        raw=[r for r in b if r['key']=='1420×20' and r['standard']=='ГОСТ 10704-91']
        row=cat.lookup('1420×20','ГОСТ 10704-91')['profile']
        require(len(raw)==2 and len(row['provenance'])==2 and len({r['record_id'] for r in row['provenance']})==2)
    test(10, 'Identical duplicate raw rows retained; one canonical identity', duplicate)
    known={'L125×80×7','L125×80×8','L75×50×5','L75×50×6','L100×8','L140×90×8'}
    test(11, 'Disagreement creates conflict including all six known mismatches', lambda: require(known <= {c['key'] for c in cat.conflicts}))
    test(12, 'Conflict lookup blocked; canonical dimensions blank; no initial resolution', lambda: require(all(cat.lookup(c['key'],c['standard'])['status']=='CONFLICT' and not cat.lookup(c['key'],c['standard'])['profile']['geometry'] and not c['resolution'] for c in cat.conflicts)))
    test(13, '35Б1 explicit user preference and basis', lambda: require(cat.user_state['preferences'][0]['basis']=='USER_CONFIRMED_UZK_328x6_348x9'))
    def preference_conflict():
        copy=deepcopy(records)
        modern_record=next(r for r in copy if r['key']=='35Б1' and r['standard']=='ГОСТ Р 57837-2017')
        bad=deepcopy(modern_record); bad['record_id']='INJECTED'; bad['source']='TEST'; bad['geometry']['b']=175
        require(Catalog(copy+[bad],sources).lookup('35Б1')['status']=='CONFLICT')
    test(14, 'Preference cannot override same-ND source conflict', preference_conflict)
    test(15, 'Explicit normalized standard exact lookup, no fallback', lambda: require(modern['status']=='FOUND' and cat.lookup('35Б1','ГОСТ 26020-1983')['status']=='NOT_FOUND' and cat.lookup('35Б1','ГОСТР57837 - 2017')['status']=='FOUND'))
    test(16, 'No-standard multiple geometry -> AMBIGUOUS', lambda: require(cat.lookup('20Ш1')['status']=='AMBIGUOUS'))
    test(17, 'No-standard valid preference -> FOUND_PREFERRED', lambda: require(preferred['status']=='FOUND_PREFERRED'))
    test(18, 'Unknown profile -> NOT_FOUND', lambda: require(cat.lookup('НЕСУЩЕСТВУЕТ123')['status']=='NOT_FOUND'))
    test(19, 'I / Двутавр normalization preserves significant series', lambda: require(all(cat.lookup(v)['status']=='FOUND_PREFERRED' for v in ('35Б1','I35Б1','I 35Б1','Двутавр 35Б1')) and normalize_profile('35Ш1')!='35Б1'))
    test(20, 'x / х / × / Гнз normalization', lambda: require(all(cat.lookup(v,'ГОСТ 30245-2003')['status']=='FOUND' for v in ('80x40x4','80х40х4','80×40×4','Гнз 80 х 40 х 4'))))
    test(21, '35Б1 ГОСТ Р geometry', lambda: require(core_tuple(modern)==(346,174,6,9)))
    test(22, '35Б1 ГОСТ 26020 geometry', lambda: require(core_tuple(old)==(346,155,6.2,8.5)))
    uzk_modern=calculate_uzk(modern); uzk_old=calculate_uzk(old)
    pairs=lambda result: [(r['L'],r['S']) for r in result['rows']]
    test(23, '35Б1 ГОСТ Р wall 328×6', lambda: require((328,6) in pairs(uzk_modern)))
    test(24, '35Б1 ГОСТ Р flanges 348×9', lambda: require((348,9) in pairs(uzk_modern)))
    test(25, '35Б1 ГОСТ 26020 wall 329×6.2', lambda: require((329,6.2) in pairs(uzk_old)))
    test(26, '35Б1 ГОСТ 26020 flanges 310×8.5', lambda: require((310,8.5) in pairs(uzk_old)))
    test(27, '11.5 -> 5 MHz / 70 degrees', lambda: require(frequency_angle(11.5)=={'frequency_mhz':5,'angle_degrees':70}))
    test(28, '12 -> 2.5 MHz / 65 degrees', lambda: require(frequency_angle(12)=={'frequency_mhz':2.5,'angle_degrees':65}))
    def unsupported():
        for family in ('Швеллер','Уголок равнополочный','Уголок неравнополочный','Квадратный профиль','Прямоугольный профиль','Круглая труба'):
            row=next(r for r in cat.rows if r['family']==family and r['status']!='КОНФЛИКТ')
            require(calculate_uzk(cat.lookup(row['key'],row['standard']))=={'status':NOT_CONFIGURED,'rows':[]})
    test(29, 'Unsupported types have no automatic UZK formula', unsupported)
    test(30, 'Square tube 100×100×6.5', lambda: require(cat.lookup('100×100×6.5','ГОСТ 30245-2003')['profile']['family']=='Квадратный профиль' and geometry(cat.lookup('100×100×6.5','ГОСТ 30245-2003'))['t']==6.5))
    test(31, 'Rectangular tube 80×40×4', lambda: require(cat.lookup('80×40×4','ГОСТ 30245-2003')['profile']['family']=='Прямоугольный профиль'))
    test(32, 'Round pipe D/s from numeric fields', lambda: require(geometry(cat.lookup('159×6','ГОСТ 10704-91'))=={'D':159,'s':6}))
    test(33, 'Missing optional radius tolerated', lambda: require(modern['status']=='FOUND' and geometry(modern).get('R') is None))
    test(34, 'All raw values/provenance preserved in local catalogue', lambda: require(local.raw_records==cat.raw_records and all(r['raw'] and r['evidence'] and r['date'] for r in local.raw_records)))
    manual_build=add_manual('Ручной тест 1','Двутавр','ТЕСТ НД',{'h':100,'b':50,'s':4,'t':6},'fixture only',workbook=fixture)
    test(35, 'Safe manual add with РУЧНОЙ and no automatic UZK', lambda: require(lookup('Ручной тест 1','ТЕСТ НД',fixture)['profile']['status']=='РУЧНОЙ' and lookup('Ручной тест 1','ТЕСТ НД',fixture)['profile']['uzk_rule']==NOT_CONFIGURED))
    set_preference('20Ш1','ГОСТ 26020-83','explicit fixture decision',workbook=fixture)
    fixture_cat=load_catalog(fixture)
    conflict=next(c for c in fixture_cat.conflicts if c['key']=='L125×80×7')
    candidate=next(m for m in conflict['members'] if m['source']=='PTOflow')
    resolve_conflict(conflict['id'],candidate['record_id'],'explicit fixture decision',workbook=fixture)
    state_before=read_json(fixture.parent/STATE)
    rebuild=build(fixture)
    state_after=read_json(fixture.parent/STATE)
    test(36, 'Rebuild preserves manual rows exactly', lambda: require(state_before['manual_rows']==state_after['manual_rows']))
    test(37, 'Rebuild preserves preferences exactly', lambda: require(state_before['preferences']==state_after['preferences'] and lookup('20Ш1',workbook=fixture)['status']=='FOUND_PREFERRED'))
    test(38, 'Verified complete backup before safe replace', lambda: require(rebuild['backup'] is not None and read_json(Path(rebuild['backup'])/STATE)==state_before and (Path(rebuild['backup'])/fixture.name).is_file()))
    def failures():
        files=[fixture.name,STATE,CATALOG,COMMIT]
        before_failure={f:sha(fixture.parent/f) for f in files}
        for point in ('before_commit','after_first_replace'):
            require(raises(lambda: build(fixture,fail_at=point)))
            require(before_failure=={f:sha(fixture.parent/f) for f in files})
            require(load_catalog(fixture).lookup('35Б1')['status']=='FOUND_PREFERRED')
    test(39, 'Safe write failures preserve previous complete generation', failures)
    test(40, 'Reopen final fixture workbook all cells/types', lambda: require(verify_workbook(fixture,workbook_payload(load_catalog(fixture)))['reopened']))
    test(41, 'DEV-only destination guard, validation isolation', lambda: require(guard_target(DEFAULT_WORKBOOK)==DEFAULT_WORKBOOK.resolve() and raises(lambda: guard_target(ROOT/'Профили.xlsx')) and raises(lambda: guard_target(ROOT.parent.parent/'Профили.xlsx'))))
    test(42, 'No production reads/writes observed; paths denied', lambda: require(not production_events and raises(lambda: build(ROOT.parent.parent/'Профили.xlsx'))))
    def smoke():
        for module in ['passport_app.config','passport_app.registry','passport_app.service','passport_app.vik',
                       'passport_app.catalog','passport_app.catalog_gui','stage2','stage3']:
            importlib.import_module(module)
        checkpoint=read_json(ROOT/'DEV/DEV_REGISTRY_CHECKPOINT.json')
        require(checkpoint['sha256']==sha(ROOT/'DEV/ВЗМК_DEV.xlsx'))
        require(all(sha(p)==h for p,h in protected.items()))
    test(43, 'Stage 1–3 imports, readable DEV checkpoint, protected hashes unchanged', smoke)
    test(44, 'Explicit conflict resolution preserved, no mutation of raw records', lambda: require(state_before['resolutions']==state_after['resolutions'] and lookup('L125×80×7','ГОСТ 8510-86',fixture)['status']=='FOUND' and geometry(lookup('L125×80×7','ГОСТ 8510-86',fixture))['b2']==80))
    def stale_resolution():
        changed=deepcopy(records); changed[0]['raw']['tamper_fixture']='changed'
        conflict_row=next(r for r in changed if r['key']=='L125×80×7')
        conflict_row['raw']['tamper_fixture']='changed'
        require(Catalog(changed,sources,state_after).lookup('L125×80×7','ГОСТ 8510-86')['status']=='CONFLICT')
    test(45, 'Stale resolution fingerprint does not resolve changed evidence', stale_resolution)
    def normalized_collision():
        r=deepcopy(next(r for r in records if r['key']=='35Б1' and r['standard']=='ГОСТ Р 57837-2017'))
        r['family']='Швеллер'; r['record_id']='COLLISION'; r['source']='TEST'
        require(Catalog(records+[r],sources).lookup('35Б1','ГОСТ Р 57837-2017')['status']=='AMBIGUOUS')
    test(46, 'Different real families with normalized collision -> AMBIGUOUS', normalized_collision)
    def equivalent():
        require(cat.lookup('25Б1')['status']=='FOUND' and set(cat.lookup('25Б1')['equivalent_standards'])=={'СТО АСЧМ 20-93','ГОСТ Р 57837-2017'})
        require(len(cat.lookup('25Б1')['profile']['provenance'])==2)
    test(47, 'Equivalent geometry across standards -> FOUND with all provenance', equivalent)
    def radius_conflict():
        r=deepcopy(next(r for r in records if r['key']=='35Б1' and r['standard']=='СТО АСЧМ 20-93'))
        r['record_id']='RADIUS'; r['source']='TEST'; r['geometry']['R']=15
        c=Catalog(records+[r],sources)
        require(c.lookup('35Б1','СТО АСЧМ 20-93')['status']=='CONFLICT' and any(x['fields']==['R'] for x in c.conflicts))
    test(48, 'Radius-only conflict is explicit, never silently chosen', radius_conflict)
    def same_thickness():
        r=deepcopy(modern); r['profile']['geometry']['t']=6
        out=calculate_uzk(r)
        require(len(out['rows'])==1 and pairs(out)==[(682,6)])
        require(raises(lambda: calculate_uzk({'status':'CONFLICT'})))
        require(raises(lambda: frequency_angle(0)))
    test(49, 'Same S sums within one item; conflict/invalid thickness rejected', same_thickness)
    def offline_guard():
        with no_network() as probes:
            require(raises(lambda: socket.create_connection(('example.invalid',443))))
        require(len(probes)==1)
    test(50, 'Negative no-network guard probe blocked before network IO', offline_guard)
    test(51, 'Excluded/unproved standards/category-only rows absent', lambda: require(not any(any(v in r['standard'] for v in ('8639','8645','ТУ','ВНЕ ГОСТ')) for r in records) and cat.lookup('90×50×6.5','ГОСТ 30245-2003')['status']=='NOT_FOUND'))
    test(52, 'Rounded pipe designation does not overwrite numeric identity', lambda: require(geometry(cat.lookup('10.2×1','ГОСТ 10704-91'))['D']==10.2 and geometry(cat.lookup('10×1','ГОСТ 10704-91'))['D']==10))
    def manual_confirmation():
        record=load_catalog(fixture).user_state['manual_rows'][0]
        confirm_manual_uzk(record['record_id'],'explicit fixture rule confirmation',workbook=fixture)
        require(lookup('Ручной тест 1','ТЕСТ НД',fixture)['profile']['uzk_rule']==RULE)
    test(53, 'Manual I-beam UZK only after separate explicit confirmation', manual_confirmation)
    require(all(t['status']=='PASS' for t in tests), 'Targeted tests failed; DEV build is gated')
    final_build=build(previews=OUT/'final_previews')
    final_cat=load_catalog()
    def ui_smoke():
        from profiles_stage4.gui import ProfileWindow
        app=ProfileWindow()
        try:
            app.update()
            for text,nd,status in [('35Б1','','FOUND_PREFERRED'),('35Б1','ГОСТ 26020-83','FOUND'),
                                   ('20Ш1','','AMBIGUOUS'),('L125×80×7','ГОСТ 8510-86','CONFLICT'),('UNKNOWN','','NOT_FOUND')]:
                app.profile.delete(0,'end'); app.profile.insert(0,text)
                app.standard.delete(0,'end'); app.standard.insert(0,nd)
                app.search(); app.update()
                require(app.status.cget('text')==status)
                require(status in app.output.get('1.0','end'))
            with patch('profiles_stage4.gui.os.startfile') as start:
                app.open_workbook(); start.assert_called_once_with(str(DEFAULT_WORKBOOK))
        finally:
            app.destroy()
    test(54, 'Real Tk widgets: five lookup statuses, UZK/provenance output and DEV open dispatch', ui_smoke)
    def tampering():
        path=fixture.parent/STATE
        saved=path.read_bytes()
        workbook_hash=sha(fixture)
        try:
            data=read_json(path); data['preferences']=[]; dump(path,data)
            require(raises(lambda: load_catalog(fixture)))
            require(raises(lambda: build(fixture)))
            require(sha(fixture)==workbook_hash)
        finally:
            path.write_bytes(saved)
        require(load_catalog(fixture).lookup('35Б1')['status']=='FOUND_PREFERRED')
    test(55, 'Local user state tampering fails closed without rebuild overwrite', tampering)
    test(56, 'Normalization preserves series digits and opaque word boundaries', lambda: require(
        normalize_profile('35Б01')!='35Б1' and normalize_profile('Модель A B')!=normalize_profile('Модель AB')
        and normalize_profile('I 35Б01')=='35Б01' and normalize_profile('Серия X12')=='СЕРИЯ X12'))
    golden={nd: {'lookup':final_cat.lookup('35Б1',nd), 'uzk':calculate_uzk(final_cat.lookup('35Б1',nd))}
            for nd in ('ГОСТ Р 57837-2017','ГОСТ 26020-83')}
    golden['preferred']=final_cat.lookup('35Б1')
    dump(OUT/'golden_35Б1.json',golden)
    dump(OUT/'build_result.json',final_build)
    dump(OUT/'conflicts.json',final_cat.conflicts)
    dump(OUT/'ambiguous_keys.json',[k for k in final_cat.by_key if final_cat.lookup(k)['status']=='AMBIGUOUS'])
    dump(OUT/'source_ingestion.json',{'source_raw':len(records),'A':len(a),'B':len(b),'rounding_records':len(rounded),
        'rounding_fields':sum(len(r['display_rounding']) for r in rounded),'sources':sources,
        'excluded':'All category/detail rows, 8639/8645/TU/special/outside-GOST; only nine accepted group datasets imported'})
except Exception as exc:
    tests.append({'id':'FATAL','name':'Validation runner','status':'FAIL','error':str(exc)})
    print('FATAL',str(exc),flush=True)
finally:
    after=hashes()
    dump(OUT/'product_after.json',after)
    preserved={p:sha(p) for p in protected}
    dump(OUT/'protected_after.json',preserved)
    report={'tests':tests,'mandatory_ids':list(range(1,44)),'passed':sum(t['status']=='PASS' for t in tests),
        'failed':sum(t['status']=='FAIL' for t in tests),'product_before_equals_after':before==after,
        'protected_unchanged':protected==preserved,'protected_count':len(protected),
        'network_events':network_events,'production_events':production_events,'tested_product_hashes':before}
    dump(OUT/'tests.json',report)
    print(json.dumps({k:v for k,v in report.items() if k not in ('tests','tested_product_hashes')},ensure_ascii=False),flush=True)
sys.exit(0 if report['failed']==0 and report['product_before_equals_after'] and report['protected_unchanged'] else 1)
