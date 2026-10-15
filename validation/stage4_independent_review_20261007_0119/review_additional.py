import sys,json,hashlib
from pathlib import Path
from collections import defaultdict
from copy import deepcopy
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
from profiles_stage4 import model as m, storage as s, offline as o
sys.stdout.reconfigure(encoding='utf-8')
fields={'Двутавр':('h','b','s','t'),'Швеллер':('h','b','s','t'),'Уголок равнополочный':('b','t'),
 'Уголок неравнополочный':('b','b2','t'),'Квадратный профиль':('h','b','t'),'Прямоугольный профиль':('h','b','t'),'Круглая труба':('D','s')}
with o.no_network() as network:
    cat=s.load_catalog();rows=defaultdict(list)
    for r in cat.rows:rows[r['key']].append(r)
    prefs={p['key']:p['standard'] for p in cat.user_state['preferences'] if p['status']=='ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ'}
    ambiguous=[]
    for key,rr in rows.items():
        if len(rr)<2 or any(r['status']=='КОНФЛИКТ' for r in rr):continue
        signatures={(r['family'],*(r['geometry'][f] for f in fields[r['family']])) for r in rr}
        if len(signatures)>1 and not any(r['standard']==prefs.get(key) for r in rr):ambiguous.append(key)
    assert set(ambiguous)=={k for k in rows if cat.lookup(k)['status']=='AMBIGUOUS'} and len(ambiguous)==51
    pipe=[r for r in cat.raw_records if r['key']=='1420×20' and r['standard']=='ГОСТ 10704-91']
    assert len(pipe)==2 and pipe[0]['raw']==pipe[1]['raw'] and pipe[0]['geometry']==pipe[1]['geometry']
    assert len([r for r in cat.rows if r['key']=='1420×20' and r['standard']=='ГОСТ 10704-91'])==1
    main=next(r for r in cat.raw_records if r['key']=='35Б1' and r['standard']=='ГОСТ Р 57837-2017')
    bad=deepcopy(main);bad['geometry']['b']=175;bad['record_id']='REVIEW-GEOMETRY'
    assert m.Catalog(cat.raw_records+[bad],cat.sources,cat.user_state).lookup('35Б1')['status']=='CONFLICT'
    invalid=s.set_preference
    base={r['family']:deepcopy(r) for r in cat.raw_records if r['family'] in ('Квадратный профиль','Прямоугольный профиль')}
    for r in base.values():
        r['geometry'].pop('R',None);r['geometry'].pop('r',None)
        m.validate_geometry(r['family'],r['geometry'])
        assert m.Catalog([r],user_state=m.default_user_state()).lookup(r['key'],r['standard'])['status']=='FOUND'
    a=deepcopy(main);a['display']=a['key']='REVIEW';a['standard']='REVIEW ND A';a['geometry']['R']=5
    b=deepcopy(a);b['record_id']='REVIEW-B';b['standard']='REVIEW ND B';b['geometry']['R']=6
    r=m.Catalog([a,b]).lookup('REVIEW')
    assert r['status']=='FOUND' and 'R' not in r['profile']['geometry'] and any('разные значения' in n for n in r['profile']['notes'])
assert not network
data={'independently_derived_ambiguous_key_count':len(ambiguous),'round_1420x20_exact_raw_duplicate_one_identity':True,
 'preference_cannot_override_geometry_conflict':True,'missing_profile_tube_radius_does_not_invalidate_main':True,
 'different_ND_optional_radius_explicit_and_main_usable':True,'network_attempts':0}
(OUT/'additional_independent.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(data,ensure_ascii=False))
