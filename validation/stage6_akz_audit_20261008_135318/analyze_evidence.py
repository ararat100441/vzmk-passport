import sys
sys.dont_write_bytecode=True
from audit_readonly import OUT, ROOT, PROJECT, PASSPORT, TEMPLATES, save
import json, re, unicodedata
from pathlib import Path
from collections import Counter, defaultdict
from decimal import Decimal

books={}
for p in OUT.glob('xlsx_*.json'):
    obj=json.loads(p.read_text(encoding='utf-8')); books[obj['id']]=obj

def key(v): return re.sub(' +',' ',unicodedata.normalize('NFC',str(v)).strip()).casefold()
def blank(v): return v is None or (isinstance(v,str) and not v.strip())
def row_eq(a,b):
    return all(abs(Decimal(str(x))-Decimal(str(y)))<=Decimal('0.0000001') if isinstance(x,(int,float)) and isinstance(y,(int,float)) else x==y for x,y in zip(a,b))
def matrix(s): return {c['cell']:c.get('cached') if c['type']=='f' else c['value'] for c in s['cells']}
def val_rows(s,start):
    m=matrix(s); out=[]; total=None
    for r in range(start,1000):
        v=m.get('A'+str(r))
        if blank(v): continue
        if key(v)=='итого':
            total={c:m.get(c+str(r)) for c in 'CEG'} | {'row':r}; break
        out.append({'row':r,'mark':v,'values':[m.get(c+str(r)) for c in 'ABCDEFG']})
    return out,total

cat=books[33]; m=matrix(cat['sheets'][0]); rows=[]
for r in range(2,197):
    vals=[m.get(c+str(r)) for c in 'ABCD']
    if any(not blank(v) for v in vals): rows.append({'row':r,'mark':vals[0],'profile':vals[1],'paint':vals[2],'akz':vals[3]})
counts={f:sum(not blank(x[f]) for x in rows) for f in ['mark','profile','paint','akz']}
groups=defaultdict(list)
for x in rows:
    if not blank(x['mark']):groups[key(x['mark'])].append(x)
stats={'path':cat['path'],'sha256':cat['sha256'],'sheet':'Марки','range':'A1:D196','rows':len(rows),'filled':counts,
       'blank':{k:len(rows)-v for k,v in counts.items()},'unique_marks':len(groups),
       'duplicates':{k:v for k,v in groups.items() if len(v)>1},
       'values':{k:dict(Counter(str(x[k]) for x in rows if not blank(x[k]))) for k in ['profile','paint','akz']},
       'incomplete_assignments':[x for x in rows if blank(x['paint'])!=blank(x['akz'])],
       'invalid_akz':[x for x in rows if not blank(x['akz']) and (isinstance(x['akz'],bool) or not isinstance(x['akz'],(int,float)) or x['akz']<0 or int(x['akz'])!=x['akz'])],
       'check_profile':[x for x in rows if key(x['profile'])=='проверить']}
save('catalog_assessment.json',stats)
print('CATALOG',json.dumps(stats,ensure_ascii=False))

reg=json.loads((OUT/'registry_project.json').read_text(encoding='utf-8'))
reg_rows={x['values']['№ паспорта']:x for x in reg['full_project_rows']}
history=[]
acts=[b for b in books.values() if Path(b['path']).name.startswith('Акт')] if False else [b for b in books.values() if re.search(r'\\Акт',b['path'])]
apps=[b for b in books.values() if ' - Приложение к АКЗ' in b['path']]
app1=[b for b in books.values() if ' - Приложение № 1' in b['path']]
for a in sorted(apps,key=lambda b:int(re.search(r'(\d+)_2026',b['path'])[1])):
    num=int(re.search(r'(\d+)_2026',a['path'])[1]); canon=f'{num}/2026'
    act=next(x for x in acts if str(Path(x['path']).parent)==str(Path(a['path']).parent))
    primary=next(x for x in app1 if str(Path(x['path']).parent)==str(Path(a['path']).parent))
    am=matrix(act['sheets'][0]); ar,at=val_rows(a['sheets'][0],4); vr,vt=val_rows(primary['sheets'][0],5)
    vd={key(x['mark']):x for x in vr}
    missing=[x['mark'] for x in ar if key(x['mark']) not in vd]
    diff=[{'akz':x,'vom':[v for v in vr if key(v['mark'])==key(x['mark'])]} for x in ar if key(x['mark']) in vd and not any(row_eq(x['values'],v['values']) for v in vr if key(v['mark'])==key(x['mark']))]
    no_akz=[x for x in vr if key(x['mark']) not in {key(y['mark']) for y in ar}]
    products=[]
    for x in ar:
        v=x['values']
        for u,t in [(3,4),(5,6)]:
            if all(isinstance(v[k],(int,float)) for k in [2,u,t]):
                delta=Decimal(str(v[2]))*Decimal(str(v[u]))-Decimal(str(v[t]))
                if abs(delta)>Decimal('0.00001'): products.append({'mark':v[0],'field':'mass' if u==3 else 'area','unit_times_q_minus_total':str(delta)})
    sums={c:float(sum(Decimal(str(x['values'][j])) for x in ar)) if all(isinstance(x['values'][j],(int,float)) for x in ar) else None for c,j in [('C',2),('E',4),('G',6)]}
    details={'passport':canon,'akz_id':a['id'],'act_id':act['id'],'appendix1_id':primary['id'],
        'header':matrix(a['sheets'][0]).get('A1'),'act_number':am.get('E4'),'act_date':am.get('H4'),
        'paint':am.get('E18'),'colour':am.get('E21'),'dry_actual':am.get('E28'),'required':am.get('E25'),
        'climate':{c:am.get(c) for c in ['E31','E32','E33','E34']},
        'rows':len(ar),'quantity':sums['C'],'mass':sums['E'],'area':sums['G'],'totals':at,
        'appendix1_rows':len(vr),'appendix1_totals':vt,'not_in_appendix1':missing,'row_differences':diff,
        'blank_areas':[x['mark'] for x in ar if blank(x['values'][5]) or blank(x['values'][6])],
        'appendix1_excluded':no_akz,'recomputed_product_differences':products,
        'registry_akz_date':reg_rows.get(canon,{}).get('values',{}).get('Дата Акз'),
        'catalog_missing_marks':[x['mark'] for x in ar if key(x['mark']) not in groups],
        'marks':[x['mark'] for x in ar]}
    history.append(details)
save('history_comparison.json',history)
for h in history: print('HISTORY',json.dumps({k:v for k,v in h.items() if k not in {'marks','appendix1_excluded','recomputed_product_differences','row_differences'}},ensure_ascii=False))

variants=[]
for i in [32,39,40,41,35,36,37,38]:
    b=books[i]
    variants.append({'id':i,'path':b['path'],'sha256':b['sha256'],'sheets':[{k:s[k] for k in ['name','state','dimension','print_area','merged','formula_count','conditional_formats','data_validations']} | {'formulas':[c for c in s['cells'] if c['type']=='f']} for s in b['sheets']]})
save('variant_structure.json',variants)

# All changing historical act cells, excluding formula cache changes.
actmaps=[{c['cell']:c['value'] for c in b['sheets'][0]['cells']} for b in acts]
vary=[]
for c in sorted(set().union(*(set(m) for m in actmaps))):
    vs={str(m.get(c)) for m in actmaps}
    if len(vs)>1:vary.append({'cell':c,'values':sorted(vs)})
save('historical_dynamic_cells.json',vary)
print('VARY',json.dumps(vary,ensure_ascii=False))

native=json.loads((OUT/'native_excel.json').read_text(encoding='utf-8'))
nat2=json.loads((OUT/'native_excel_extra.json').read_text(encoding='utf-8'))
for r in nat2['files']: print('EXTRA_PAGES',Path(r['path']).name,[(s['name'],s.get('native_print_pages'),s.get('pagination_error')) for s in r['sheets']])

voms=[]
visibility={x['path']:set(x['visible_rows_B']) for x in json.loads((OUT/'native_visibility.json').read_text(encoding='utf-8'))}
for f in native['files']:
    if not f['path'].endswith('.xls'):continue
    s=f['sheets'][0]; m={c['cell']:c['value'] for c in s['cells']}; rlist=[]
    for r in range(4,1000):
        mk=m.get('B'+str(r)); q=m.get('D'+str(r))
        if blank(mk) or key(mk)=='итого' or not isinstance(q,(int,float)) or q<=0:continue
        rlist.append({'row':r,'visible':r in visibility[f['path']], 'mark':mk,'values':[m.get(c+str(r)) for c in ['B','C','D','H','I','J','K']]})
    voms.append({'path':f['path'],'sheet':s['name'],'zero_quantity_rows_omitted':sum(isinstance(m.get('D'+str(r)),(int,float)) and m.get('D'+str(r))==0 for r in range(4,1000)), 'rows':rlist})
save('vom_rows.json',voms)
matches=[]
for h in history:
    a=books[h['akz_id']]; ar,_=val_rows(a['sheets'][0],4)
    candidates=[]
    for v in voms:
        vp=Path(v['path'])
        if not vp.is_relative_to(PASSPORT):continue
        rel=vp.relative_to(PASSPORT)
        vn=1 if len(rel.parts)==1 else int(re.match(r'^(\d+)',rel.parts[0])[1])
        if vn==int(h['passport'].split('/')[0]):candidates.append(v)
    for v in candidates:
        vd={key(x['mark']):x for x in v['rows']}
        matches.append({'passport':h['passport'],'vom':v['path'],
            'unmatched':[x['mark'] for x in ar if key(x['mark']) not in vd],
            'differences':[{'akz':x,'vom':[r for r in v['rows'] if key(r['mark'])==key(x['mark'])]} for x in ar if key(x['mark']) in vd and not any(row_eq(x['values'],r['values']) for r in v['rows'] if key(r['mark'])==key(x['mark']))]})
save('vom_matches.json',matches)
print('VOM_MATCH',json.dumps([{'passport':x['passport'],'vom':x['vom'],'missing_count':len(x['unmatched']),'diff_count':len(x['differences'])} for x in matches],ensure_ascii=False))

aggregates=[]
def agg(rows):
    d={}
    for r in rows:
        k=key(r['mark']); v=r['values']
        if not isinstance(v[2],(int,float)) or not isinstance(v[4],(int,float)):continue
        if k not in d:d[k]={'quantity':Decimal(0),'mass':Decimal(0),'units':set()}
        d[k]['quantity']+=Decimal(str(v[2]));d[k]['mass']+=Decimal(str(v[4])); d[k]['units'].add(str(v[3]))
    return d
for h in history:
    a=books[h['akz_id']]; ar,_=val_rows(a['sheets'][0],4); ag=agg(ar)
    for match in [v for v in matches if v['passport']==h['passport']]:
        vom=next(v for v in voms if v['path']==match['vom']); vg=agg([r for r in vom['rows'] if r['visible']])
        conflicts=[]
        for mk,x in ag.items():
            y=vg.get(mk)
            if y is None or x['quantity']!=y['quantity'] or abs(x['mass']-y['mass'])>Decimal('0.0000001'):
                conflicts.append({'mark':mk,'akz':x,'vom':y})
        aggregates.append({'passport':h['passport'],'vom':match['vom'],'akz_unique':len(ag),'vom_unique':len(vg),'unparsed_numeric_rows':[r for r in vom['rows'] if not isinstance(r['values'][2],(int,float)) or not isinstance(r['values'][4],(int,float))], 'conflicts':conflicts,'vom_excluded':[mk for mk in vg if mk not in ag]})
save('vom_aggregate_matches.json',aggregates)
print('AGGREGATE_MATCH',json.dumps([{'passport':x['passport'],'vom':x['vom'],'conflicts':len(x['conflicts']),'vom_excluded':len(x['vom_excluded'])} for x in aggregates],ensure_ascii=False))
visible_summary=[]
for v in voms:
    rs=[r for r in v['rows'] if r['visible']]
    valid=[r for r in rs if all(isinstance(r['values'][j],(int,float)) for j in [2,4])]
    rec={'path':v['path'],'visible_data_rows':len(rs),'unique_visible':len(set(key(r['mark']) for r in rs)),
         'quantity':str(sum(Decimal(str(r['values'][2])) for r in valid)),
         'mass':str(sum(Decimal(str(r['values'][4])) for r in valid)),
         'exact_historical_mark_sets':[h['passport'] for h in history if set(map(key,h['marks']))==set(key(r['mark']) for r in rs)]}
    visible_summary.append(rec)
save('visible_vom_summary.json',visible_summary)
print('VISIBLE_SUMMARY',json.dumps(visible_summary,ensure_ascii=False))
