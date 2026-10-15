import sys,json,hashlib
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
from uzk_stage5.inputs import visible_vom,registry_rows,key
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
paths=[ROOT.parent/'Гараж казахстан/Паспорт/ВОМ отгрузочный счет 183  (3).xls',ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls']
registry,evidence=registry_rows()
result=[]
for path in paths:
    before=sha(path);vom=visible_vom(path)
    matched=[{'vom':r,'registry':registry[(key('Гараж Казахстан'),key(r['mark']))]} for r in vom['rows'] if (key('Гараж Казахстан'),key(r['mark'])) in registry]
    result.append({'sha_before':before,'sha_after':sha(path),'vom':vom,'quantity_total':sum(r['quantity'] for r in vom['rows']),'uzk_required_matches':matched,'uzk_status':'UZK_REQUIRED' if matched else 'NO_UZK_REQUIRED'})
(OUT/'live_vom_uzk_readonly.json').write_text(json.dumps({'registry_evidence':evidence,'shipments':result},ensure_ascii=False,indent=2,default=str),encoding='utf8')
print(json.dumps([{'path':r['vom']['path'],'visible_rows':r['vom']['visible_count'],'quantity':r['quantity_total'],'mass':r['vom']['mass_kg'],'uzk_status':r['uzk_status'],'uzk_marks':[m['vom']['mark'] for m in r['uzk_required_matches']]} for r in result],ensure_ascii=False))
