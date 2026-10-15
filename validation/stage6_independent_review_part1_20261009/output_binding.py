from review import *
from akz_stage6 import core as c
old=load(I/'preview.json');m=load(I/'EVIDENCE_MANIFEST.json');out=Path(m['canonical_output'])
fixture=Path(load(I/'test_assignments.json')['fixture'])
p=c.preview('Гараж Казахстан','9/2026',old['payload']['vom']['path'],fixture)
assert p.seal==old['seal']==c.digest(old['payload'])
assert load(out/'PAIR.json')['act_id']==p.seal[:16]=='e9fc12e4433feafd'
with c.native_excel() as app:
    c.validate_saved_pair(app,out/'ТЕСТ АКЗ №0.xlsx',out/'ТЕСТ Приложение к АКЗ №0.xlsx',p)
save('actual_output_product_binding.json',{'current_product_seal':p.seal,'implementation_seal':old['seal'],
    'seal_exact_match':True,'native_saved_pair_validation':'PASS','outputs':{str(x):sha(x) for x in out.glob('*.xlsx')},
    'sources':p.payload['sources'],'tested_product_current':{path:sha(path) for path in load(I/'tested_product_hashes.json')['CURRENT']}})
print('EXACT OUTPUT CURRENT PRODUCT BINDING PASS',p.seal)
