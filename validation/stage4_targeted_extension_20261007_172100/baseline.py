"""Read-only acceptance preflight; snapshot only into this new validation folder."""
from pathlib import Path
import json, shutil, sys
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT)); sys.stdout.reconfigure(encoding='utf-8')
from profiles_stage4.offline import sha,read_json,no_network,load_sources
from profiles_stage4 import storage as s
OLD=ROOT/'validation/stage4_excel_table_fix_20261007_121500'
manifest=read_json(OLD/'REVIEW_MANIFEST.json')
assert sha(s.DEFAULT_WORKBOOK)=='39f14b040d3a97365f9e44f5fdc01c87e1f2ce068d5d4f26458a91602759ca91'
for p,h in manifest['product_hashes'].items(): assert sha(ROOT/p)==h,p
for p,h in manifest['protected_hashes'].items(): assert sha(p)==h,p
for p,h in manifest['DEV_artifacts'].items(): assert sha(p)==h,p
protected=dict(manifest['protected_hashes'])
for p in OLD.rglob('*'):
    if p.is_file() and '__pycache__' not in p.parts: protected[str(p)]=sha(p)
for folder in [ROOT/'passport_app',ROOT/'DEV',ROOT/'validation']:
    if folder==ROOT/'validation':
        paths=[p for d in folder.glob('*stage5*') for p in d.rglob('*') if p.is_file()]
    else: paths=[p for p in folder.rglob('*') if p.is_file() and '__pycache__' not in p.parts]
    for p in paths:
        if p.parent==s.DEFAULT_WORKBOOK.parent or p.is_relative_to(s.DEFAULT_WORKBOOK.parent/'backups'): continue
        protected[str(p)]=sha(p)
project=ROOT.parent/'Гараж казахстан'
for name in ['Справочник марок.xlsx','УЗК_DEV.xlsm']:
    p=project/name
    if p.is_file(): protected[str(p)]=sha(p)
for name in ['profiles_stage4/model.py','profiles_stage4/storage.py','profiles_stage4/workbook.mjs','STAGE4_README.md']:
    protected.pop(str(ROOT/name),None)
with no_network() as attempts:
    cat=s.load_catalog(); records,sources=load_sources()
assert not attempts and len(records)==3037
assert cat.lookup('16П','ГОСТ 8240-97')['profile']['geometry']=={'h':160,'b':64,'s':5,'t':8.4,'R':8.5,'r':5}
assert cat.lookup('[16П')['status']=='NOT_FOUND'
assert cat.lookup('120×120×3','ГОСТ 30245-2003')['status']=='NOT_FOUND'
assert not any(r['standard']=='ГОСТ 8278-83' for r in cat.rows)
assert not (OUT/'baseline_generation').exists()
(OUT/'baseline_generation').mkdir()
for name in ['Профили.xlsx',s.STATE,s.CATALOG,s.COMMIT]: shutil.copy2(s.DEFAULT_WORKBOOK.parent/name,OUT/'baseline_generation'/name)
s.dump(OUT/'generation_before.json',manifest['DEV_artifacts'])
s.dump(OUT/'protected_before.json',protected)
s.dump(OUT/'baseline_product.json',{p:h for p,h in manifest['product_hashes'].items() if not p.startswith('validation')})
s.dump(OUT/'baseline_result.json',{'status':'PASS','accepted_manifest':str(OLD/'REVIEW_MANIFEST.json'),'accepted_manifest_sha256':sha(OLD/'REVIEW_MANIFEST.json'),'counts':cat.counts(),'state_revision':cat.user_state['revision'],'protected_files':len(protected),'source_raw':len(records),'network_attempts':0,'channel_16П':cat.lookup('16П','ГОСТ 8240-97')})
print(json.dumps({'baseline':'PASS','counts':cat.counts(),'protected':len(protected)},ensure_ascii=False))
