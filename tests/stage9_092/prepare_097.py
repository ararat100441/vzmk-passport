from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
for rel in ('stage9_app/__init__.py','tests/stage9_092/common.py','tests/stage9_092/e2e.py','tests/stage9_092/inventory.py','tests/stage9_092/major_restore_repro.py'):
    p=ROOT/rel;p.write_text(p.read_text(encoding='utf8').replace('0.9.6','0.9.7').replace('stage9_dist_096','stage9_dist_097').replace('FIXED_096','FIXED_097'),encoding='utf8')
p=ROOT/'tests/stage9_092/inventory.py';text=p.read_text(encoding='utf8').replace("ROOT/'stage9_096.spec']","ROOT/'stage9_096.spec',ROOT/'stage9_097.spec']");p.write_text(text,encoding='utf8')
for old,new in (('stage9_096.spec','stage9_097.spec'),('tests/stage9_092/build_096.py','tests/stage9_092/build_097.py')):
    p=ROOT/new;assert not p.exists();p.write_text((ROOT/old).read_text(encoding='utf8').replace('0.9.6','0.9.7').replace('_096','_097'),encoding='utf8')
for name in ('LIFECYCLE_FINAL.json','OWNERSHIP_STOP_FINAL.json'):
    p=ROOT/'validation/stage9_092_20261012_01'/name
    if p.exists():p.rename(p.with_name(p.stem+'_PRE_097.json'))
