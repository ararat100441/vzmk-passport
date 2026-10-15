from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
for rel in ('stage9_app/__init__.py','tests/stage9_092/common.py','tests/stage9_092/e2e.py','tests/stage9_092/inventory.py','tests/stage9_092/major_restore_repro.py'):
    p=ROOT/rel;p.write_text(p.read_text(encoding='utf8').replace('0.9.5','0.9.6').replace('stage9_dist_095','stage9_dist_096').replace('FIXED_095','FIXED_096'),encoding='utf8')
p=ROOT/'tests/stage9_092/inventory.py';text=p.read_text(encoding='utf8').replace("ROOT/'stage9_095.spec']","ROOT/'stage9_095.spec',ROOT/'stage9_096.spec']");p.write_text(text,encoding='utf8')
for old,new in (('stage9_095.spec','stage9_096.spec'),('tests/stage9_092/build_095.py','tests/stage9_092/build_096.py')):
    p=ROOT/new;assert not p.exists();p.write_text((ROOT/old).read_text(encoding='utf8').replace('0.9.5','0.9.6').replace('_095','_096'),encoding='utf8')
