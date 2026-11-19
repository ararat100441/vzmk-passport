import sys,json,os,traceback
from decimal import Decimal
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save
from stage9_app import core
from stage8_closure import storage as S
OUT=ROOT/'validation/stage9_20261012_01'
work=Path(S.read(OUT/'WORKSPACE.json')['path'])
vom=ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
try:
    plan=core.check(work,'Гараж Казахстан','9/2026',vom,(0,))
    assert plan['data']['visible_rows']==104 and plan['data']['hidden_excluded']==397,plan['data']
    assert plan['data']['quantity']=='288' and Decimal(plan['data']['mass_kg'])==Decimal('15822.14')
    result=core.generate(plan,True)
    m=core.verify_manifest(result)
    save(OUT/'E2E.json',{'status':'PASS','manifest':str(result),'data':plan['data'],'quality':m['quality_validation'],'native':m['native_quality']})
    print(result)
except Exception as exc:
    save(OUT/('E2E_FAIL_'+__import__('uuid').uuid4().hex[:6]+'.json'),{'status':'FAIL','error':str(exc),'traceback':traceback.format_exc()})
    raise
