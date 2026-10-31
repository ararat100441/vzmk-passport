import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app import core
from stage9_app.workspace import bootstrap,save,PROJECT
from stage9_app.ownership import owned_office
import os,traceback
OUT=ROOT/'validation/stage9_20261012_01'
from uuid import uuid4
attempt=uuid4().hex[:6]
with owned_office(OUT/f'NEW_PASSPORT_OFFICE_{attempt}.json'):
    work=bootstrap(ROOT,Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/workspaces')
    _,next_number=core.passports(work,PROJECT)
    vom=ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
    try:
        plan=core.check(work,PROJECT,next_number,vom,(0,),True)
        output=core.generate(plan,True)
        m=core.verify_manifest(output)
        assert m['passport']==next_number and m['state']=='COMPLETE_TEST'
        save(OUT/'NEW_PASSPORT.json',{'status':'PASS','workspace':str(work),'number':next_number,'manifest':str(output)})
    except Exception as exc:
        save(OUT/f'NEW_PASSPORT_FAIL_{attempt}.json',{'status':'FAIL','error':str(exc),'traceback':traceback.format_exc(),'workspace':str(work)})
        raise
