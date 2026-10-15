import sys,json,traceback
from pathlib import Path
OUT=Path(__file__).resolve().parent;ROOT=OUT.parents[1];sys.path.insert(0,str(ROOT))
from stage9_app.ownership import owned_office
p=OUT/'OLD_EMPTY_GUARD.json'
try:
    with owned_office(p):pass
except Exception as exc:
    result={'status':'REPRODUCED','type':type(exc).__name__,'traceback':traceback.format_exc(),'source_only':True}
    assert result['type']=='FileExistsError'
else:raise AssertionError('Expected original duplicate write')
(OUT/'OLD_EMPTY_GUARD_REPRO.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
print(result['type'])
