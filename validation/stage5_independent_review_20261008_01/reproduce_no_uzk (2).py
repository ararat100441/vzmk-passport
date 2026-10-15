"""Repeatable read-only reproduction against retained isolated fixture bytes."""
import json,sys
from pathlib import Path
from contextlib import ExitStack
from unittest.mock import patch
RUN=Path(__file__).resolve().parent;ROOT=RUN.parents[1];sys.path.insert(0,str(ROOT))
import uzk_stage5.inputs as I
import uzk_stage5.core as C
root=RUN/'scenarios/no_uzk_True'
paths={'ROOT':root,'REGISTRY':root/'DEV/ВЗМК_DEV.xlsx','UZK':root/'DEV/УЗК_DEV.xlsm','PROFILES':root/'DEV/Справочники/Профили.xlsx','CHECKPOINT':root/'DEV/DEV_UZK_CHECKPOINT.json'}
before={str(p):I.sha(p) for p in root.rglob('*') if p.is_file()}
original_passports=I.passports;original_registry=I.registry_rows
with ExitStack() as stack:
    for module in (I,C):
        for name,value in paths.items():stack.enter_context(patch.object(module,name,value))
    stack.enter_context(patch.object(I,'PINS',{paths[k]:I.sha(paths[k]) for k in ('REGISTRY','UZK','PROFILES')}))
    stack.enter_context(patch.object(I,'passports',side_effect=lambda:original_passports(paths['REGISTRY'])))
    stack.enter_context(patch.object(C,'registry_rows',side_effect=lambda:original_registry(paths['UZK'])))
    stock,records=C.registry_rows();assert all(p!='гараж казахстан' for p,m in stock)
    out={'finding':'S5-R1','registry_valid_rows':len(records),'selected_project_registry_matches':0,'expected':'empty preview; no DOCX; no catalog error'}
    try:C.preview('Гараж Казахстан','9/2026',root/'Документы/Гараж казахстан/source.xls')
    except I.Stage5Error as exc:
        out['actual']=str(exc);out['reproduced']='MARK_PROFILE_CONFLICT: UNRELATED_MARK' in str(exc)
    else:out['reproduced']=False
after={str(p):I.sha(p) for p in root.rglob('*') if p.is_file()}
out['fixture_unchanged']=before==after
out['generated_docx']=list(root.glob('validation/**/*.docx'))
assert out['reproduced'] and out['fixture_unchanged'] and not out['generated_docx']
(RUN/'S5-R1_reproduction.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf8')
print(json.dumps(out,ensure_ascii=False,indent=2))
