"""Baseline accepted product and all existing workbook/document inputs before tests."""
import sys,json,hashlib,os
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import save,bootstrap
from stage8_closure.storage import sha
OUT=ROOT/'validation/stage9_20261012_01'
OUT.mkdir(exist_ok=False)
files=[]
for folder in ('passport_app','profiles_stage4','uzk_stage5','akz_stage6','stage7_quality','stage8_closure'):
    files+=list((ROOT/folder).glob('*.py'))
files+=list(ROOT.glob('stage[1-8].py'))
files+=list((ROOT/'DEV').rglob('*.xlsx'))+list((ROOT/'DEV').rglob('*.xlsm'))+list((ROOT/'DEV').rglob('*.json'))
project=ROOT.parent/'Гараж казахстан'
files+=list(project.rglob('*.xlsx'))+list(project.rglob('*.xls'))+list(project.rglob('*.xlsm'))+list(project.rglob('*.docx'))
files += [ROOT.parent.parent/'ВЗМК.xlsx',ROOT/'settings.json']
before={str(p):sha(p) for p in dict.fromkeys(files) if p.is_file() and not p.name.startswith('~$')}
save(OUT/'PROTECTED_BEFORE.json',before)
work=bootstrap(ROOT,Path(os.environ['LOCALAPPDATA'])/'VZMKPassport'/'workspaces')
save(OUT/'WORKSPACE.json',{'path':str(work)})
print(work)
