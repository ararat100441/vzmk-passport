"""Explicit copy-in; source files are only read under Windows guards."""
import json
import os
import re
from pathlib import Path
from uuid import uuid4
from stage8_closure import storage as S

PROJECT = 'Гараж Казахстан'

def save(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8') as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2, default=str)
        stream.flush(); os.fsync(stream.fileno())

def load(root):
    root = S.no_links(root)
    meta = S.read(root/'STAGE9_WORKSPACE.json')
    if (not re.fullmatch(r'[0-9a-f]{32}',root.name) or meta.get('schema') != 'STAGE9_WORKSPACE_V1' or meta.get('official_issue') is not False
        or meta.get('workspace_id') != root.name or meta.get('seal') != S.digest({k:v for k,v in meta.items() if k!='seal'})):
        raise S.Stop('Некорректный DEV-стенд. Запись запрещена.')
    for name in ('DEV','documents','validation'):
        S.no_links(root/name)
    return root, meta

def bootstrap(source, parent):
    source, parent = S.no_links(source), S.no_links(parent)
    # Each import creates a new named workspace; no reuse or overwrite.
    root = parent/uuid4().hex
    from passport_app.config import Settings
    from passport_app.registry import read_registry, select_project
    from uzk_stage5.result_samples import ACCEPTED_SOURCES
    from akz_stage6 import core as a, part2
    cfg = Settings.load(source/'settings.json')
    reg = read_registry(source/'DEV/ВЗМК_DEV.xlsx'); select_project(reg, PROJECT)
    passports, _ = cfg.project_paths(PROJECT)
    project = passports.parent
    inputs = {}
    def add(src, relative, expected=None):
        src = S.no_links(src)
        actual = S.sha(src)
        if expected and actual != expected: raise S.Stop('SOURCE_SHA_CONFLICT: '+str(src))
        inputs[relative] = {'source':str(src),'sha256':actual}
    from uzk_stage5.inputs import PINS as pins
    for path, pin in pins.items():
        relative = path.relative_to(Path(__file__).resolve().parents[1])
        add(source/relative, relative.as_posix(), pin)
    for name in ('profiles.user.json','profiles.catalog.json','profiles.build.json'):
        add(source/'DEV/Справочники'/name, 'DEV/Справочники/'+name)
    add(project/'Справочник марок.xlsx','documents/project/Справочник марок.xlsx')
    for name in ('1_2026 - Приложение № 1.xlsx','1_2026 - ВИК.docx','1_2026 - Документ о качестве.docx'):
        add(passports/name,'documents/project/Паспорт/'+name)
    for relative, pin, _ in ACCEPTED_SOURCES:
        add(project/relative,'documents/project/'+relative,pin)
    for path,pin in a.PINS.items(): add(passports/path.name,'documents/project/Паспорт/'+path.name,pin)
    for n, hashes in part2.PINS.items():
        for name,pin in zip((f'АКЗ-{n}.xlsx',f'Приложение к АКЗ-{n}.xlsx'),hashes):
            add(passports/name,'documents/project/Паспорт/'+name,pin)
    from stage7_quality.sources import PINS as quality_pins, TEMPLATE
    expected = quality_pins[TEMPLATE]
    if inputs['documents/project/Паспорт/'+TEMPLATE.name]['sha256'] != expected:
        raise S.Stop('QUALITY_TEMPLATE_SHA_CONFLICT')
    # Use the exact accepted assignment input, not the newest similar workbook.
    preview = S.read(source/'validation/stage6_fix_r1_r2_20261009_094707/current_preview.json')
    if preview['seal'] != a.digest(preview['payload']): raise S.Stop('ACCEPTED_AKZ_PREVIEW_DAMAGED')
    candidates = [p for p in preview['payload']['sources'] if 'Stage6_Fixtures' in p and p.endswith('.xlsx')]
    if len(candidates)!=1: raise S.Stop('AKZ_ASSIGNMENT_INPUT_AMBIGUOUS')
    add(candidates[0],'DEV/Stage6_Fixtures/assignments.xlsx',preview['payload']['sources'][candidates[0]])
    root.mkdir(parents=True, exist_ok=False)
    with S.frozen([v['source'] for v in inputs.values()]):
        for relative, ref in inputs.items():
            data = Path(ref['source']).read_bytes()
            import hashlib
            if hashlib.sha256(data).hexdigest()!=ref['sha256']: raise S.Stop('INPUT_CHANGED_DURING_IMPORT')
            target=root/relative;target.parent.mkdir(parents=True,exist_ok=True)
            S.exclusive_bytes(target,data)
            if S.sha(target)!=ref['sha256']: raise S.Stop('COPY_SHA_CONFLICT')
    (root/'validation').mkdir()
    meta={'schema':'STAGE9_WORKSPACE_V1','workspace_id':root.name,'mode':'DEV_TEST_ONLY',
          'official_issue':False,'project':PROJECT,'inputs':inputs,'created_at':S.now()}
    save(root/'STAGE9_WORKSPACE.json',meta|{'seal':S.digest(meta)})
    return root

def bind(root):
    """Bind accepted storage boundaries to the owned DEV copy, never to sources.

    Only paths are adapted. Original template pins and business functions remain.
    One process has one workspace, eliminating concurrent global rebinding.
    """
    root,meta=load(root)
    from passport_app import vik
    from profiles_stage4 import storage as profiles
    from stage8_closure import fixtures
    from akz_stage6 import core as a, part2
    from stage7_quality import word
    vik.APP_ROOT=root
    profiles.ROOT=root
    profiles.DEFAULT_WORKBOOK=root/'DEV/Справочники/Профили.xlsx'
    S.ROOT=root;fixtures.ROOT=root
    a.ROOT=root;a.DEV=root/'DEV';a.REGISTRY=root/'DEV/ВЗМК_DEV.xlsx'
    a.PROJECT=root/'documents/project';a.FIXTURES=root/'DEV/Stage6_Fixtures'
    a.OUTPUT=root/'DEV/Stage6_Выход';a.OUTPUT.mkdir(exist_ok=True)
    old=dict(a.PINS)
    a.ACT=a.PROJECT/'Паспорт/Акт № 4-2026.xlsx'
    a.APPLICATION=a.PROJECT/'Паспорт/1_2026 - Приложение к АКЗ.xlsx'
    a.PINS={a.PROJECT/'Паспорт'/p.name:h for p,h in old.items()}
    part2.OUTPUT=a.OUTPUT/'Part2'
    pin=word.PINS[word.TEMPLATE]
    word.TEMPLATE=a.PROJECT/'Паспорт/1_2026 - Документ о качестве.docx'
    word.PINS={word.TEMPLATE:pin}
    return root,meta
