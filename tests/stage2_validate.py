"""Targeted Stage 2 E2E on a new copy, never on accepted Stage 1 evidence."""
import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from passport_app.config import Settings, PassportError
from passport_app.excel import sha256
from passport_app.vik import generate_vik


def prepare(base, run, name, fallback=False):
    root = run / name
    passport = root / 'Документы/Гараж казахстан/Паспорт'
    passport.mkdir(parents=True)
    shutil.copy2(run / 'ВЗМК.xlsx', root / 'ВЗМК.xlsx')
    if not fallback:
        shutil.copy2(run / 'template.docx', passport / '1_2026 - ВИК.docx')
    current = passport / '9-2026 от 06.10.2026'
    current.mkdir()
    source_current = base / 'e2e_user_stage1/Документы/Гараж казахстан/Паспорт/9-2026 от 06.10.2026'
    for file in source_current.iterdir():
        if file.is_file():
            shutil.copy2(file, current / file.name)
    settings = {'registry_path': 'ВЗМК.xlsx', 'documents_root': 'Документы',
                'projects': {'Гараж Казахстан': {'passport_root': 'Гараж казахстан/Паспорт'}}}
    (root / 'settings.json').write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding='utf-8')
    return root, Settings.load(root / 'settings.json')


def main():
    base = Path(__file__).resolve().parents[1]
    run = Path(sys.argv[1]).resolve()
    before = json.loads((run / 'protected_before.json').read_text(encoding='utf-8'))
    general = base.parent.parent / 'Шаблоны/ВИК.docx'
    before[str(general)] = sha256(general)
    (run / 'protected_before.json').write_text(json.dumps(before, ensure_ascii=False, indent=2), encoding='utf-8')
    root, settings = prepare(base, run, 'project_e2e')
    result = generate_vik(settings, 'Гараж Казахстан', '9/2026', test_root=root,
                          fallback_root=general.parent)
    (run / 'project_result.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    after = {p: sha256(p) for p in before}
    (run / 'protected_after.json').write_text(json.dumps(after, ensure_ascii=False, indent=2), encoding='utf-8')
    assert before == after
    evidence = json.loads(Path(result['evidence']).read_text(encoding='utf-8'))
    assert evidence['requisites']['vik_date'] == '29.09.2026'
    assert evidence['requisites']['connections'] == 'Балка'
    assert sha256(root / 'ВЗМК.xlsx') == sha256(run / 'ВЗМК.xlsx')
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
