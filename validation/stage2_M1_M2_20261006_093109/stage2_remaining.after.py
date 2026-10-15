"""Resume-only targeted checks. Preserve Stage 1 and pre-STOP evidence."""
import json
import re
import shutil
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path
from unittest.mock import patch
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from passport_app.config import Settings, PassportError
from passport_app.excel import sha256
from passport_app.service import validate
from passport_app.vik import generate_vik

BASE = Path(__file__).resolve().parents[1]
PREVIOUS = BASE / 'validation/stage2_20261006_084159'
CURRENT_HASH = 'b0ece2f9e5cbb1db73a0487da320c4e69c0c8fd815345a417bd9482fd7029b0a'
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def assert_protected(before):
    mismatches = {p: {'expected': h, 'actual': sha256(p)} for p, h in before.items() if sha256(p) != h}
    if mismatches:
        raise RuntimeError('STOP: protected source changed: ' + json.dumps(mismatches, ensure_ascii=False))


def prepare(run, name, *, project_template=True, general_template=False):
    root = run / name
    passport = root / 'Документы/Гараж казахстан/Паспорт'
    passport.mkdir(parents=True, exist_ok=True)
    shutil.copy2(PREVIOUS / 'ВЗМК.xlsx', root / 'ВЗМК.xlsx')
    if project_template:
        shutil.copy2(PREVIOUS / 'template.docx', passport / '1_2026 - ВИК.docx')
    if general_template:
        general = root / 'Шаблоны'
        general.mkdir(exist_ok=True)
        shutil.copy2(BASE.parent.parent / 'Шаблоны/ВИК.docx', general / 'ВИК.docx')
    (passport / '9-2026 от 06.10.2026').mkdir(exist_ok=True)
    settings = {'registry_path': 'ВЗМК.xlsx', 'documents_root': 'Документы',
                'projects': {'Гараж Казахстан': {'passport_root': 'Гараж казахстан/Паспорт'}}}
    (root / 'settings.json').write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding='utf-8')
    return root, Settings.load(root / 'settings.json')


def patch_formula(path, formula):
    """Change only a disposable fixture's M400 formula, never a source book."""
    assert path.resolve().is_relative_to(BASE / 'validation')
    with zipfile.ZipFile(path) as z:
        entries = [(info, z.read(info.filename)) for info in z.infolist()]
    part = 'xl/worksheets/sheet1.xml'
    updated = []
    for info, data in entries:
        if info.filename == part:
            # Keep the original namespace declarations, including prefixes named
            # in mc:Ignorable. Re-serializing the entire sheet breaks that contract.
            text = data.decode('utf-8')
            pattern = r'(<c\b[^>]*\br="M400"[^>]*>)(.*?)(</c>)'
            matches = list(re.finditer(pattern, text, re.S))
            assert len(matches) == 1
            match = matches[0]
            content, count = re.subn(r'<f(?:\s[^>]*)?>.*?</f>', '<f>' + escape(formula) + '</f>', match[2], flags=re.S)
            assert count == 1
            content = re.sub(r'<v(?:\s[^>]*)?>.*?</v>', '', content, flags=re.S)
            text = text[:match.start()] + match[1] + content + match[3] + text[match.end():]
            data = text.encode('utf-8')
        updated.append((info, data))
    temp = path.with_suffix('.fixture.tmp')
    with zipfile.ZipFile(temp, 'w') as z:
        for info, data in updated:
            z.writestr(info, data)
    temp.replace(path)


def main():
    resuming = len(sys.argv) > 1
    run = Path(sys.argv[1]).resolve() if resuming else BASE / 'validation' / ('stage2_resume_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
    assert run.is_relative_to(BASE / 'validation')
    if not resuming:
        run.mkdir()
    before = json.loads(((run if resuming else PREVIOUS) / 'protected_before.json').read_text(encoding='utf-8'))
    production = str(BASE.parent.parent / 'ВЗМК.xlsx')
    before[production] = CURRENT_HASH
    for folder in [PREVIOUS]:
        for p in folder.rglob('*'):
            if p.is_file():
                before[str(p)] = sha256(p)
    for p in ['stage2.py', 'tests/stage2_audit.py',
              'tests/stage2_validate.py', 'tests/render_word_pages.ps1', 'STAGE2_CHECKPOINT.json']:
        before[str(BASE / p)] = sha256(BASE / p)
    if not resuming:
        before[str(BASE / 'passport_app/vik.py')] = sha256(BASE / 'passport_app/vik.py')
    elif 'product_fix_authorized' not in json.loads((run / 'remaining_results.json').read_text(encoding='utf-8')):
        assert sha256(BASE / 'passport_app/vik.py') == before[str(BASE / 'passport_app/vik.py')]
    else:
        recorded = json.loads((run / 'remaining_results.json').read_text(encoding='utf-8'))['product_fix_authorized']
        assert sha256(BASE / 'passport_app/vik.py') == recorded['after_sha256']
        # Keep source before/after evidence intact; current validation guards the
        # explicitly recorded targeted fix as well as all unchanged sources.
        before[str(BASE / 'passport_app/vik.py')] = recorded['after_sha256']
    assert_protected(before)
    if not resuming:
        (run / 'protected_before.json').write_text(json.dumps(before, ensure_ascii=False, indent=2), encoding='utf-8')
    (run / 'protected_validation_expected.json').write_text(json.dumps(before, ensure_ascii=False, indent=2), encoding='utf-8')
    results = json.loads((run / 'remaining_results.json').read_text(encoding='utf-8')) if resuming else {
        'scope': 'remaining Stage 2 only', 'negative_checks': [],
        'reused_project_generation': str(PREVIOUS / 'project_result.json')}
    if resuming:
        results['test_fixture_repair'] = 'TEST DEFECT: whole-sheet XML serialization removed namespace declarations referenced by mc:Ignorable. Fixture now patches only M400 formula bytes; product unchanged.'

    def save():
        (run / 'remaining_results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2, default=str), encoding='utf-8')

    def expect_error(name, root, settings, *, number='9/2026', contains, verifier_failure=False):
        assert_protected(before)
        original_files = {str(p): sha256(p) for p in root.rglob('*') if p.is_file()}
        error = None
        try:
            if verifier_failure:
                # Actual Excel and Word generation/save/reopen; only the final
                # verification decision is deliberately failed before publication.
                with patch('passport_app.vik.verify_document', side_effect=PassportError('Тестовый отказ финальной проверки DOCX')):
                    generate_vik(settings, 'Гараж Казахстан', number, test_root=root)
            else:
                generate_vik(settings, 'Гараж Казахстан', number, test_root=root)
        except PassportError as exc:
            error = str(exc)
        assert error is not None and contains in error, (name, error)
        assert all(sha256(p) == h for p, h in original_files.items()), name + ': input modified'
        output = root / 'Документы/Гараж казахстан/Паспорт/9-2026 от 06.10.2026/9_2026 - ВИК.docx'
        if str(output) not in original_files:
            assert not output.exists(), name + ': failed output published'
        assert not list(root.rglob('.vik-pending-*.docx')), name + ': temporary DOCX left'
        assert not list(root.rglob('generation.json')), name + ': failed run marked successful'
        assert_protected(before)
        results['negative_checks'].append({'name': name, 'status': 'PASS', 'error': error,
                                           'input_hashes_unchanged': True, 'no_successful_publication': True,
                                           'verification_failure_injected': verifier_failure})
        save()
        print(name + ': PASS', flush=True)

    completed = {c['name'] for c in results['negative_checks'] if c['status'] == 'PASS'}
    if 'missing_project_and_fallback' not in completed:
        root, settings = prepare(run, 'missing_templates', project_template=False)
        expect_error('missing_project_and_fallback', root, settings, contains='отсутствует точный общий шаблон')
    if 'current_row_not_found' not in completed:
        root, settings = prepare(run, 'missing_current_row')
        expect_error('current_row_not_found', root, settings, number='99/2026', contains='Текущий паспорт не найден')
    for name, formula in [('blank_vik_date', '""'), ('error_vik_date', '1/0')]:
        if name in completed:
            continue
        root, settings = prepare(run, name)
        patch_formula(root / 'ВЗМК.xlsx', formula)
        expect_error(name, root, settings, contains='после пересчёта пустая или ошибочная')
    if 'existing_output_not_overwritten' not in completed:
        root, settings = prepare(run, 'existing_output')
        result = json.loads((PREVIOUS / 'project_result.json').read_text(encoding='utf-8'))
        existing = root / 'Документы/Гараж казахстан/Паспорт/9-2026 от 06.10.2026/9_2026 - ВИК.docx'
        shutil.copy2(result['output'], existing)
        expect_error('existing_output_not_overwritten', root, settings, contains='Перезапись запрещена')
    if 'failed_final_verification_not_published' not in completed:
        root, settings = prepare(run, 'failed_final_verification')
        expect_error('failed_final_verification_not_published', root, settings,
                     contains='Тестовый отказ финальной проверки', verifier_failure=True)

    assert_protected(before)
    root, settings = prepare(run, 'fallback_e2e', project_template=False, general_template=True)
    source = root / 'Шаблоны/ВИК.docx'
    template_hash = sha256(source)
    registry_hash = sha256(root / 'ВЗМК.xlsx')
    fallback_result = generate_vik(settings, 'Гараж Казахстан', '9/2026', test_root=root)
    evidence = json.loads(Path(fallback_result['evidence']).read_text(encoding='utf-8'))
    assert evidence['template_origin'] == 'general_fallback'
    assert evidence['pages'] == 2 and evidence['constant_text_preserved'] and evidence['layout_properties_preserved']
    assert evidence['requisites']['vik_date'] == '29.09.2026' and evidence['requisites']['connections'] == 'Балка'
    assert sha256(source) == template_hash and sha256(root / 'ВЗМК.xlsx') == registry_hash
    results['fallback'] = {'status': 'PASS', **fallback_result, 'source_template_sha256': template_hash,
                           'template_and_registry_unchanged': True}
    assert_protected(before)
    save()
    print('fallback: PASS', flush=True)

    # Minimal accepted Stage 1 validate(), no generation or full regression.
    root = run / 'stage1_smoke'
    passport = root / 'Документы/Гараж казахстан/Паспорт'
    passport.mkdir(parents=True)
    shutil.copy2(BASE.parent.parent / 'ВЗМК.xlsx', root / 'ВЗМК.xlsx')
    prod_passport = BASE.parent / 'Гараж казахстан/Паспорт'
    shutil.copy2(prod_passport / '1_2026 - Приложение № 1.xlsx', passport / '1_2026 - Приложение № 1.xlsx')
    shutil.copy2(prod_passport / '8 от 02.10.2026/ВОМ отгрузочный.xlsx', root / 'ВОМ.xlsx')
    settings = Settings(root / 'Документы', root / 'ВЗМК.xlsx',
                        {'Гараж Казахстан': {'passport_root': 'Гараж казахстан/Паспорт',
                         'template': 'Гараж казахстан/Паспорт/1_2026 - Приложение № 1.xlsx'}})
    smoke_before = {str(p): sha256(p) for p in root.rglob('*') if p.is_file()}
    plan = validate(settings, 'Гараж Казахстан', root / 'ВОМ.xlsx', today=date(2026, 10, 6))
    assert plan.project.first_number == '1/2026' and plan.project.last_number == '8/2026' and plan.project.next_number == '9/2026'
    assert len(plan.vom.rows) == 32 and float(plan.vom.quantity) == 50 and float(plan.vom.mass) == 15071.44
    assert float(plan.vom.area) == 477.749 and plan.vom.connections == 'Балка'
    assert {str(p): sha256(p) for p in root.rglob('*') if p.is_file()} == smoke_before
    results['stage1_smoke'] = {'status': 'PASS', 'operation': 'unchanged Stage 1 validate()',
                               'first': plan.project.first_number, 'last': plan.project.last_number,
                               'next': plan.project.next_number, 'rows': 32, 'quantity': 50,
                               'mass_kg': str(plan.vom.mass), 'area_m2': str(plan.vom.area),
                               'connections': plan.vom.connections, 'no_generation': True,
                               'input_hashes_unchanged': True}
    assert_protected(before)
    after = {p: sha256(p) for p in before}
    (run / 'protected_after.json').write_text(json.dumps(after, ensure_ascii=False, indent=2), encoding='utf-8')
    results['protected_hashes_unchanged'] = before == after
    save()
    print('stage1_smoke: PASS', flush=True)
    print(str(run), flush=True)


if __name__ == '__main__':
    main()
