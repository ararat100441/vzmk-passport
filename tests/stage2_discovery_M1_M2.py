"""Consolidated M1/M2 discovery checks; one native control run on copies."""
import ast
import hashlib
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from passport_app.config import Settings, PassportError
from passport_app.excel import sha256
from passport_app.registry import read_registry, select_project
from passport_app.vik import find_vik_template, general_vik_template_path, package_layout, NS

BASE = Path(__file__).resolve().parents[1]
PROJECT_VIK = BASE.parent / 'Гараж казахстан/Паспорт/1_2026 - ВИК.docx'
GENERAL_VIK = BASE.parent.parent / 'Шаблоны/ВИК.docx'
PROJECT = SimpleNamespace(first_number='1/2026')


def copy(path, source=PROJECT_VIK):
    path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, path)
    return path.resolve()


def main():
    run = Path(sys.argv[1]).resolve()
    assert run.is_relative_to(BASE / 'validation') and run.is_dir()
    original = json.loads((run / 'protected_before.json').read_text(encoding='utf-8'))
    allowed = [BASE / 'passport_app/vik.py', BASE / 'stage2.py', BASE / 'tests/stage2_remaining.py']
    expected = dict(original)
    for p in allowed:
        expected[str(p)] = sha256(p)
    # Protect the new runner too for the duration of this run.
    expected[str(Path(__file__).resolve())] = sha256(__file__)

    def guard():
        mismatches = [p for p, h in expected.items() if sha256(p) != h]
        assert not mismatches, 'STOP: unexpected protected changes: ' + repr(mismatches)

    guard()
    (run / 'protected_expected.json').write_text(json.dumps(expected, ensure_ascii=False, indent=2), encoding='utf-8')
    results = {'scope': 'M1/M2 discovery only, one native control scenario', 'discovery_checks': []}
    discovery_only = '--discovery-only' in sys.argv[2:]

    def finish():
        results['protected_hashes_unchanged'] = True
        results['allowed_source_changes'] = {str(p): {'before': original[str(p)], 'after': sha256(p)} for p in allowed}
        guard()
        after = {p: sha256(p) for p in expected}
        (run / 'protected_after.json').write_text(json.dumps(after, ensure_ascii=False, indent=2), encoding='utf-8')
        (run / 'targeted_results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
        print('real_Garage_Kazakhstan: PASS', flush=True)
        print('control_stage2: PASS', flush=True)
        print('protected_hashes: PASS', flush=True)

    def record(name, detail):
        guard()
        results['discovery_checks'].append({'name': name, 'status': 'PASS', **detail})
        (run / 'targeted_results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
        print(name + ': PASS', flush=True)

    def paths(name):
        case = run / ('cases_retest' if discovery_only else 'cases') / name
        root = case / 'Паспорт'
        root.mkdir(parents=True)
        return root, case / 'Шаблоны/ВИК.docx'

    def selected(name, root, fallback, expected_path, expected_origin):
        before = {str(p): sha256(p) for p in root.parent.rglob('*') if p.is_file()}
        path, origin = find_vik_template(root, PROJECT, fallback)
        assert path == expected_path and origin == expected_origin, (name, path, origin)
        assert all(sha256(p) == h for p, h in before.items())
        record(name, {'selected': str(path), 'origin': origin, 'inputs_unchanged': True})

    def rejected(name, root, fallback, phrase):
        try:
            find_vik_template(root, PROJECT, fallback)
        except PassportError as exc:
            assert phrase in str(exc), str(exc)
            record(name, {'error': str(exc)})
        else:
            raise AssertionError(name + ': an invalid template was selected')

    root, fallback = paths('M1_A_correct_year')
    candidate = copy(root / '1-2026 от 09.09.2026/ВИК.docx')
    selected('M1_A_correct_year', root, fallback, candidate, 'project_first_passport')

    root, fallback = paths('M1_B_wrong_year')
    copy(root / '1-2025 от 01.01.2025/ВИК.docx')  # Its content is 1/2026: folder mismatch still excludes it.
    rejected('M1_B_wrong_year', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M1_C_both_years')
    copy(root / '1-2025 от 01.01.2025/ВИК.docx')
    candidate = copy(root / '1-2026 от 09.09.2026/ВИК.docx')
    selected('M1_C_both_years', root, fallback, candidate, 'project_first_passport')

    root, fallback = paths('M1_D_ambiguity')
    copy(root / '1-2026 от 09.09.2026/ВИК.docx')
    copy(root / '1_2026 от 09.09.2026/ВИК.docx')
    rejected('M1_D_ambiguity', root, fallback, '2 кандидатов')

    # Legacy location without a year: real content controls prove the complete number.
    root, fallback = paths('M1_legacy_yearless_valid')
    candidate = copy(root / '1 от 09.09.2026/ВИК.docx')
    selected('M1_legacy_yearless_valid', root, fallback, candidate, 'project_first_passport')

    root, fallback = paths('M1_legacy_yearless_wrong_content_year')
    candidate = copy(root / '1 от 09.09.2026/ВИК.docx')
    with zipfile.ZipFile(candidate) as z:
        entries = [(info, z.read(info.filename)) for info in z.infolist()]
    with zipfile.ZipFile(candidate, 'w') as z:
        for info, data in entries:
            if info.filename == 'word/document.xml':
                assert data.count(b'>1/2026<') == 2
                data = data.replace(b'>1/2026<', b'>1/2025<')
            z.writestr(info, data)
    rejected('M1_legacy_yearless_wrong_content_year', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M1_legacy_unproven_STOP')
    copy(root / '1 от 09.09.2026/ВИК.docx', GENERAL_VIK)
    copy(fallback, GENERAL_VIK)
    rejected('M1_legacy_unproven_STOP', root, fallback, 'Нельзя однозначно связать')

    root, fallback = paths('M1_wrong_year_file_boundary')
    copy(root / '1_2025-ВИК.docx')
    rejected('M1_wrong_year_file_boundary', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M1_wrong_year_file_suffix')
    copy(root / '1_2025x ВИК.docx')
    rejected('M1_wrong_year_file_suffix', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M1_nested_year_conflict')
    copy(root / '1-2026 от 09.09.2026/1-2025 от 01.01.2025/ВИК.docx')
    rejected('M1_nested_year_conflict', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M1_wrong_number')
    copy(root / '2-2026 от 09.09.2026/ВИК.docx')
    rejected('M1_wrong_number', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M2_A_exact')
    exact = copy(fallback, GENERAL_VIK)
    selected('M2_A_exact', root, fallback, exact, 'general_fallback')

    root, fallback = paths('M2_B_archive_only')
    copy(fallback.parent / 'Archive/ВИК старый.docx', GENERAL_VIK)
    rejected('M2_B_archive_only', root, fallback, 'отсутствует точный общий шаблон')

    root, fallback = paths('M2_C_exact_and_archive')
    exact = copy(fallback, GENERAL_VIK)
    copy(fallback.parent / 'Archive/ВИК старый.docx', GENERAL_VIK)
    selected('M2_C_exact_and_archive', root, fallback, exact, 'general_fallback')

    root, fallback = paths('M2_D_project_priority')
    candidate = copy(root / '1_2026 - ВИК.docx')
    copy(fallback, GENERAL_VIK)
    selected('M2_D_project_priority', root, fallback, candidate, 'project_first_passport')

    root, fallback = paths('M2_E_both_missing')
    rejected('M2_E_both_missing', root, fallback, 'отсутствует точный общий шаблон')

    settings = Settings.load(BASE / 'settings.json')
    registry = read_registry(settings.registry_path)
    project = select_project(registry, 'Гараж Казахстан')
    passport_root, _ = settings.project_paths(project.name)
    assert general_vik_template_path(settings) == GENERAL_VIK.resolve()
    assert general_vik_template_path(settings, fallback_root=GENERAL_VIK.parent) == GENERAL_VIK.resolve()
    assert general_vik_template_path(settings, fallback_template=GENERAL_VIK) == GENERAL_VIK.resolve()
    real, origin = find_vik_template(passport_root, project, general_vik_template_path(settings))
    assert real == PROJECT_VIK.resolve() and origin == 'project_first_passport'
    results['real_Garage_Kazakhstan'] = {'status': 'PASS', 'first_number': project.first_number,
                                        'selected': str(real), 'fallback_default': str(GENERAL_VIK.resolve())}
    guard()

    # Prove that substitutions, calculation, Word automation and sdtEndPr logic
    # were untouched. Only the discovery assignment in generate_vik may differ.
    old = ast.parse((run / 'vik.before.py').read_text(encoding='utf-8'))
    new = ast.parse((BASE / 'passport_app/vik.py').read_text(encoding='utf-8'))
    old_defs = {node.name: node for node in old.body if isinstance(node, ast.FunctionDef)}
    new_defs = {node.name: node for node in new.body if isinstance(node, ast.FunctionDef)}
    untouched = []
    for name, node in old_defs.items():
        if name not in {'find_vik_template', 'generate_vik'}:
            assert ast.dump(node) == ast.dump(new_defs[name]), name + ' changed outside discovery'
            untouched.append(name)
    def without_discovery(node):
        return [n for n in node.body if not (isinstance(n, ast.Assign) and any(isinstance(c, ast.Call) and isinstance(c.func, ast.Name) and c.func.id == 'find_vik_template' for c in ast.walk(n)))]
    assert [ast.dump(n) for n in without_discovery(old_defs['generate_vik'])] == [ast.dump(n) for n in without_discovery(new_defs['generate_vik'])]
    results['output_logic_unchanged'] = {'status': 'PASS', 'unchanged_functions': untouched,
                                        'generate_body_except_discovery_equal': True,
                                        'Stage1_config_unchanged': True, 'Stage1_smoke': 'NOT_RUN: no Stage 1/config changes; accepted smoke reused'}

    if discovery_only:
        prior = json.loads((run / 'targeted_results.initial_pass.json').read_text(encoding='utf-8'))
        control = prior['control_stage2']
        generation = json.loads(Path(control['evidence']).read_text(encoding='utf-8'))
        assert sha256(control['output']) == generation['output_sha256']
        control['evidence_reused_after_discovery_hardening'] = True
        results['control_stage2'] = control
        finish()
        return

    # One control generation, using the current accepted Stage 1 E2E row on a copy.
    control = run / 'control_stage2'
    control_passport = control / 'Документы/Гараж казахстан/Паспорт'
    control_passport.mkdir(parents=True)
    copy(control_passport / '1_2026 - ВИК.docx')
    general = copy(control / 'Шаблоны/ВИК.docx', GENERAL_VIK)
    copy(control / 'ВЗМК.xlsx', BASE / 'e2e_user_stage1/ВЗМК.xlsx')
    (control_passport / '9-2026 от 06.10.2026').mkdir()
    config = {'registry_path': 'ВЗМК.xlsx', 'documents_root': 'Документы',
              'projects': {'Гараж Казахстан': {'passport_root': 'Гараж казахстан/Паспорт'}}}
    config_path = control / 'settings.json'
    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding='utf-8')
    inputs = {str(p): sha256(p) for p in control.rglob('*') if p.is_file()}
    command = [sys.executable, str(BASE / 'stage2.py'), '--settings', str(config_path),
               '--test-root', str(control), '--project', 'Гараж Казахстан', '--number', '9/2026',
               '--fallback-template', str(general)]
    process = subprocess.run(command, capture_output=True, text=True, encoding='utf-8')
    (run / 'control_cli_stdout.txt').write_text(process.stdout, encoding='utf-8')
    (run / 'control_cli_stderr.txt').write_text(process.stderr, encoding='utf-8')
    guard()
    assert process.returncode == 0, process.stderr
    output = json.loads(process.stdout)
    generation = json.loads(Path(output['evidence']).read_text(encoding='utf-8'))
    req = generation['requisites']
    assert req['number'] == '9/2026' and req['vik_date'] == '29.09.2026' and req['connections'] == 'Балка'
    assert generation['template_origin'] == 'project_first_passport' and generation['pages'] == 2
    assert all(sha256(p) == h for p, h in inputs.items())
    accepted = json.loads((BASE / 'validation/stage2_20261006_084159/project_result.json').read_text(encoding='utf-8'))
    def paragraph_texts(path):
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read('word/document.xml'))
        return [''.join(t.text or '' for t in p.findall('.//w:t', NS)) for p in root.findall('.//w:p', NS)]
    texts = paragraph_texts(output['output'])
    assert texts == paragraph_texts(accepted['output'])
    assert package_layout(output['output']) == package_layout(accepted['output'])
    assert 'АКТ № 9/2026 от 29.09.2026' in texts and 'АКТ № 9/2026-1 от 29.09.2026' in texts
    results['control_stage2'] = {'status': 'PASS', **output, 'requisites': req,
                                'pages': 2, 'origin': generation['template_origin'],
                                'all_paragraph_texts_equal_to_accepted_output': True,
                                'layout_equal_to_accepted_output': True, 'inputs_unchanged': True,
                                'visual_evidence': 'REUSED: accepted visuals, unchanged output logic and matching text/layout'}
    finish()


if __name__ == '__main__':
    main()
