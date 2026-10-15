"""Fixed-path USER E2E wrapper. Calls an unchanged copy of Stage 2 product code."""
import argparse
import hashlib
import importlib
import json
import os
import stat
import sys
from pathlib import Path

EXPECTED_ROOT = Path(r'C:\Users\arara\Desktop\ВЗМК рабочий стол\Работа\Документы\Pasport\e2e_user_stage2')
ROOT = Path(__file__).absolute().parent
TEST_ROOT = ROOT / 'validation' / 'user_e2e'
REGISTRY = TEST_ROOT / 'ВЗМК.xlsx'
DOCUMENTS = TEST_ROOT / 'Документы'
PASSPORTS = DOCUMENTS / 'Гараж казахстан' / 'Паспорт'
TEMPLATE = PASSPORTS / '1_2026 - ВИК.docx'
FALLBACK = TEST_ROOT / 'Шаблоны' / 'ВИК.docx'
OUTPUT = PASSPORTS / '9-2026 от 06.10.2026' / '9_2026 - ВИК.docx'
PROJECT = 'Гараж Казахстан'
NUMBER = '9/2026'
EXPECTED = {
    'number': NUMBER,
    'contract': '№ 06/26 от 03.06.2026',
    'object': 'Архитектурно-строительные решения здания парковки подвижного состава',
    'code': 'WAA0007144-AC2-КМД',
    'steel': 'С255, С345',
    'connections': 'Балка',
    'vik_date': '29.09.2026',
    'end_date': '06.10.2026',
}


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def check_path(path):
    path = Path(path)
    if not path.is_relative_to(EXPECTED_ROOT) or path.resolve() != path:
        raise ValueError(f'Путь выходит из фиксированной E2E-копии или перенаправлен: {path}')
    for entry in (path, *path.parents):
        if entry.exists():
            info = entry.lstat()
            if stat.S_ISLNK(info.st_mode) or getattr(info, 'st_file_attributes', 0) & 0x400:
                raise ValueError(f'Ссылка/junction запрещена в E2E-пути: {entry}')


def check_tree():
    if ROOT != EXPECTED_ROOT or ROOT.resolve() != ROOT:
        raise ValueError('E2E launcher перемещён или перенаправлен. Запуск остановлен.')
    check_path(ROOT)
    for directory, dirs, files in os.walk(ROOT, followlinks=False):
        for name in [*dirs, *files]:
            check_path(Path(directory) / name)


def frozen_files(checkpoint):
    for relative, expected in checkpoint['copied_files'].items():
        path = ROOT / relative
        check_path(path)
        if not path.is_file() or digest(path) != expected:
            raise ValueError(f'Изменилась принятая E2E-копия: {path}. Автоматического сброса нет.')
    for absolute, expected in checkpoint['protected_before'].items():
        if digest(absolute) != expected:
            raise ValueError(f'Изменился production/checkpoint/защищённый источник: {absolute}')


def runtime_preflight():
    print('Python sys.executable: ' + sys.executable)
    print('Python version: ' + sys.version.split()[0])
    failures = []
    for name in ('win32com.client', 'win32con', 'pythoncom'):
        try:
            importlib.import_module(name)
        except Exception as exc:
            print(name + ' — FAIL: ' + str(exc), flush=True)
            failures.append(name)
        else:
            print(name + ' — PASS', flush=True)
    if failures:
        raise RuntimeError('COM preflight FAIL: ' + ', '.join(failures)
                           + '. Нужна рабочая pywin32 в указанном Python. '
                           + 'Генерация не запускалась.')


def preflight():
    check_tree()
    checkpoint = json.loads((ROOT / 'E2E_CHECKPOINT.json').read_text(encoding='utf-8'))
    if checkpoint['root'] != str(ROOT) or checkpoint['test_root'] != str(TEST_ROOT):
        raise ValueError('E2E checkpoint содержит другие пути.')
    if checkpoint['expected_requisites'] != EXPECTED:
        raise ValueError('E2E checkpoint содержит другие контрольные реквизиты.')
    frozen_files(checkpoint)
    for path in (REGISTRY, TEMPLATE, FALLBACK, OUTPUT, TEST_ROOT / '.vik-stage'):
        check_path(path)
    if OUTPUT.exists() or (OUTPUT.parent / ('~$' + OUTPUT.name)).exists():
        raise ValueError(f'ВИК уже существует или открыт. Перезапись/автоматический сброс запрещены: {OUTPUT}')
    if not OUTPUT.parent.is_dir():
        raise ValueError('Нет принятой папки текущего Stage 1 паспорта.')

    # Import only the hash-checked application copy after all path checks.
    sys.path.insert(0, str(ROOT))
    from passport_app.config import Settings
    from passport_app.excel import load_book
    from passport_app.registry import read_registry, select_project
    from passport_app import vik
    if Path(vik.__file__).resolve() != ROOT / 'passport_app' / 'vik.py' or vik.APP_ROOT != ROOT:
        raise ValueError('Загружена другая копия product code.')
    settings = Settings(DOCUMENTS, REGISTRY,
                        {PROJECT: {'passport_root': 'Гараж казахстан/Паспорт'}})
    vik.validation_boundary(settings, TEST_ROOT)
    registry = read_registry(REGISTRY)
    project = select_project(registry, PROJECT)
    rows = [r for r in registry.projects[PROJECT] if r['number'] == NUMBER]
    if len(rows) != 1 or project.first_number != '1/2026' or project.first_row != 353:
        raise ValueError('Не подтверждена текущая строка 9/2026 или первый паспорт проекта.')
    selected, origin = vik.find_vik_template(PASSPORTS, project, FALLBACK)
    if selected != TEMPLATE or origin != 'project_first_passport':
        raise ValueError('Template source отличается от первого паспорта проекта.')
    columns = vik._columns(registry)
    formula_book = load_book(REGISTRY)
    values_book = load_book(REGISTRY, values=True)
    try:
        row = rows[0]['row']
        formula = formula_book[registry.sheet].cell(row, columns['vik_date'])
        if formula.data_type != 'f':
            raise ValueError('ВИК Дата не содержит существующую Excel-формулу.')
        sheet = values_book[registry.sheet]
        actual = {}
        for key, column in columns.items():
            value = sheet.cell(row, column).value
            if key in {'vik_date', 'end_date'}:
                value = value.strftime('%d.%m.%Y') if hasattr(value, 'strftime') else value
            actual[key] = value
        if actual != EXPECTED:
            raise ValueError('Контрольная строка E2E не совпадает с принятым Stage 1 output.')
    finally:
        formula_book.close()
        values_book.close()
    return checkpoint, settings, vik


def show_paths():
    print('\nТЕСТОВЫЙ E2E — документы формируются только в e2e_user_stage2.')
    print('Проект = Гараж Казахстан; Паспорт = 9/2026')
    print('Дата окончания = 06.10.2026; ВИК Дата = 29.09.2026; Соединения = Балка')
    print('ВИК Дата показана из сохранённого результата формулы; product выполнит пересчёт копии Excel.')
    print('Template source = первый паспорт проекта (1/2026)')
    print('\nACTIVE PATHS')
    print('Registry: ' + str(REGISTRY))
    print('Project VIK: ' + str(TEMPLATE))
    print('Output: ' + str(OUTPUT))


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description='Безопасный USER E2E Stage 2 на фиксированных копиях')
    parser.add_argument('--check-only', action='store_true')
    args = parser.parse_args()
    runtime_preflight()
    checkpoint, settings, vik = preflight()
    show_paths()
    if args.check_only:
        print('\nPREFLIGHT PASS. Генерация и открытие Word: NOT_RUN.')
        return 0
    if input('\nВведите 1 и нажмите Enter, чтобы сформировать ВИК (Enter — выход): ').strip() != '1':
        print('Формирование не запускалось.')
        return 0
    checkpoint, settings, vik = preflight()
    try:
        # The real, unchanged Stage 2 generator. No alternative E2E generation.
        result = vik.generate_vik(settings, PROJECT, NUMBER,
                                  test_root=TEST_ROOT, fallback_template=FALLBACK)
    finally:
        frozen_files(checkpoint)
    if Path(result['output']) != OUTPUT:
        raise ValueError('Product вернул неожиданный output path.')
    for key in ('output', 'evidence', 'stage'):
        check_path(Path(result[key]))
    evidence = json.loads(Path(result['evidence']).read_text(encoding='utf-8'))
    if (evidence['requisites'] != EXPECTED or evidence['template'] != str(TEMPLATE)
            or evidence['template_origin'] != 'project_first_passport'
            or evidence['pages'] != 2 or evidence['output_sha256'] != digest(OUTPUT)):
        raise ValueError('Контрольный результат отличается от ожидаемого. Успех не подтверждён.')
    frozen_files(checkpoint)
    report = {'status': 'GENERATION_PASS_VISUAL_USER_CHECK_PENDING', **result,
              'output_sha256': digest(OUTPUT), 'production_hashes_unchanged': True,
              'user_visual_e2e': 'NOT_RUN'}
    (ROOT / 'E2E_RESULT.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print('\nВИК успешно создан. Production и исходные копии не изменились.')
    print(str(OUTPUT))
    print('Проверьте в Word ровно 2 страницы, реквизиты и оформление по инструкции.')
    os.startfile(str(OUTPUT))
    return 0


if __name__ == '__main__':
    try:
        exit_code = main()
    except Exception as exc:
        print('STOP: ' + str(exc), file=sys.stderr)
        exit_code = 1
    raise SystemExit(exit_code)
