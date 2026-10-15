import re
from datetime import date, timedelta
from dataclasses import dataclass
from pathlib import Path
from .config import PassportError
from .excel import load_book, sha256, open_excel, assert_no_formula_errors
from .tables import norm


REQUIRED = {'project': 'примечание', 'number': '№ паспорта', 'mass': 'масса кг',
            'date': 'дата окончания', 'connections': 'соединения', 'object': 'наименование объекта'}


@dataclass(frozen=True)
class Registry:
    path: Path
    digest: str
    sheet: str
    table: str
    columns: dict
    projects: dict
    width: int


@dataclass(frozen=True)
class Project:
    name: str
    first_row: int
    first_number: str
    first_date: object
    object_name: str
    last_row: int
    last_number: str
    next_number: str


def read_registry(path):
    initial_digest = sha256(path)
    book = load_book(path)
    try:
        candidates = []
        for sheet in book.worksheets:
            for table in sheet.tables.values():
                header_row = sheet[table.ref.split(':')[0]].row
                headers = {}
                for cell in sheet[header_row]:
                    headers.setdefault(norm(cell.value), []).append(cell.column)
                if all(label in headers for label in REQUIRED.values()):
                    if any(len(headers[label]) != 1 for label in REQUIRED.values()):
                        raise PassportError('В реестре повторяются обязательные заголовки.')
                    columns = {key: headers[label][0] for key, label in REQUIRED.items()}
                    candidates.append((sheet, table, columns, header_row))
        if len(candidates) != 1:
            raise PassportError('В реестре не найдена единственная Excel-таблица с обязательными столбцами.')
        sheet, table, columns, header_row = candidates[0]
        if sheet.protection.sheet:
            raise PassportError('Лист реестра защищён. Автоматическое снятие защиты не выполняется.')
        last_row = sheet[table.ref.split(':')[1]].row - (1 if table.totalsRowShown else 0)
        projects = {}
        for r in range(header_row + 1, last_row + 1):
            key = sheet.cell(r, columns['project']).value
            if key is not None and str(key).strip():
                projects.setdefault(str(key), []).append({
                    key: sheet.cell(r, col).value for key, col in columns.items()
                } | {'row': r})
        if sha256(path) != initial_digest:
            raise PassportError('Реестр изменился во время чтения. Повторите проверку.')
        return Registry(Path(path).resolve(), initial_digest, sheet.title, table.name, columns,
                        projects, sheet.max_column)
    finally:
        book.close()


def select_project(registry, name):
    entries = registry.projects.get(name)
    if not entries:
        raise PassportError('Выберите проект из реестра.')
    parsed = []
    for e in entries:
        match = re.fullmatch(r'([1-9]\d*)/(\d{4})', str(e['number']))
        if not match:
            raise PassportError(f'Номер {e["number"]!r}, строка {e["row"]}, не соответствует формату число/год.')
        parsed.append((int(match[1]), match[2], e))
    if len({e['number'] for e in entries}) != len(entries):
        raise PassportError('У проекта есть повторяющиеся номера паспортов.')
    first = [e for n, year, e in parsed if n == 1]
    if len(first) != 1:
        raise PassportError('Нельзя однозначно определить первый паспорт проекта: требуется единственный номер 1/год.')
    highest = max(n for n, year, e in parsed)
    last = [(year, e) for n, year, e in parsed if n == highest]
    if len(last) != 1:
        raise PassportError('Нельзя однозначно определить последний паспорт: совпадает числовая часть номера.')
    year, end = last[0]
    start = first[0]
    if not hasattr(start['date'], 'date') or not str(start['object'] or '').strip():
        raise PassportError('В первой строке проекта отсутствует дата окончания или наименование объекта.')
    next_number = f'{highest + 1}/{year}'
    if any(e['number'] == next_number for e in entries):
        raise PassportError('Предлагаемый номер уже существует.')
    return Project(name, start['row'], start['number'], start['date'].date(), str(start['object']),
                   end['row'], end['number'], next_number)


def append_row(excel, stage_path, registry, project, vom, day):
    book = open_excel(excel, stage_path)
    try:
        sheet = book.Worksheets(registry.sheet)
        table = sheet.ListObjects(registry.table)
        source = sheet.Rows(project.first_row)
        formulas = {c: sheet.Cells(project.first_row, c).FormulaR1C1
                    for c in range(1, registry.width + 1) if sheet.Cells(project.first_row, c).HasFormula}
        new = table.ListRows.Add(AlwaysInsert=True)
        target_row = new.Range.Row
        source.Copy(Destination=sheet.Rows(target_row))
        sheet.Rows(target_row).RowHeight = source.RowHeight
        sheet.Rows(target_row).Hidden = source.Hidden
        changed = {'number': project.next_number, 'mass': float(vom.mass),
                   'date': (day - date(1899, 12, 30)).days,
                   'connections': vom.connections}
        for key, value in changed.items():
            sheet.Cells(target_row, registry.columns[key]).Value2 = value
        excel.CalculateFull()
        for col, formula in formulas.items():
            cell = sheet.Cells(target_row, col)
            if not cell.HasFormula or cell.FormulaR1C1 != formula:
                raise PassportError(f'Не сохранилась исходная формула в новой строке, столбец {col}.')
        assert_no_formula_errors(sheet, target_row, registry.width)
        # The appendix gets all its header requisites from this new registry row.
        data = {key: sheet.Cells(target_row, col).Value2 for key, col in registry.columns.items()}
        data['date'] = date(1899, 12, 30) + timedelta(days=int(data['date']))
        if data['number'] != project.next_number or data['date'] != day:
            raise PassportError('Не совпадают реквизиты записанной строки.')
        book.Save()
        return target_row, data
    finally:
        book.Close(SaveChanges=False)
