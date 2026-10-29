"""Stage 2: native Word generation, currently restricted to validation copies.

No VOM input and no Stage 1 mutation. The registry is recalculated in a private
Excel copy. The published DOCX is checked after a normal Word reopen.
"""
import hashlib
import json
import math
import os
import re
import uuid
import zipfile
from contextlib import ExitStack, contextmanager
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

from .config import PassportError
from .excel import excel_session, locked_file, open_excel, sha256
from .registry import read_registry, select_project
from .service import ensure_no_pending
from .tables import norm

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = {'w': W[1:-1]}
FIELDS = {'number': '№ паспорта', 'contract': 'Договор №',
          'object': 'Наименование объекта', 'code': 'шифр КМД',
          'steel': 'Марка стали', 'connections': 'Соединения', 'vik_date': 'ВИК Дата',
          'end_date': 'Дата окончания'}
TAGS = {'PASSPORT_NO': ('number', 2), 'VIK_DATE': ('vik_date', 2),
        'DOGOVOR_NO': ('contract', 1), 'KMD_CODE': ('code', 1), 'STEEL_MARK': ('steel', 1)}
TOKENS = {key: '{' + label + '}' for key, label in FIELDS.items()}
APP_ROOT = Path(__file__).resolve().parents[1]


def validation_boundary(settings, test_root):
    root = Path(test_root).resolve()
    validation = (APP_ROOT / 'validation').resolve()
    if root == validation or not root.is_relative_to(validation) or not root.is_dir():
        raise PassportError('Этап 2 разрешён только в отдельной копии внутри Pasport/validation.')
    if not settings.registry_path.resolve().is_relative_to(root) or not settings.documents_root.resolve().is_relative_to(root):
        raise PassportError('Реестр и документы должны находиться внутри Stage 2 validation-копии.')
    return root


def _passport_name_number(name):
    """Read an explicit number/year, or a legacy number without a year."""
    match = re.match(r'^(?:Паспорт\s*(?:№\s*)?|№\s*)?(\d+)(?=$|[\s._-])', name, re.I)
    if match is None:
        return None
    year = re.match(r'^[-_ ](\d+)', name[match.end():])
    return int(match[1]), year[1] if year else None


def _legacy_vik_number_matches(path, canonical):
    """A yearless location needs both real PASSPORT_NO controls as evidence."""
    digest = sha256(path)
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read('word/document.xml'))
        numbers = []
        for control in root.iter(W + 'sdt'):
            tag = control.find('w:sdtPr/w:tag', NS)
            if tag is not None and tag.get(W + 'val') == 'PASSPORT_NO':
                numbers.append(''.join(t.text or '' for t in control.findall('w:sdtContent//w:t', NS)))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise PassportError(f'Нельзя подтвердить номер первого паспорта в ВИК: {path}.') from exc
    if sha256(path) != digest:
        raise PassportError('ВИК изменился во время проверки номера первого паспорта.')
    if len(numbers) != 2 or any(not re.fullmatch(r'[1-9]\d*/\d{4}', n) for n in numbers) or numbers[0] != numbers[1]:
        raise PassportError(f'Нельзя однозначно связать ВИК без года в имени с паспортом {canonical}: {path}.')
    return numbers[0] == canonical


def general_vik_template_path(settings, *, fallback_template=None, fallback_root=None):
    """One exact file; the legacy directory option also resolves only ВИК.docx."""
    if fallback_template is not None and fallback_root is not None:
        raise PassportError('Задайте либо точный fallback-template, либо fallback-root, но не оба.')
    if fallback_template is not None:
        return Path(fallback_template).resolve()
    return (Path(fallback_root) if fallback_root is not None else settings.documents_root.parent / 'Шаблоны').resolve() / 'ВИК.docx'


def find_vik_template(passport_root, project, fallback_template):
    """Canonical first passport only; fallback is a single explicit file."""
    canonical = re.fullmatch(r'([1-9]\d*)/(\d{4})', project.first_number)
    if not canonical:
        raise PassportError('Первый паспорт должен иметь однозначный номер число/год.')
    first_number, first_year = int(canonical[1]), canonical[2]
    candidates = []
    for path in Path(passport_root).rglob('*'):
        if not path.is_file() or path.name.startswith('~$') or not re.search(r'(?<!\w)ВИК(?!\w)', path.stem, re.I):
            continue
        relative = path.relative_to(passport_root)
        identities = [_passport_name_number(path.stem)]
        if len(relative.parts) > 1:
            folder_number = _passport_name_number(relative.parts[0])
            if folder_number is None:
                continue  # An unrelated directory is not a first-passport location.
            identities.extend(_passport_name_number(name) for name in relative.parts[:-1])
        explicit = [n for n in identities if n is not None]
        if any(n != first_number or (year is not None and year != first_year) for n, year in explicit):
            continue
        if not any(year is not None for _, year in explicit):
            if path.suffix.lower() != '.docx':
                raise PassportError(f'Неаудированный формат ВИК без номера/года: {path.suffix}.')
            if not _legacy_vik_number_matches(path, project.first_number):
                continue
        candidates.append(path)
    if len(candidates) > 1:
        raise PassportError(f'ВИК-шаблон не определён однозначно: {len(candidates)} кандидатов (project_first_passport).')
    if candidates:
        path, origin = candidates[0].resolve(), 'project_first_passport'
    else:
        path, origin = Path(fallback_template).resolve(), 'general_fallback'
        if not path.is_file():
            raise PassportError(f'Не найден ВИК первого паспорта {project.first_number}; отсутствует точный общий шаблон: {path}.')
    if path.suffix.lower() != '.docx':
        raise PassportError(f'Неаудированный формат ВИК: {path.suffix}. Требуется DOCX.')
    return path, origin


def _columns(registry):
    from .excel import load_book
    book = load_book(registry.path)
    try:
        sheet = book[registry.sheet]
        header = sheet[sheet.tables[registry.table].ref.split(':')[0]].row
        labels = {}
        for cell in sheet[header]:
            labels.setdefault(norm(cell.value), []).append(cell.column)
        if any(len(labels.get(norm(label), [])) != 1 for label in FIELDS.values()):
            raise PassportError('Отсутствует или повторяется обязательный заголовок ВИК в ВЗМК.xlsx.')
        return {key: labels[norm(label)][0] for key, label in FIELDS.items()}
    finally:
        book.close()


def _text(value, label):
    if not isinstance(value, str) or not value.strip() or any(c in value for c in '\r\n\t\x07{}'):
        raise PassportError(f'Пустое или небезопасное значение «{label}» в строке паспорта.')
    return value  # Preserve the registry's spelling, punctuation and spaces.


def read_requisites(excel, snapshot, registry, columns, first_row, current_row):
    book = open_excel(excel, snapshot)
    try:
        sheet = book.Worksheets(registry.sheet)
        cell = sheet.Cells(current_row, columns['vik_date'])
        if not cell.HasFormula:
            raise PassportError('«ВИК Дата» текущего паспорта должна содержать существующую формулу Excel.')
        formula = cell.FormulaR1C1
        excel.CalculateFullRebuild()
        if excel.CalculationState != 0 or not cell.HasFormula or cell.FormulaR1C1 != formula:
            raise PassportError('Пересчёт Excel не завершён или изменилась формула «ВИК Дата».')
        raw = cell.Value2
        if str(cell.Text).startswith('#') or isinstance(raw, bool) or not isinstance(raw, (float, int)) or not math.isfinite(raw) or raw <= 0:
            raise PassportError('«ВИК Дата» после пересчёта пустая или ошибочная.')
        try:
            calculated = date(1899, 12, 30) + timedelta(days=int(raw))
        except (ValueError, OverflowError) as exc:
            raise PassportError('«ВИК Дата» не является допустимой датой Excel.') from exc
        current, first = {}, {}
        for key, col in columns.items():
            if key in {'vik_date', 'end_date'}:
                continue
            current[key] = _text(sheet.Cells(current_row, col).Value2, FIELDS[key])
            first[key] = _text(sheet.Cells(first_row, col).Value2, FIELDS[key])
        current['vik_date'] = calculated.strftime('%d.%m.%Y')
        end_raw = sheet.Cells(current_row, columns['end_date']).Value2
        if isinstance(end_raw, bool) or not isinstance(end_raw, (int, float)) or not math.isfinite(end_raw) or end_raw <= 0:
            raise PassportError('Недопустимая «Дата окончания» текущего паспорта.')
        current['end_date'] = (date(1899, 12, 30) + timedelta(days=int(end_raw))).strftime('%d.%m.%Y')
        return first, current, {'sheet': registry.sheet, 'row': current_row,
                                'headers': {key: {'label': FIELDS[key], 'column': col} for key, col in columns.items()},
                                'formula_r1c1': formula, 'calculated_value2': raw, 'requisites': current}
    finally:
        book.Close(SaveChanges=False)  # Neither snapshot nor source registry is saved.


@contextmanager
def word_session():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        word.Options.UpdateLinksAtOpen = False
        yield word
    except PassportError:
        raise
    except Exception as exc:
        raise PassportError(f'Ошибка Microsoft Word; формирование остановлено: {exc}') from exc
    finally:
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def open_word(word, path, *, readonly):
    doc = word.Documents.Open(FileName=str(Path(path).resolve()), ReadOnly=readonly,
                              AddToRecentFiles=False, ConfirmConversions=False,
                              OpenAndRepair=False, NoEncodingDialog=True)
    if not readonly and doc.ReadOnly:
        doc.Close(False)
        raise PassportError('Word открыл временный ВИК только для чтения. Закройте занятый файл.')
    if doc.ProtectionType != -1 or doc.Revisions.Count or doc.TrackRevisions:
        doc.Close(False)
        raise PassportError('Защита или исправления Word: безопасная подстановка не подтверждена.')
    doc.Repaginate()
    return doc


def _paragraphs(doc):
    return [doc.Paragraphs(i).Range.Text for i in range(1, doc.Paragraphs.Count + 1)]


def _font(range_):
    props = ('Name', 'NameFarEast', 'NameBi', 'Size', 'SizeBi', 'Bold', 'BoldBi', 'Italic',
             'ItalicBi', 'Underline', 'Color', 'Hidden', 'StrikeThrough', 'Superscript',
             'Subscript', 'Spacing', 'Scaling', 'Position')
    values = {p: getattr(range_.Font, p) for p in props}
    if any(v == 9999999 for v in values.values()):
        raise PassportError('Изменяемое поле имеет смешанное оформление. Нужен отдельный аудит.')
    return values


def _locate(paragraph, old):
    r = paragraph.Range.Duplicate
    finder = r.Find
    finder.ClearFormatting()
    finder.Text = old
    finder.Forward = True
    finder.Wrap = 0
    finder.MatchCase = True
    finder.MatchWildcards = False
    finder.MatchWholeWord = False
    finder.Format = False
    if not finder.Execute() or r.Text != old:
        raise PassportError('Word не смог точно выделить изменяемое поле.')
    return r


def plan_slots(doc, first, current, origin):
    if doc.ComputeStatistics(2) != 2:
        raise PassportError('ВИК-шаблон должен состоять ровно из двух страниц.')
    controls = {}
    for i in range(1, doc.ContentControls.Count + 1):
        c = doc.ContentControls(i)
        controls.setdefault(c.Tag, []).append(c)
    if any(len(controls.get(tag, [])) != count for tag, (_, count) in TAGS.items()):
        raise PassportError('Набор именованных полей ВИК не соответствует аудированной структуре.')
    slots = []
    before = _paragraphs(doc)
    expected = list(before)

    def add(range_, key, expected_page, *, control=None):
        if range_.Information(3) != expected_page or range_.StoryType != 1:
            raise PassportError(f'Поле «{FIELDS[key]}» находится вне ожидаемой страницы/основного текста.')
        old = range_.Text
        if control is not None and (control.Type != 1 or control.LockContents or control.LockContentControl):
            raise PassportError(f'Поле {control.Tag} заблокировано или имеет неаудированный тип.')
        paragraph_index = next((i for i in range(1, doc.Paragraphs.Count + 1)
                                if doc.Paragraphs(i).Range.Start <= range_.Start < doc.Paragraphs(i).Range.End), None)
        if paragraph_index is None or expected[paragraph_index - 1].count(old) != 1:
            raise PassportError(f'Неоднозначный диапазон поля «{FIELDS[key]}».')
        expected[paragraph_index - 1] = expected[paragraph_index - 1].replace(old, current[key], 1)
        slots.append({'range': range_.Duplicate, 'key': key, 'old': old, 'new': current[key],
                      'page': expected_page, 'paragraph': paragraph_index, 'font': _font(range_),
                      'tag': control.Tag if control is not None else None})

    for tag, (key, count) in TAGS.items():
        items = sorted(controls[tag], key=lambda c: c.Range.Start)
        for index, c in enumerate(items):
            page = index + 1 if count == 2 else 1
            old = c.Range.Text
            if key == 'vik_date':
                valid = old == TOKENS[key] or (origin != 'general_fallback' and re.fullmatch(r'\d{2}\.\d{2}\.\d{4}', old))
            else:
                valid = old == TOKENS[key] or (origin != 'general_fallback' and old == first[key])
            if not valid:
                raise PassportError(f'Поле {tag} не совпадает с первым паспортом или placeholder.')
            add(c.Range, key, page, control=c)

    headers = [p.rstrip('\r\x07') for p in before if p.startswith('АКТ № ')]
    old_number = controls['PASSPORT_NO'][0].Range.Text
    old_date = controls['VIK_DATE'][0].Range.Text
    if headers != [f'АКТ № {old_number} от {old_date}', f'АКТ № {old_number}-1 от {old_date}']:
        raise PassportError('Неоднозначные заголовки двух актов ВИК.')
    if controls['PASSPORT_NO'][1].Range.Text != old_number or controls['VIK_DATE'][1].Range.Text != old_date:
        raise PassportError('Номер/дата двух актов шаблона различаются.')

    obj_old = TOKENS['object'] if origin == 'general_fallback' else first['object']
    code_old = TOKENS['code'] if origin == 'general_fallback' else first['code']
    con_old = TOKENS['connections'] if origin == 'general_fallback' else first['connections']
    object_line = f'{obj_old} ({code_old})'
    connections_line = f'НШ, ТШ, СШ соединения {con_old} МП в соответствии с требованиями технологии сварки '
    for exact, fields in [(object_line, [('object', obj_old), ('code', code_old)]),
                          (connections_line, [('connections', con_old)])]:
        matches = [i for i, text in enumerate(before, 1) if text.rstrip('\r\x07') == exact]
        if len(matches) != 1:
            raise PassportError('Обычный текст второго акта не совпадает с аудированными границами поля.')
        paragraph = doc.Paragraphs(matches[0])
        for key, old in fields:
            add(_locate(paragraph, old), key, 2)
    for a, b in zip(sorted(slots, key=lambda s: s['range'].Start), sorted(slots, key=lambda s: s['range'].Start)[1:]):
        if a['range'].End > b['range'].Start:
            raise PassportError('Перекрывающиеся изменяемые диапазоны ВИК.')
    return slots, expected


def _semantic(element):
    if element is None:
        return None
    attrs = tuple(sorted((k, v) for k, v in element.attrib.items() if 'rsid' not in k))
    return element.tag, attrs, element.text, tuple(_semantic(c) for c in element)


def package_layout(path):
    """Exact layout-property and opaque-part comparison after native Word save."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
        properties = {tag: [_semantic(x) for x in root.iter(W + tag)]
                      for tag in ('pPr', 'tblPr', 'tblGrid', 'trPr', 'tcPr', 'sectPr', 'sdtPr')}
        # Word inserts empty end-property elements when saving the general
        # template. Absent and empty mean no end formatting. Keep one entry per
        # control so any actual end formatting or relocation still fails.
        properties['sdtEndPr'] = []
        for control in root.iter(W + 'sdt'):
            end = control.find(W + 'sdtEndPr')
            empty = end is None or (not end.attrib and not len(end) and not (end.text or '').strip())
            properties['sdtEndPr'].append(None if empty else _semantic(end))
        opaque = {n: z.read(n) for n in z.namelist()
                  if n.startswith(('word/media/', 'word/header', 'word/footer', 'word/embeddings/'))}
        return properties, opaque


def verify_document(doc, expected, baseline, layout, path, slots):
    doc.Repaginate()
    if doc.ComputeStatistics(2) != 2 or doc.Tables.Count != baseline['tables'] or doc.Sections.Count != baseline['sections']:
        raise PassportError('Нарушена двухстраничная структура ВИК.')
    if _paragraphs(doc) != expected:
        raise PassportError('Изменился постоянный текст или не выполнена подстановка ВИК.')
    if any(token in doc.Content.Text for token in TOKENS.values()):
        raise PassportError('Остались неподставленные placeholders ВИК.')
    if package_layout(path) != layout:
        raise PassportError('Изменилось оформление таблиц, абзацев, секций или постоянных частей DOCX.')
    for slot in slots:
        p = doc.Paragraphs(slot['paragraph'])
        r = _locate(p, slot['new'])
        if r.Information(3) != slot['page'] or _font(r) != slot['font']:
            raise PassportError('Изменилось оформление или страница динамического поля ВИК.')


def generate_vik(settings, project_name, number, *, test_root, fallback_root=None, fallback_template=None):
    root = validation_boundary(settings, test_root)
    if not re.fullmatch(r'[1-9]\d*/\d{4}', number):
        raise PassportError('Номер паспорта должен иметь формат число/год.')
    ensure_no_pending(settings.registry_path)
    registry = read_registry(settings.registry_path)
    project = select_project(registry, project_name)
    if number == project.first_number:
        raise PassportError('Первый паспорт является шаблоном и не может быть текущим output.')
    rows = [r for r in registry.projects[project_name] if r['number'] == number]
    if len(rows) != 1:
        raise PassportError('Текущий паспорт не найден однозначно в выбранном проекте ВЗМК.xlsx.')
    passport_root, _ = settings.project_paths(project_name)
    if not passport_root.is_relative_to(root):
        raise PassportError('Папка паспортов выходит за validation-копию.')
    template, origin = find_vik_template(passport_root, project,
        general_vik_template_path(settings, fallback_template=fallback_template, fallback_root=fallback_root))
    columns = _columns(registry)
    stage = root / '.vik-stage' / uuid.uuid4().hex
    published = False
    temporary = None
    try:
        with ExitStack() as stack:
            registry_bytes = stack.enter_context(locked_file(registry.path))
            template_bytes = stack.enter_context(locked_file(template))
            if hashlib.sha256(registry_bytes).hexdigest() != registry.digest:
                raise PassportError('Реестр изменился во время проверки.')
            template_digest = hashlib.sha256(template_bytes).hexdigest()
            stage.mkdir(parents=True)
            snapshot = stage / 'registry.xlsx'
            snapshot.write_bytes(registry_bytes)
            with excel_session() as excel:
                first, current, evidence = read_requisites(excel, snapshot, registry, columns, project.first_row, rows[0]['row'])
            if current['number'] != number:
                raise PassportError('Пересчитанная строка не соответствует текущему паспорту.')
            folder = passport_root / f'{number.replace("/", "-")} от {current["end_date"]}'
            if not folder.is_dir() or folder.is_symlink() or not folder.resolve().is_relative_to(root):
                raise PassportError('Папка текущего паспорта должна уже существовать после Этапа 1 внутри validation-копии.')
            output = folder / f'{number.replace("/", "_")} - ВИК.docx'
            if output.exists() or (folder / ('~$' + output.name)).exists():
                raise PassportError('ВИК уже существует или открыт в Word. Перезапись запрещена.')
            temporary = folder / ('.vik-pending-' + uuid.uuid4().hex + '.docx')
            with temporary.open('xb') as f:
                f.write(template_bytes)
                f.flush()
                os.fsync(f.fileno())
            layout = package_layout(temporary)
            with word_session() as word:
                doc = open_word(word, temporary, readonly=False)
                try:
                    slots, expected = plan_slots(doc, first, current, origin)
                    baseline = {'tables': doc.Tables.Count, 'sections': doc.Sections.Count}
                    for slot in sorted(slots, key=lambda s: s['range'].Start, reverse=True):
                        if slot['old'] != slot['new']:
                            slot['range'].Text = slot['new']
                    doc.Repaginate()
                    doc.Save()
                finally:
                    doc.Close(False)
                # Normal reopen; recovery is explicitly disabled.
                doc = open_word(word, temporary, readonly=True)
                try:
                    verify_document(doc, expected, baseline, layout, temporary, slots)
                    doc.ActiveWindow.View.Type = 3
                    for i in (1, 2):
                        (stage / f'output-page-{i}.emf').write_bytes(bytes(doc.ActiveWindow.Panes(1).Pages(i).EnhMetaFileBits))
                finally:
                    doc.Close(False)
            evidence.update(template=str(template), template_origin=origin, template_sha256=template_digest,
                            output=str(output), output_sha256=sha256(temporary), pages=2,
                            tables=baseline['tables'], sections=baseline['sections'],
                            constant_text_preserved=True, layout_properties_preserved=True,
                            normal_word_reopen=True, registry_sha256=registry.digest,
                            slots=[{k: v for k, v in s.items() if k != 'range'} for s in slots])
            evidence_path = stage / 'generation.json'
            evidence_path.write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding='utf-8')
            # Windows rename fails if the target appears concurrently. No overwrite.
            temporary.rename(output)
            published = True
            return {'output': str(output), 'evidence': str(evidence_path), 'stage': str(stage)}
    except PassportError:
        raise
    except Exception as exc:
        raise PassportError(f'ВИК не опубликован: файл занят, недоступен или не прошёл проверку: {exc}') from exc
    finally:
        # Only this invocation's private temporary DOCX can be removed.
        if not published and temporary is not None and temporary.exists():
            try:
                temporary.unlink()
            except OSError:
                pass  # It retains the pending name and is never reported as success.
