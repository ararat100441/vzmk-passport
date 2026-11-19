"""Read-only, project-specific discovery of the FIRST issued passport.

This is a source locator and provenance checker, NOT a license to reuse the
first passport's old inspection findings, signatures, dates or paint batches.
The first Appendix #1 is a safe source for its existing, separately verified
dynamic table generator. Other issued forms must be converted into EMPTY
approved forms before passing them into the production renderer.

No changes to registry, historical Word/Excel, filesystem or settings.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree as ET

from passport_app.appendix import inspect_template
from passport_app.excel import load_book
from .storage import Stop, sha, safe


_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
_FIRST_ROLES = {
    'appendix': ('Приложение № 1.xlsx', 'Приложение №1.xlsx','appendix.xlsx'),
    'vik': ('ВИК.docx','vik.docx'),
    'uzk': ('УЗК.docx','uzk.docx'),
    'quality': ('Документ о качестве.docx','quality.docx'),
    'akz_appendix0': ('Приложение к АКЗ.xlsx','akz_appendix0.xlsx'),
}


def _key(s):
    return ' '.join(str(s).replace('ё', 'е').casefold().split())


def _safe_child(parent, item):
    """Reject redirects/hard links even when their target is inside the tree."""
    if item.is_symlink() or (getattr(item,'is_junction',lambda:False)()):
        raise Stop('Папка или файл первого паспорта является ссылкой: '+str(item))
    if item.is_file() and item.stat().st_nlink != 1:
        raise Stop('Hardlink недопустим для исходного файла: '+str(item))
    candidate=item.resolve()
    if not candidate.is_relative_to(parent.resolve()):
        raise Stop('Источник первого паспорта выходит за пределы папки проекта.')
    return candidate


def _unique_subdirectory(root, title, *, optional=False):
    matches=[p for p in root.iterdir()
             if p.is_dir() and _key(p.name)==_key(title)
             and not p.name.startswith('.')]
    if len(matches)!=1:
        if optional and not matches:
            return None
        raise Stop('Не найдена единственная папка «'+title+
                   '»; найдено '+str(len(matches))+'.')
    return _safe_child(root,matches[0])


def _first_roots(passport_root, project):
    """First passport documents may be at root OR inside its dated folder.

    Accept only exact first-number/date spellings, e.g.
    1_2026 от 09.09.2026, 1 от 09.09.2026, 1_1_2026 от 12.06.2026.
    Never treat passport 11/2026 or a different series as number 1.
    """
    roots=[passport_root]
    stamp=project.first_date.strftime('%d.%m.%Y')
    number=str(project.first_number)
    token=number.replace('/','_')
    n=number.split('/')[0]
    candidates={
        token,
        token+' от '+stamp,
        token+' - '+stamp,
        n+' от '+stamp,
        n+' - '+stamp,
    }
    desired={_key(c) for c in candidates}
    for item in passport_root.iterdir():
        if item.is_dir() and _key(item.name) in desired:
            roots.append(_safe_child(passport_root,item))
    return roots


def _exact_file(roots,basename):
    matches=[]
    for root in roots:
        for item in root.iterdir():
            if (item.is_file() and not item.name.startswith('~$')
                    and _key(item.name)==_key(basename)):
                matches.append(_safe_child(root,item))
    if len(matches)>1:
        raise Stop('Несколько документов первого паспорта с именем '+
                   basename+'; нельзя выбирать автоматически.')
    return matches[0] if matches else None


def _read_word_identity(path,first_number):
    """Reject a corrupt/non-first issued DOCX with stale different passport."""
    try:
        with ZipFile(path) as z:
            data=z.read('word/document.xml')
        root=ET.fromstring(data)
        text=''.join((node.text or '') for node in root.iter(_W+'t'))
    except (BadZipFile, KeyError, ET.ParseError, OSError) as exc:
        raise Stop('Повреждён исторический Word-файл: '+str(path)) from exc
    if not text.strip():
        raise Stop('Пустая форма первого паспорта: '+str(path))
    # Some historic Word fields split the number between XML runs. The exact
    # file identity plus the selected project's isolated Passport folder
    # are primary locators; a clearly different 'passport № N/...' blocks.
    numbers=re.findall(r'(?:№\s*|паспорт\s*)([0-9]+(?:/[0-9]+){1,2})',
                       text,flags=re.I)
    if numbers and first_number not in numbers:
        raise Stop('Word содержит другой номер паспорта: '+str(path))
    return True


def _read_akz_identity(path,first_number):
    book=load_book(path,values=True)
    try:
        if len(book.worksheets)!=1:
            raise Stop('В первом акте АКЗ найдено несколько листов.')
        s=book.active
        # The issued garage act is historically named 'Акт № 4-2026.xlsx',
        # while E4 contains the ACTUAL first passport '1/2026 от'.
        number=str(s['E4'].value or '')
        if not number.startswith(first_number+' от'):
            raise Stop('Первый акт АКЗ содержит другой номер: '+str(path))
    finally:
        book.close()
    return True


def discover_first_passport(documents_root,project,*,require_all=False,catalog_path=None):
    """Find first source files strictly inside ONE selected project folder.

    Returns a read-only inventory with pinned hashes. No templates or
    register values are rewritten; historical findings are never approved.
    """
    documents=safe(documents_root).resolve(strict=True)
    if not documents.is_dir():
        raise Stop('Не выбрана папка рабочих документов.')
    project_root=_unique_subdirectory(documents,project.name)
    passport_root=_unique_subdirectory(project_root,'Паспорт')
    roots=_first_roots(passport_root,project)
    number=str(project.first_number)
    prefix=number.replace('/','_')
    entries={};errors={}
    for role,endings in _FIRST_ROLES.items():
        candidates=[]
        for suffix in endings:
            path=_exact_file(roots,prefix+' - '+suffix)
            if path is not None:
                candidates.append(path)
        if len(set(candidates))>1:
            raise Stop('Неоднозначный документ первого паспорта: '+role)
        if not candidates:
            continue
        path=candidates[0]
        before=sha(path)
        if role=='appendix':
            # Strong validation of first passport's number/date/object and
            # the 7-column layout already used by the production engine.
            inspect_template(path,project)
        elif path.suffix.lower()=='.docx':
            _read_word_identity(path,number)
        else:
            with ZipFile(path) as z:
                if 'xl/workbook.xml' not in z.namelist():
                    raise Stop('Не Excel XLSX: '+str(path))
        if sha(path)!=before:
            raise Stop('Исторический источник изменился при чтении.')
        entries[role]={'path':str(path),'sha256':before,'origin':'FIRST_ISSUED_READ_ONLY'}

    # The original Kazakhstan 1-layer AKZ act has historically misleading
    # filename, 'Акт № 4-2026.xlsx'. Only use it in this exact project and
    # only after validating that the INNER E4 number is 1/2026.
    if _key(project.name)==_key('Гараж Казахстан') and number.split('/')[0]=='1':
        special='Акт № 4-'+number.split('/')[-1]+'.xlsx'
        path=_exact_file(roots,special)
        if path is not None:
            original=sha(path)
            _read_akz_identity(path,number)
            if sha(path)!=original:
                raise Stop('Исторический акт АКЗ изменился при чтении.')
            entries['akz0']={
                'path':str(path),'sha256':original,
                'origin':'FIRST_ISSUED_READ_ONLY',
            }

    from .project_catalog import resolve_catalog
    catalog=resolve_catalog(project_root,project.name,catalog_path)
    # Acts can have historical filenames. Bind them by the exact INNER act
    # number, inside this passport's exact directories, and reject duplicates.
    for root in roots:
        for candidate in root.iterdir():
            if candidate.suffix.lower()!='.xlsx' or candidate.name.startswith('~$'):
                continue
            is_act=('акт' in _key(candidate.stem) and 'приложени' not in _key(candidate.stem) or bool(re.search(r' - akz[0-3]$',candidate.stem)))
            is_appendix=('приложени' in _key(candidate.stem) and 'акз' in _key(candidate.stem) or bool(re.search(r' - akz_appendix[0-3]$',candidate.stem)))
            if not (is_act or is_appendix):continue
            path=_safe_child(root,candidate);before=sha(path)
            book=load_book(path,values=True)
            try:
                value=str(book.active['E4' if is_act else 'A1'].value or '')
            finally:book.close()
            if is_appendix and 'к акту АКЗ №' not in value:continue
            pattern=(r'^'+re.escape(number)+r'(?P<suffix>-[1-3])?\s+от\s' if is_act
                     else r'№\s*'+re.escape(number)+r'(?P<suffix>-[1-3])?\s+от\s')
            found=re.search(pattern,value)
            if not found:continue
            variant=int(found['suffix'][1:]) if found['suffix'] else 0
            role=('akz' if is_act else 'akz_appendix')+str(variant)
            if role in errors:continue
            if role in entries and Path(entries[role]['path'])!=path:
                errors[role]='Неоднозначный исходный документ '+role+' паспорта '+number
                entries.pop(role,None);continue
            if sha(path)!=before:raise Stop('Изменился исходный документ '+str(path))
            entries[role]={'path':str(path),'sha256':before,'origin':'SELECTED_ISSUED_READ_ONLY'}
    # Catalog is NOT a result or evidence of control. Bind only as a
    # candidate for explicitly confirmed mark/profile/paint/AKZ assignments.
    missing=[r for r in ('appendix','vik','uzk','quality') if r not in entries]
    if require_all and missing:
        raise Stop('В первом паспорте '+project.name+
                   ' не найдены исходные формы: '+', '.join(missing)+
                   '. Другие проекты и общие бланки не подставлены.')
    return {
        'project':project.name,
        'first_number':number,
        'first_date':project.first_date.isoformat(),
        'project_root':str(project_root),
        'passport_root':str(passport_root),
        'sources':entries,
        'missing':missing,'errors':errors,
        'catalog':({'path':str(catalog),'sha256':sha(catalog)}
                   if catalog else None),
        'status':'READ_ONLY_FIRST_PASSPORT',
    }


def archive_root_for_project(path,project_name):
    """Recognize a selected project's folder/passport leaf, without searching.

    Older settings can hold <archive>/<project>/Паспорт/<dated passport>.
    Only that exact project's existing hierarchy may identify the archive;
    unrelated selected directories and other projects remain unchanged.
    """
    root=safe(path)
    if not root.is_dir():
        return root
    for folder in (root,*root.parents):
        if _key(folder.name)!=_key(project_name):
            continue
        passports=_unique_subdirectory(folder,'Паспорт',optional=True)
        if passports is not None and (root==folder or root.is_relative_to(passports)):
            return safe(folder.parent)
        return root
    return root


def discover_first_passport_for_config(cfg,project,*,require_all=False):
    """Separate historical archive from the clean-template directory.

    Installed RC3-interface has:
      cfg['documents'] = C:\\VZMK-Passport\\...\\Templates
      cfg['output']    = ...\\Работа\\Документы (actual first passports)
    Never attempt to read the first passports from the clean form directory.
    If the user explicitly sets 'first_passport_documents', use that alone.
    Otherwise use exactly one already user-selected root containing the
    project/Pasпорт folder. Zero/multiple viable roots => safe STOP.
    """
    # Source selection changes ONLY the document reference, never numbering
    # or the project's manufacturing-start row in the working registry.
    selected=(cfg.get('source_passports') or {}).get(project.name)
    if selected and selected!=project.first_number:
        from dataclasses import replace
        from passport_app.registry import read_registry
        registry=read_registry(cfg['registry'])
        matches=[e for e in registry.projects.get(project.name,[])
                 if str(e['number'])==selected]
        if len(matches)!=1:
            raise Stop('Не найден единственный исходный паспорт '+selected+' объекта '+project.name)
        entry=matches[0]
        day=entry['date'].date() if isinstance(entry['date'],datetime) else entry['date']
        project=replace(project,first_number=selected,first_row=entry['row'],first_date=day)
    configured=str(cfg.get('first_passport_documents') or '').strip()
    if configured:
        return discover_first_passport(
            archive_root_for_project(configured,project.name),project,
            require_all=require_all,catalog_path=cfg.get('catalog_choices',{}).get(project.name),
        )
    roots=[]
    for kind in ('output','documents'):
        path=cfg.get(kind)
        if not path:
            continue
        root=safe(path)
        if not root.is_dir():
            continue
        if root not in roots:
            roots.append(root)
    found=[]
    for root in roots:
        candidates=[p for p in root.iterdir()
                    if p.is_dir() and _key(p.name)==_key(project.name)]
        if len(candidates)>1:
            raise Stop('Несколько папок объекта в '+str(root)+
                       '. Выберите точную папку первых паспортов в Настройках.')
        if not candidates:
            continue
        project_folder=_safe_child(root,candidates[0])
        passports=[p for p in project_folder.iterdir()
                   if p.is_dir() and _key(p.name)==_key('Паспорт')]
        if len(passports)>1:
            raise Stop('Несколько папок «Паспорт» у проекта '+project.name)
        if passports:
            found.append(root)
    if len(found)!=1:
        raise Stop('Не найдена единственная папка первых паспортов для '+
                   project.name+'. Найдено корневых папок: '+str(len(found))+
                   '. В «Настройки путей» выберите «Архив первых паспортов»: '
                   'папку Работа/Документы, где находятся паспорта проекта. '
                   'Каталог очищенных Templates не является архивом.')
    return discover_first_passport(
        found[0],project,require_all=require_all,catalog_path=cfg.get('catalog_choices',{}).get(project.name),
    )


def validate_source_pins(inventory):
    """Detect deleted or modified source BEFORE using it as a reference."""
    for role,record in inventory['sources'].items():
        path=Path(record['path'])
        if not path.is_file() or sha(path)!=record['sha256']:
            raise Stop('Первый паспорт изменился: '+role+' ('+str(path)+')')
    return True
