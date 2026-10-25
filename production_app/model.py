import re
from dataclasses import replace
from datetime import date,datetime
from pathlib import Path
from decimal import Decimal
from passport_app.registry import read_registry,Project
from passport_app.excel import load_book
from passport_app.vom import read_vom
from .storage import Stop,safe,contained,sha,save,local_root

HEADERS={'NUMBER':'№ паспорта','CONTRACT':'Договор №','CUSTOMER':'Заказчик','OBJECT':'Наименование объекта','CIPHER':'шифр КМД','MASS':'Масса кг',
    'START_DATE':'Дата начала','END_DATE':'Дата окончания','KM':'КМ Организация и шифр','KMD_COMPANY':'КМД Организация','STEEL':'Марка стали',
    'VIK_DATE':'ВИК Дата','UZK_DATE':'УЗК Дата','AKZ_DATE':'Дата Акз','CONNECTIONS':'Соединения'}

def columns(registry):
    from passport_app.tables import norm
    w=load_book(registry.path)
    try:
        s=w[registry.sheet];table=s.tables[registry.table];row=s[table.ref.split(':')[0]].row
        labels={}
        for cell in s[row]:labels.setdefault(norm(cell.value),[]).append(cell.column)
        if any(len(labels.get(norm(v),[]))!=1 for v in HEADERS.values()):raise Stop('Неоднозначная/неполная структура рабочих реквизитов реестра.')
        return {k:labels[norm(v)][0] for k,v in HEADERS.items()}
    finally:w.close()

def parse_number(value):
    m=re.fullmatch(r'([1-9]\d*)/(?:([1-9]\d*)/)?(20\d{2})',str(value))
    if not m:raise Stop('Номер должен иметь формат N/YYYY или N/S/YYYY: '+str(value))
    return int(m[1]),m[2],m[3]

def number(n,series,year):return f'{n}/'+(series+'/' if series else '')+year
def filename(value):parse_number(value);return value.replace('/','_')

# User-approved READ-ONLY interpretation of exactly two historical source cells.
# Never rewrite source values, generated document numbers or existing filenames.
TEMPO1_LEGACY_ALIASES = {
    (401, '10/2026'): '10/1/2026',
    (407, '11/2026'): '11/1/2026',
}

def approved_legacy_aliases(name, entries):
    if name != 'ТЭМПО 1':
        return []
    return [(e, TEMPO1_LEGACY_ALIASES[(e['row'], str(e['number']))])
            for e in entries if (e['row'], str(e['number'])) in TEMPO1_LEGACY_ALIASES]

def select(registry,name):
    if any(c in name for c in '<>:"/\\|?*') or name in ('.','..') or name.endswith((' ','.')):raise Stop('Имя объекта небезопасно для папки результата.')
    entries=registry.projects.get(name,[])
    if not entries:raise Stop('Объект не найден в реестре: '+name)
    aliases=approved_legacy_aliases(name,entries)
    if aliases:
        # Fail closed if the documented first passport or either historical
        # source row is missing, moved, or contains an unexpected number.
        by_row={e['row']:str(e['number']) for e in entries}
        if (by_row.get(342)!='1/1/2026'
            or any(by_row.get(row) not in (raw,normalized)
                   for (row,raw),normalized in TEMPO1_LEGACY_ALIASES.items())):
            raise Stop('ТЭМПО 1: подтверждённые опечатки номеров не совпали со строками 401/407. Реестр не изменён.')
    normalized={(e['row'],str(e['number'])):alias for e,alias in aliases}
    parsed=[(parse_number(normalized.get((e['row'],str(e['number'])),e['number'])),e) for e in entries]
    if len({str(e['number']) for e in entries})!=len(entries):raise Stop('Дубли номеров у объекта '+name)
    if len({number(*p) for p,e in parsed})!=len(parsed):
        raise Stop('Дубли номеров с учётом подтверждённых опечаток у объекта '+name)
    groups={(p[1],p[2]) for p,e in parsed}
    if len(groups)!=1:
        detail='; '.join(f"строка {e['row']}: {e['number']}" for _,e in parsed)
        raise Stop('КОНФЛИКТ СЕРИИ/ГОДА '+name+': '+detail+'. Исходный реестр не исправлен.')
    series,year=next(iter(groups))
    first=[e for p,e in parsed if p[0]==1]
    if len(first)!=1:raise Stop('Нет единственного первого паспорта '+name)
    start=first[0];last=max(parsed,key=lambda t:t[0][0])[1]
    if not isinstance(start['date'],(date,datetime)):raise Stop('Не заполнена дата первого паспорта.')
    # Contract/object/cipher must remain identical within the selected series.
    cols=columns(registry);w=load_book(registry.path,values=True)
    try:
        s=w[registry.sheet]
        def requisites(r):
            values=[s.cell(r,cols[k]).value for k in ('CONTRACT','CUSTOMER','OBJECT','CIPHER')]
            # An optional printed number sign is not a different contract.
            # Comparison only; never rewrite the source text.
            values[0]=re.sub(r'^\s*№\s*','',str(values[0]))
            return values
        baseline=requisites(start['row'])
        if aliases and (baseline[0]!='МК-3465 от 05.05.2026'
                        or baseline[3]!='SKL-VD-(L3-41-030)-N-ZX 2511.3.135.021'):
            raise Stop('ТЭМПО 1: опечатки нельзя учитывать — договор/шифр первого паспорта не совпали с подтверждённым объектом.')
        for _,e in parsed:
            if requisites(e['row'])!=baseline:
                raise Stop(f"КОНФЛИКТ РЕКВИЗИТОВ: {name}, строка {e['row']}; договор/заказчик/объект/шифр отличаются от первого паспорта.")
    finally:w.close()
    day=start['date'].date() if isinstance(start['date'],datetime) else start['date']
    nxt=number(max(p[0] for p,e in parsed)+1,series,year)
    return Project(name,start['row'],start['number'],day,str(start['object']),last['row'],last['number'],nxt)

def validate_paths(cfg):
    registry=safe(cfg['registry']);documents=safe(cfg['documents']);output=safe(cfg['output'])
    if not registry.is_file() or registry.suffix.lower()!='.xlsx':raise Stop('Выберите существующий реестр .xlsx.')
    if not documents.is_dir() or not output.is_dir():raise Stop('Выберите существующие папки документов и сохранения.')
    if cfg.get('first_passport_documents'):
        first=safe(cfg['first_passport_documents'])
        if not first.is_dir():raise Stop('Выберите существующую папку архива первых паспортов.')
    return registry,documents,output

def profile(cfg,name):
    p=cfg.get('profiles',{}).get(name)
    if not p:
        if not cfg.get('auto_first_passport',False):
            raise Stop('Не подключены формы проекта: '+name)
        from .prepared_forms import prepared_project
        p=prepared_project(cfg,name)
    root=contained(p['root'],cfg['documents'])
    if not root.is_dir():raise Stop('Папка объекта отсутствует.')
    templates=p.get('templates',{})
    required={'vik','uzk','quality'} if cfg.get('auto_first_passport',False) else {'appendix','vik','uzk','quality'}
    if not required.issubset(templates):
        raise Stop('Не хватает подготовленных безопасных Word-форм: '+
                   ', '.join(sorted(required-set(templates)))+
                   '. Заполненные акты первого паспорта нельзя выпускать '
                   'повторно без очистки старых результатов.')
    for role,ref in templates.items():
        # The actual first Appendix will be validated/read separately below.
        # An old manually cached appendix must NOT become a prerequisite.
        if role=='appendix' and cfg.get('auto_first_passport',False):
            continue
        path=contained(ref['path'],root)
        if not path.is_file() or sha(path)!=ref['sha256']:
            raise Stop('Шаблон изменён/отсутствует: '+str(path))
    if p.get('catalog') and not cfg.get('auto_first_passport',False):
        catalog=contained(p['catalog'],root)
        if not catalog.is_file() or sha(catalog)!=p.get('catalog_sha256'):
            raise Stop('Изменён/отсутствует справочник марок объекта.')
    # The approved EMPTY dynamic Word/AKZ forms remain pinned above. When
    # the user enables automatic first-passport sources, the FIRST Appendix
    # is taken directly from the project's historical issued workbook, which
    # the established dynamic appendix engine can already safely process.
    # Do NOT bind filled VIK/UZK/quality/AKZ acts as empty forms.
    if cfg.get('auto_first_passport',False):
        from .first_passport import discover_first_passport_for_config
        project=select(read_registry(cfg['registry']),name)
        inventory=discover_first_passport_for_config(cfg,project)
        original=inventory['sources'].get('appendix')
        if original is None:
            raise Stop('Первый паспорт найден, но первое Приложение №1 '
                       'отсутствует или не подтверждено. Не использован '
                       'чужой подготовленный шаблон.')
        result=dict(p)
        result['templates']=dict(templates)
        result['templates']['appendix']=dict(original)
        result['first_passport_sources']=inventory
        if inventory.get('catalog'):
            # Engineering assignments come from the selected REAL project
            # folder, not from the installed application's Templates root.
            result['catalog']=inventory['catalog']['path']
            result['catalog_sha256']=inventory['catalog']['sha256']
            result['catalog_origin']='SELECTED_PROJECT_FIRST_PASSPORT'
        elif p.get('catalog'):
            # A user-pinned manual catalogue remains possible for projects
            # without a standard project-local file.
            manual=Path(p['catalog'])
            if not manual.is_file() or sha(manual)!=p.get('catalog_sha256'):
                raise Stop('Изменён вручную выбранный справочник марок.')
            result['catalog']=str(manual)
            result['catalog_sha256']=p['catalog_sha256']
            result['catalog_origin']='MANUAL_PINNED'
        return result
    return p

def output_passport_directory(cfg,project,profile=None):
    """Exact existing <Работа/Документы>/<PROJECT>/Паспорт folder.

    New real passports MUST be siblings of the first passport, including
    case-preserved names (e.g. 'Гараж казахстан'). For a sandbox write under
    the isolated training output root as before. Never write in the archive
    when training=True.
    """
    if cfg.get('training'):
        return safe(cfg['output'])/project/'Паспорт'
    output=safe(cfg['output'])
    inventory=(profile or {}).get('first_passport_sources')
    if inventory:
        actual=contained(inventory['passport_root'],output)
        if actual.parent.parent!=output or not actual.is_dir():
            raise Stop('Папка первого паспорта не является папкой '
                       'Работа/Документы/<проект>/Паспорт: '+str(actual))
        if actual.name.casefold()!='паспорт':
            raise Stop('Неверное имя папки паспортов.')
        if actual.parent.name.casefold()!=project.casefold():
            raise Stop('Папка первого паспорта принадлежит другому проекту.')
        return actual
    # Compatibility when automatic first-source mode is not enabled.
    return output/project/'Паспорт'


def visible_vom(path):
    # Native parser: remove only hidden/filter-excluded rows in the private copy.
    # Source is read without writing. Existing parser validates every numeric cell.
    path=safe(path)
    if path.suffix.lower()=='.xlsm':
        # The existing XLSX parser reads an immutable byte copy. No macros,
        # external refresh, saving or conversion through Office are involved.
        import tempfile,hashlib
        with tempfile.TemporaryDirectory(prefix='vzmk-vom-') as temp:
            payload=path.read_bytes();h=hashlib.sha256(payload).hexdigest()
            snapshot=Path(temp)/'read_only.xlsx';snapshot.write_bytes(payload)
            v=visible_vom(snapshot)
            if sha(path)!=h:raise Stop('ВОМ изменился во время чтения.')
            return replace(v,path=path,digest=h)
    v=read_vom(path);w=load_book(path)
    try:hidden={r for r in v.source_rows if w[v.layout.sheet].row_dimensions[r].hidden}
    finally:w.close()
    rows=tuple(row for r,row in zip(v.source_rows,v.rows) if r not in hidden)
    if not rows:raise Stop('В ВОМ нет видимых марок.')
    names=list(dict.fromkeys(str(r[1]).strip() for r in rows if r[1]))
    return replace(v,rows=rows,source_rows=tuple(r for r in v.source_rows if r not in hidden),
        quantity=sum((r[2] for r in rows),Decimal(0)),mass=sum((r[4] for r in rows),Decimal(0)),
        area=sum((r[6] for r in rows),Decimal(0)),connections=', '.join(names))

def draft_catalog(vom,target):
    # Explicitly requested standalone draft. No profile, paint or AKZ defaults.
    target=safe(target)
    if target.exists():raise Stop('Файл уже существует. Перезапись запрещена.')
    from passport_app.excel import excel_session
    with excel_session() as app:
        book=app.Workbooks.Add(-4167)
        try:
            s=book.Worksheets(1);s.Name='Черновик марок'
            s.Cells(1,1).Value2='ЧЕРНОВИК — назначения не подтверждены'
            s.Range('A2:F2').Value2=(('Марка','Наименование','Количество','Профиль','Краска','№ АКЗ'),)
            for n,row in enumerate(visible_vom(vom).rows,3):s.Range(f'A{n}:F{n}').Value2=((str(row[0]),row[1],float(row[2]),None,None,None),)
            s.Columns('A:F').AutoFit();s.Rows(2).Font.Bold=True
            book.SaveAs(str(target),FileFormat=51)
        finally:book.Close(SaveChanges=False)
    return str(target)
