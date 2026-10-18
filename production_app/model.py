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
    return registry,documents,output

def profile(cfg,name):
    p=cfg.get('profiles',{}).get(name)
    if not p:raise Stop('Настройте собственные шаблоны объекта: '+name)
    root=contained(p['root'],cfg['documents'])
    if not root.is_dir():raise Stop('Папка объекта отсутствует.')
    templates=p.get('templates',{})
    required={'appendix','vik','uzk','quality'}
    if not required.issubset(templates):raise Stop('Не привязаны шаблоны Приложение №1, ВИК, УЗК, качество.')
    for role,ref in templates.items():
        path=contained(ref['path'],root)
        if not path.is_file() or sha(path)!=ref['sha256']:raise Stop('Шаблон изменён/отсутствует: '+str(path))
    if p.get('catalog'):
        catalog=contained(p['catalog'],root)
        if not catalog.is_file() or sha(catalog)!=p.get('catalog_sha256'):raise Stop('Изменён/отсутствует справочник марок объекта.')
    return p

def visible_vom(path):
    # Native parser: remove only hidden/filter-excluded rows in the private copy.
    # Source is read without writing. Existing parser validates every numeric cell.
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
