"""Read passport dates from the VZMK registry, with project-specific provenance.

For EVERY passport of a project, manufacturing START_DATE is the SAME:
copy the exact date from the FIRST issued passport's registered row (H353 in
Garage Kazakhstan), even when this is a fixed Excel date, not a formula.

For a NEW passport the ending date is proposed from the current date and
shown at confirmation. VIK/UZK/AKZ dates must come from live formulas in
a disposable registry copy, never stale control dates of the first passport.
For existing passports use their exact registered row.

The source VZMK registry, VOM and issued documents remain read-only.
"""
from datetime import date,datetime
from pathlib import Path
from shutil import copyfile
from tempfile import TemporaryDirectory

from passport_app.excel import load_book, excel_session
from passport_app.registry import read_registry, append_row
from . import model
from .storage import Stop, sha

MAP={'start_date':'START_DATE','date':'END_DATE','vik_date':'VIK_DATE',
     'uzk_date':'UZK_DATE','akz_date':'AKZ_DATE'}
CURRENT_CONTROL=('vik_date','uzk_date','akz_date')


def _coerce_date(value,label):
    if isinstance(value,datetime):value=value.date()
    if type(value) is not date:
        raise Stop('В реестре нет даты '+label+'; не копируем даты первого '
                   'паспорта и не создаём их автоматически.')
    return value.isoformat()


def row_dates(workbook,registry,row):
    col=model.columns(registry)
    sheet=workbook[registry.sheet]
    return {key:_coerce_date(sheet.cell(row,col[column]).value,key)
            for key,column in MAP.items()}


def first_passport_start(registry,project):
    """Read manufacturing start from the FIRST registered passport, never
    the next shipment's date and never from another project's row.
    A fixed date is expected and VALID under the approved VZMK procedure.
    """
    book=load_book(registry.path,values=True)
    try:
        columns=model.columns(registry)
        value=book[registry.sheet].cell(
            project.first_row,columns['START_DATE']).value
        return _coerce_date(value,'начала изготовления первого паспорта')
    finally:
        book.close()


def resolve_registry_dates(cfg,project_name,passport_number,vom,new,*,day=None):
    path=Path(cfg['registry']).resolve(strict=True)
    baseline=sha(path)
    registry=read_registry(path)
    project=model.select(registry,project_name)
    if new and passport_number!=project.next_number:
        raise Stop('Изменился следующий номер паспорта.')
    if new:
        day=day or date.today()
        if type(day) is not date:
            raise Stop('Неверная дата нового паспорта.')
        # The project manufacturing start is intentionally constant across
        # all passports; other inspection dates must still be NEW.
        approved_start=first_passport_start(registry,project)
        formula_book=load_book(path)
        try:
            columns=model.columns(registry)
            original=formula_book[registry.sheet]
            for key in CURRENT_CONTROL:
                cell=original.cell(project.first_row,columns[MAP[key]])
                if cell.data_type!='f':
                    raise Stop('В первом паспорте колонка '+MAP[key]+
                               ' содержит фиксированную дату старого контроля. '
                               'Для НОВОЙ отгрузки нужна подтверждённая '
                               'формула или иной согласованный источник даты '
                               'текущего контроля. Дата начала изготовления '
                               'при этом берётся из первого паспорта.')
        finally:
            formula_book.close()

        with TemporaryDirectory(prefix='vzmk-registry-dates-') as temp:
            temporary=Path(temp)/'registry-copy.xlsx'
            copyfile(path,temporary)
            copied=read_registry(temporary)
            selected=model.select(copied,project_name)
            with excel_session() as excel:
                append_row(excel,temporary,copied,selected,vom,day)
            actual=read_registry(temporary)
            rows=[e['row'] for e in actual.projects[project_name]
                  if str(e['number'])==passport_number]
            if len(rows)!=1:
                raise Stop('Не удалось проверить новую строку только в копии реестра.')
            book=load_book(temporary,values=True)
            try:
                dates=row_dates(book,actual,rows[0])
            finally:book.close()
        if dates['date']!=day.isoformat():
            raise Stop('Дата новой строки не совпадает с предложенной датой выпуска.')
        if dates['start_date']!=approved_start:
            raise Stop('Дата начала изготовления нового паспорта не совпала '
                       'с датой первого паспорта проекта. Источники не '
                       'изменены, выпуск остановлен.')
        origin='NEW_ROW_CALCULATED_IN_DISPOSABLE_REGISTRY_COPY'
    else:
        candidates=[e['row'] for e in registry.projects.get(project_name,[])
                    if str(e['number'])==passport_number]
        if len(candidates)!=1:
            raise Stop('Не найден один выбранный паспорт в реестре.')
        book=load_book(path,values=True)
        try:dates=row_dates(book,registry,candidates[0])
        finally:book.close()
        origin='EXACT_EXISTING_REGISTRY_ROW'
    if sha(path)!=baseline:
        raise Stop('Реестр изменился во время чтения дат.')
    return dates|{
        'registry_date_origin':origin,
        'start_date_origin':(
            'FIRST_PASSPORT_REGISTERED_START_DATE' if new
            else 'EXACT_EXISTING_REGISTRY_ROW'
        ),
    }
