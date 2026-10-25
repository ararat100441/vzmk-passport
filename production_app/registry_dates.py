"""Read current passport dates from the VZMK registry, never from old acts.

For an existing passport, read the exact registered row. For a new passport,
calculate the new row in a disposable *private* Excel copy using the existing
append_row engine; the passport's ending date is today's proposal and must
be shown to the user for explicit review. Every other current-work date MUST
come from a formula in the registered first-row model. Stale literal dates
of the first passport are not valid inspection dates for a new shipment.

Actual source registry and VOM remain read-only; no official row is inserted.
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
CONTROL=('start_date','vik_date','uzk_date','akz_date')


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
        # A copied historical literal is not a current date. Explicitly
        # block such a source and explain which current registry field needs
        # an approved formula/confirmed mechanism before issuing a passport.
        formula_book=load_book(path)
        try:
            columns=model.columns(registry)
            original=formula_book[registry.sheet]
            for key in CONTROL:
                cell=original.cell(project.first_row,columns[MAP[key]])
                if cell.data_type!='f':
                    raise Stop('В первом паспорте колонка '+MAP[key]+
                               ' содержит старую фиксированную дату, а не '
                               'формулу. Нужен подтверждённый источник даты '
                               'для НОВОЙ отгрузки; старую дату не переносим.')
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
    return dates|{'registry_date_origin':origin}
