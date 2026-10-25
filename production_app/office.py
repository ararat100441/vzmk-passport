"""Native Office contracts. No TEST climate or Kazakhstan bindings imported."""
from datetime import date,datetime
from decimal import Decimal
from pathlib import Path
import re,shutil
from passport_app.excel import excel_session,open_excel,load_book
from passport_app.vik import word_session,open_word
from passport_app.appendix import generate_appendix,inspect_template
from .storage import Stop,sha
from .model import filename,columns
from .word_controls import fill_word_controls
from .facts import validate_layer_ranges,validate_dry_averages,validate_layer_details
from .real_forms import values_for_documents,validate_uzk_details,UZK_HEADERS,uzk_details_to_table
from .automatic_acts import ordered_vik_names,joint_text

ROLES={'appendix':'.xlsx','vik':'.docx','uzk':'.docx','quality':'.docx'}
TOKEN=re.compile(r'\{\{[A-Z_0-9]+\}\}')

def fields(registry,row):
    cols=columns(registry);w=load_book(registry.path,values=True)
    try:values={k:w[registry.sheet].cell(row,c).value for k,c in cols.items()}
    finally:w.close()
    d={k:(v.strftime('%d.%m.%Y') if isinstance(v,(date,datetime)) else str(v or '')) for k,v in values.items()}
    if any(not d[k] for k in cols):raise Stop('Не заполнены реквизиты/не пересчитаны формулы: '+str(row))
    return d

def create_forms(folder):
    """Explicit creation in a NEW directory; requires separate human approval."""
    folder=Path(folder);folder.mkdir(exist_ok=False)
    with word_session() as word:
        for role,title,body in (
            ('vik','Акты визуального и измерительного контроля','ВИК: {{VIK_DATE}}\nСоединения: {{CONNECTIONS}}\n{{VIK_TABLE}}\nПервичные протоколы: {{EVIDENCE}}'),
            ('uzk','Заключение ультразвукового контроля','УЗК: {{UZK_DATE}}\n{{UZK_TABLE}}\nПротоколы: {{EVIDENCE}}'),
            ('quality','Документ о качестве','Масса: {{MASS}} кг\nСталь: {{STEEL}}\nКМ: {{KM}}\nКМД: {{KMD_COMPANY}}\nПриложение №1; акты ВИК, УЗК и АКЗ по комплекту.\n{{EVIDENCE}}')):
            doc=word.Documents.Add()
            try:
                doc.Content.Text=title+' № {{NUMBER}} от {{END_DATE}}\rДоговор: {{CONTRACT}}\rЗаказчик: {{CUSTOMER}}\rОбъект: {{OBJECT}}\rШифр: {{CIPHER}}\r'+body.replace('\n','\r')+'\rОтветственное лицо: {{RESPONSIBLE}}\rПодпись: __________________\r'
                doc.Content.Font.Name='Times New Roman';doc.Content.Font.Size=11
                doc.Paragraphs(1).Range.Font.Bold=True
                doc.SaveAs2(str(folder/(role+'.docx')),FileFormat=16)
            finally:doc.Close(SaveChanges=False)
    with excel_session() as app:
        for role in ('akz0','akz1','akz2','akz3','akz_appendix0','akz_appendix1','akz_appendix2','akz_appendix3'):
            b=app.Workbooks.Add(-4167)
            try:
                s=b.Worksheets(1);s.Name='АКЗ'
                if role.startswith('akz_appendix'):
                    s.Range('A1:G1').Merge();s.Range('A1').Value2='Приложение к АКЗ № {{NUMBER}} от {{AKZ_DATE}}'
                    s.Range('A2:G2').Merge();s.Range('A2').Value2='{{OBJECT}}'
                    s.Range('A3:G3').Value2=(('Марка','Наименование','Количество','Масса 1 шт','Масса всего','Площадь 1 шт','Площадь всего'),)
                    s.Range('A4').Value2='{{AKZ_MARKS}}';s.Range('A5').Value2='Итого'
                    s.Range('A3:G5').Borders.LineStyle=1;s.Columns('A:B').ColumnWidth=21;s.Columns('C:G').ColumnWidth=14
                else:
                    lines=['Акт АКЗ № {{NUMBER}} от {{AKZ_DATE}}','{{OBJECT}}','Договор: {{CONTRACT}}','Краска/материал: {{MATERIAL}}',
                        'Воздух, °C: {{AIR}}','Поверхность, °C: {{SURFACE}}','Влажность, %: {{HUMIDITY}}','Точка росы, °C: {{DEW}}',
                        'Толщина, мкм: {{THICKNESS}}','Послойные замеры/основание: {{BASIS}}','Протоколы: {{EVIDENCE}}','Ответственный: {{RESPONSIBLE}}','Подпись: __________________']
                    for n,line in enumerate(lines,1):s.Cells(n,1).Value2=line
                    s.Columns('A').ColumnWidth=95
                s.UsedRange.WrapText=True;s.UsedRange.Rows.AutoFit()
                s.PageSetup.Orientation=2 if role.startswith('akz_appendix') else 1
                s.PageSetup.PrintTitleRows='$1:$3';s.PageSetup.Zoom=False;s.PageSetup.FitToPagesWide=1;s.PageSetup.FitToPagesTall=False
                b.SaveAs(str(folder/(role+'.xlsx')),FileFormat=51)
            finally:b.Close(SaveChanges=False)
    return str(folder)

def replace_word(doc,values):
    # Find replacement uses literal Range.Text, preserving adjacent fields/layout.
    for key,value in values.items():
        if '{{' in str(value):raise Stop('Недопустимый placeholder в исходных данных: '+key)
        while True:
            r=doc.Content.Duplicate
            if not r.Find.Execute(FindText='{{'+key+'}}',MatchCase=True,MatchWildcards=False,Wrap=0):break
            r.Text=str(value)

def word_table(doc,token,headers,rows):
    """Insert an independent Word table, never merge with requisites table.

    Word may join two adjacent tables when Tables.Add replaces the only
    paragraph between them. Preserve a paragraph BEFORE the insertion point.
    The marker must occupy its own main-story paragraph, outside all tables.
    """
    if not headers or not isinstance(rows, (list,tuple)):
        raise Stop('Некорректные заголовки/строки Word.')
    if str(doc.Content.Text).count(token) != 1:
        raise Stop('Маркер таблицы Word должен встречаться один раз: '+token)

    def find_anchor():
        anchor=doc.Content.Duplicate
        if not anchor.Find.Execute(
            FindText=token,MatchCase=True,MatchWildcards=False,Wrap=0
        ):
            raise Stop('Шаблон Word не содержит '+token)
        if anchor.Tables.Count:
            raise Stop('Маркер таблицы Word находится внутри другой таблицы.')
        if anchor.Paragraphs.Count != 1:
            raise Stop('Маркер Word должен занимать отдельный абзац.')
        paragraph=anchor.Paragraphs(1).Range
        if paragraph.Text.strip(' \t\r\n\x07') != token:
            raise Stop('Маркер Word должен быть единственным текстом абзаца.')
        return anchor,paragraph

    before_count=doc.Tables.Count
    anchor,paragraph=find_anchor()
    # Word intentionally treats consecutive tables as a single table. Insert
    # a durable separator paragraph before the anchor paragraph. Re-find the
    # marker after insertion because Word mutates character positions.
    paragraph.InsertParagraphBefore()
    anchor,_=find_anchor()
    anchor.Text=''
    table=doc.Tables.Add(anchor,len(rows)+1,len(headers))
    if doc.Tables.Count != before_count+1:
        raise Stop('Word присоединил новую таблицу к существующей; выпуск остановлен.')
    try:
        if table.Rows(1).Cells.Count != len(headers):
            raise Stop('Word изменил структуру первой строки таблицы: '
                       +str(table.Rows(1).Cells.Count)+' вместо '
                       +str(len(headers)))
        table.Borders.Enable=True
        # Fit the 8-column UZK table to the landscape document width.
        # Word constants: wdAutoFitWindow=2.
        table.AutoFitBehavior(2)
        for col,value in enumerate(headers,1):
            table.Cell(1,col).Range.Text=str(value)
        table.Rows(1).HeadingFormat=True
        table.Rows(1).Range.Font.Bold=True
        for n,row in enumerate(rows,2):
            if len(row) != len(headers):
                raise Stop('Строка Word №'+str(n-1)+' имеет неверное число колонок.')
            for col,value in enumerate(row,1):
                table.Cell(n,col).Range.Text=str(value)
        if (doc.Tables.Count != before_count+1 or
                table.Rows(1).Cells.Count != len(headers)):
            raise Stop('Структура таблицы Word изменилась после заполнения.')
    except Stop:
        raise
    except Exception as exc:
        raise Stop('Ошибка Word при формировании отдельной таблицы '
                   +token+': '+str(exc)) from exc
    return table



# Compact two-tier headers in the historical VZMK appendix are vertically
# merged (A3:A4, etc.). The original 15 pt rows were too short to guarantee
# legible repeats when many data lines printed across three PDF pages.
# Change heights ONLY in temporary/newly issued copies; never in sources.
_MIN_HEADER_ROW_HEIGHT_PT = 28.0
_MIN_DATA_ROW_HEIGHT_PT = 18.0
_MAX_DATA_ROW_HEIGHT_PT = 220.0


def fit_appendix_body(sheet,first_data_row,last_data_row):
    """Fit *actual text* after expanding data, not fixed-size diagnostic rows.

    All body rows must be standalone (not merged) so native AutoFit can
    increase heights for long marks/names. Reject a row that would require
    more than a reasonable page height instead of silently cropping it.
    Original column widths, fonts, number formats, formulas are preserved.
    """
    if (type(first_data_row) is not int or type(last_data_row) is not int
            or first_data_row < 4 or last_data_row < first_data_row
            or last_data_row - first_data_row > 10000):
        raise Stop('Недопустимое число строк печати приложения.')
    body=sheet.Range('A'+str(first_data_row)+':B'+str(last_data_row))
    # Excel merged cells prevent reliable automatic row-height calculation.
    if body.MergeCells:
        raise Stop('В строках марок приложения имеются объединённые ячейки; '
                   'безопасная высота для печати не подтверждена.')
    body.WrapText=True
    body.ShrinkToFit=False
    body.VerticalAlignment=-4108  # xlCenter
    sheet.Rows(str(first_data_row)+':'+str(last_data_row)).AutoFit()
    for row in range(first_data_row,last_data_row+1):
        native=sheet.Rows(row)
        if native.Hidden:
            raise Stop('Строка марки скрыта от печати: '+str(row))
        height=native.RowHeight
        if not isinstance(height,(int,float)) or height <= 0:
            raise Stop('Excel не определил высоту строки марки '+str(row))
        if height > _MAX_DATA_ROW_HEIGHT_PT:
            raise Stop('Слишком длинное наименование для печати в строке '
                       +str(row)+'; требуется проверка макета.')
        if height < _MIN_DATA_ROW_HEIGHT_PT:
            native.RowHeight=_MIN_DATA_ROW_HEIGHT_PT
    return last_data_row-first_data_row+1


def verify_appendix_print_layout(sheet,expected=None):
    """Read-only native Excel check for repeated compact headings."""
    actual=str(sheet.PageSetup.PrintTitleRows or '').replace('$','')
    # Excel may prepend the sheet name to the A1-style range.
    actual=actual.rsplit('!',1)[-1].strip().upper()
    if actual not in ('3:3','3:4'):
        raise Stop('Excel повторяет длинные строки названия/объекта или не '
                   'сохранил компактную шапку таблицы: '+actual)
    if expected is not None and actual!=expected:
        raise Stop('Excel не сохранил ожидаемую шапку '+expected+
                   '; фактически '+actual)
    if not str(sheet.PageSetup.PrintArea or '').strip():
        raise Stop('Отсутствует область печати приложения.')
    # Persisted PrintTitleRows alone is not sufficient: the PR13 PDF had
    # cropped headings on page three despite correct repeat-row metadata.
    first,last=map(int,actual.split(':'))
    for row in range(first,last+1):
        height=sheet.Rows(row).RowHeight
        if not isinstance(height,(int,float)) or height < _MIN_HEADER_ROW_HEIGHT_PT:
            raise Stop('Повторяемый заголовок слишком низкий после Excel Save/Open: '
                       +str(row)+' / '+str(height))
    return actual


def set_appendix_print_titles(sheet,first_data_row):
    """Repeat ONLY compact table labels, not long title/object on later pages.

    First page prints title+object from PrintArea (starting A1); only two
    preceding column-header rows are repeated on subsequent pages. Both
    historical Appendix #1 and converted АКЗ appendix start marks at row 5.
    """
    if type(first_data_row) is not int or first_data_row < 4:
        raise Stop('Не найдены двухстрочные заголовки приложения.')
    # Generic one-line header begins at row 3; historical two-tier header
    # occupies rows 3-4. The long OBJECT row 2 is never repeated.
    first_header=max(3,first_data_row-2)
    last_header=first_data_row-1
    if (not str(sheet.Cells(first_header,1).Value2 or '').strip()
            or not str(sheet.Cells(first_data_row,1).Value2 or '').strip()):
        raise Stop('Нет заголовка или строки марок для печати.')
    # Merged column header cells need real physical height on *every*
    # printed page, especially pages 3+ with repeated titles.
    heading=sheet.Range('A'+str(first_header)+':G'+str(last_header))
    heading.WrapText=True
    heading.VerticalAlignment=-4108  # xlCenter
    for row in range(first_header,last_header+1):
        line=sheet.Rows(row)
        if not isinstance(line.RowHeight,(int,float)):
            raise Stop('Excel не определил высоту заголовка '+str(row))
        if line.RowHeight < _MIN_HEADER_ROW_HEIGHT_PT:
            line.RowHeight=_MIN_HEADER_ROW_HEIGHT_PT
    sheet.PageSetup.PrintTitleRows='$'+str(first_header)+':$'+str(last_header)
    sheet.PageSetup.FitToPagesWide=1
    sheet.PageSetup.FitToPagesTall=False
    sheet.PageSetup.Zoom=False
    verify_appendix_print_layout(sheet,str(first_header)+':'+str(last_header))
    return first_header,last_header


# The approved VZMK historical 3-layer XLSX uses *built-in* Excel number
# format IDs 14 (dates at H4, E22/H22/J22) and 20 (times at E23/H23/J23).
# Applying text-based NumberFormat or NumberFormatLocal from Python destroyed
# those native styles on the user's ru-RU installation: 19.00.2026 / literal
# dd.mm.yyyy. Keep the native Excel styles exactly as provided in the form.
# We write only Value2, preserving the numeric serial used by workbook formulas.
# A template without correctly preformatted date/time cells must fail closed.

def _native_time_matches(visible, expected):
    """Accept a valid native h:mm as well as hh:mm, never a literal mask."""
    try:
        shown = str(visible).strip().split(':')
        target = str(expected).strip().split(':')
        if len(shown) != 2 or len(target) != 2:
            return False
        if not all(part.isascii() and part.isdecimal() for part in shown+target):
            return False
        hh, mm = map(int, shown)
        th, tm = map(int, target)
        return 0 <= hh < 24 and 0 <= mm < 60 and (hh,mm) == (th,tm)
    except (TypeError, ValueError):
        return False


def _set_ru_excel_date_time(cell, serial, rendered, *, is_time, key):
    """Require *existing* Excel date/time formatting, never set mask strings.

    Document fields H4 / E22 / E23 have built-in format styles. Do not replace
    them with a custom mask. Reject general-format cells and incorrect display;
    user must provide an approved native preformatted template.
    """
    try:
        before_local = cell.NumberFormatLocal
        before_invariant = cell.NumberFormat
        cell.Value2 = serial
        stored = cell.Value2
        visible = str(cell.Text or '').strip()
        after_local = cell.NumberFormatLocal
        after_invariant = cell.NumberFormat
    except Exception as exc:
        raise Stop('Excel: не удалось прочитать исходный формат даты/времени '
                   + key + ': ' + str(exc)) from exc

    if (isinstance(stored, bool) or not isinstance(stored, (int, float))
            or abs(float(stored) - float(serial)) > 1e-9):
        raise Stop('Excel: неверное числовое значение даты/времени ' + key)

    if after_local != before_local or after_invariant != before_invariant:
        raise Stop('Excel: изменился исходный формат ячейки '+key)

    accepted = (_native_time_matches(visible, rendered)
                if is_time else visible == rendered)
    if not accepted:
        raise Stop('Excel: исходная ячейка шаблона АКЗ не имеет корректного '
                   'встроенного формата даты/времени '+key+
                   '; отображается '+visible+'; ожидалось '+rendered+
                   '. Используйте отдельную согласованную XLSX-форму с '
                   'исходными форматами Excel; маски программой не заменяются.')

def tokens_excel(sheet,values):
    # Preserve number types in Excel when a cell contains a single numeric
    # placeholder, especially climate and coating-range bounds used by formulas.
    numeric_keys={'AIR','SURFACE','HUMIDITY','DEW','THICKNESS'}
    numeric_keys.update('LAYER_'+str(n)+'_'+edge
                        for n in (1,2,3) for edge in ('MIN','MAX','DRY_AVG','TOTAL_MIN','TOTAL_MAX'))
    numeric_keys.update('AKZ_LAYER_'+str(n)+'_WET_AVG' for n in (1,2,3))
    date_keys={'AKZ_DATE'}|{'AKZ_LAYER_'+str(n)+'_APPLICATION_DATE' for n in (1,2,3)}
    time_keys={'AKZ_LAYER_'+str(n)+'_'+suffix for n in (1,2,3) for suffix in ('TIME_START','TIME_END')}
    used=sheet.UsedRange
    for row in range(used.Row,used.Row+used.Rows.Count):
        for col in range(used.Column,used.Column+used.Columns.Count):
            cell=sheet.Cells(row,col)
            if cell.HasFormula:continue
            v=cell.Value2
            if isinstance(v,str) and '{{' in v:
                key=v[2:-2] if v.startswith('{{') and v.endswith('}}') else ''
                if key in date_keys and key in values:
                    raw=str(values[key]).strip()
                    try:
                        actual=date.fromisoformat(raw) if '-' in raw else datetime.strptime(raw,'%d.%m.%Y').date()
                    except ValueError as exc:
                        raise Stop('Некорректная дата Excel АКЗ: '+key) from exc
                    _set_ru_excel_date_time(
                        cell, (actual-date(1899,12,30)).days,
                        actual.strftime('%d.%m.%Y'), is_time=False, key=key,
                    )
                    continue
                if key in time_keys and key in values:
                    raw=str(values[key]).strip()
                    import re as regex
                    if not regex.fullmatch(r'(?:[01]\d|2[0-3]):[0-5]\d',raw):
                        raise Stop('Некорректное время слоя АКЗ: '+key)
                    hh,mm=map(int,raw.split(':'))
                    _set_ru_excel_date_time(
                        cell, (hh*60+mm)/1440,
                        f'{hh:02d}:{mm:02d}', is_time=True, key=key,
                    )
                    continue
                if key in numeric_keys and key in values:
                    numeric_value=Decimal(str(values[key]).replace(',','.'))
                    if not numeric_value.is_finite():
                        raise Stop('Некорректное числовое поле Excel: '+key)
                    cell.Value2=float(numeric_value)
                    continue
                for key,value in values.items():v=v.replace('{{'+key+'}}',str(value))
                cell.Value2=v

def akz_values(climate,facts=None):
    """One daily climate set shared by all three coating layers.

    Range tokens can be placed on an approved 3-layer Excel form:
    {{LAYER_1_MIN}}, {{LAYER_1_MAX}}, {{LAYER_1_RANGE}}, etc.
    A range is a specification, never a replacement for actual wet/dry
    measurements and their primary protocols.
    """
    values={str(key).upper():value for key,value in climate.items()
            if key not in ('variant','layers','dry_averages','layer_details')}
    if climate.get('layers'):
        checked=validate_dry_averages(climate['layers'],climate.get('dry_averages'))
        for entry in validate_layer_ranges(climate['layers']):
            prefix='LAYER_'+str(entry['layer'])+'_'
            values[prefix+'MIN']=entry['min']
            values[prefix+'MAX']=entry['max']
            values[prefix+'RANGE']=entry['min']+'–'+entry['max']
        for entry in checked:
            prefix='LAYER_'+str(entry['layer'])+'_'
            values[prefix+'DRY_AVG']=entry['average']
            values[prefix+'TOTAL_MIN']=entry['required_total_min']
            values[prefix+'TOTAL_MAX']=entry['required_total_max']
            values[prefix+'TOTAL_RANGE']=entry['required_total_min']+'–'+entry['required_total_max']
    if climate.get('layer_details'):
        if not isinstance(facts,dict):raise Stop('Нет подтверждённых фактов по слоям АКЗ.')
        for row in validate_layer_details(climate['layer_details'],facts):
            prefix='AKZ_LAYER_'+str(row['layer'])+'_'
            for key in ('material','paint_batch','hardener_batch','color','application_date','time_start','time_end'):
                values[prefix+key.upper()]=row[key]
            values[prefix+'WET_AVG']=row['wet_average']
    return values


def akz_rows(sheet,rows):
    marker=sheet.UsedRange.Find(What='{{AKZ_MARKS}}',LookAt=1)
    if marker is None:raise Stop('Шаблон приложения АКЗ не содержит {{AKZ_MARKS}}.')
    first=marker.Row;total=first+len(rows)
    if marker.Column!=1 or sheet.Cells(first+1,1).Value2!='Итого':raise Stop('Неверная структура шаблона АКЗ.')
    if len(rows)>1:
        sheet.Rows(f'{first+1}:{total-1}').Insert()
        sheet.Range(f'A{first}:G{total-1}').FillDown()
    for n,row in enumerate(rows,first):
        sheet.Range(f'A{n}:G{n}').Value2=(tuple(float(v) if isinstance(v,Decimal) else v for v in row),)
    sheet.Cells(total,1).Value2='Итого'
    for col in ('C','E','G'):sheet.Range(col+str(total)).Formula=f'=SUM({col}{first}:{col}{total-1})'
    for col in ('E','G'):sheet.Range(f'{col}{first}:{col}{total}').NumberFormatLocal='0'+str(sheet.Application.DecimalSeparator)+'00'
    sheet.Range(f'A{first}:G{total}').Borders.LineStyle=1
    sheet.Range(f'A{total}:G{total}').Font.Bold=True
    fit_appendix_body(sheet,first,total-1)
    sheet.PageSetup.PrintArea=f'$A$1:$G${total}'
    set_appendix_print_titles(sheet,first)

def render(profile,project,registry,row,vom,facts,folder,training=False):
    """Approved tokens prevent stale facts from another issued passport."""
    values=fields(registry,row);values.update(RESPONSIBLE=facts['responsible'],
        EVIDENCE='; '.join(Path(e['path']).name+' SHA256 '+e['sha256'] for e in facts['evidence']))
    for k in ('start_date','vik_date','uzk_date','akz_date'):
        actual=date.fromisoformat(facts[k]).strftime('%d.%m.%Y')
        if values[k.upper()]!=actual:raise Stop('Фактическая дата '+k+' ('+actual+') не совпадает с формулой/данными реестра ('+values[k.upper()]+'). Ничего не подставлено автоматически.')
    values['MASS']=str(vom.mass);values['CONNECTIONS']=vom.connections
    values['NUMBER']=facts['number']
    values['VIK_MARKS']='; '.join(ordered_vik_names(vom))
    values['UZK_DESIGN']='; '.join(
        row['mark']+': '+joint_text(row)
        for row in facts.get('uzk_design',[])
    )
    values.update(values_for_documents(facts,{c['variant'] for c in facts['climates']}))
    prefix='УЧЕБНЫЙ КОМПЛЕКТ — НЕ ДЛЯ ВЫПУСКА' if training else 'ЧЕРНОВИК — НЕ ДЛЯ ВЫПУСКА'
    by_mark={str(m['mark']):m for m in facts['marks']}
    totals={}
    for r in vom.rows:totals[str(r[0])]=totals.get(str(r[0]),0)+r[2]
    outputs=[]
    with excel_session() as app:
        for role,ref in profile['templates'].items():
            if role in ('vik','uzk','quality'):continue
            if role.startswith('akz'):
                variant=int(role[-1]);climate=next((c for c in facts['climates'] if c['variant']==variant),None)
                if climate is None:continue
            target=folder/(filename(facts['number'])+' - '+role+'.xlsx');shutil.copyfile(ref['path'],target)
            if role=='appendix':
                # The first-passport template is valid BEFORE number/date change.
                # A second inspect_template AFTER generation would reject the
                # newly numbered passport as no longer the first template.
                layout,_,_=inspect_template(target,project)
                generate_appendix(app,target,project,vom,{'number':facts['number'],'date':date.fromisoformat(facts['date']),'object':values['OBJECT']})
                # Set compact repeating table headers on NEW production copy only.
                b=open_excel(app,target)
                try:
                    sheet=b.Worksheets(layout.sheet)
                    fit_appendix_body(sheet,layout.start,layout.start+len(vom.rows)-1)
                    set_appendix_print_titles(sheet,layout.start)
                    b.Save()
                finally:
                    b.Close(SaveChanges=False)
            else:
                b=open_excel(app,target)
                try:
                    s=b.Worksheets(1)
                    if role.startswith('akz_appendix'):
                        selected=[r for r in vom.rows if by_mark[str(r[0])]['akz']==variant]
                        akz_rows(s,selected)
                    native_values=values|akz_values(climate,facts)
                    tokens_excel(s,native_values)
                    app.CalculateFull();b.Save()
                finally:b.Close(SaveChanges=False)
            b=open_excel(app,target)
            try:
                for s in b.Worksheets:
                    s.PageSetup.CenterHeader=prefix
                    for r in s.UsedRange.Value2:
                        if any(isinstance(v,str) and TOKEN.search(v) for v in r):raise Stop('Незаполненное поле шаблона '+role)
                b.Save()
            finally:b.Close(SaveChanges=False)
            outputs.append({'role':role,'path':str(target)})
    with word_session() as app:
        first_sources=(profile.get('first_passport_sources',{})
                       .get('sources',{}))
        for role in ('vik','uzk','quality'):
            ref=profile['templates'][role];target=folder/(filename(facts['number'])+' - '+role+'.docx')
            original=first_sources.get(role)
            if original and role in ('vik','uzk'):
                from .highlighted_reference import (
                    create_vik_copy,create_uzk_copy,
                    supports_vik_first,supports_uzk_first,
                )
                source=original['path']
                if sha(source)!=original['sha256']:
                    raise Stop('Исторический исходник '+role+
                               ' изменился после проверки.')
                if role=='vik' and supports_vik_first(source):
                    # Current item NAMES, never a semicolon list of mark IDs.
                    create_vik_copy(
                        source,target,number=facts['number'],
                        vik_date=facts['vik_date'],
                        item_names=[row[1] for row in vom.rows],
                    )
                    doc=open_word(app,target,readonly=False)
                    try:
                        doc.Range(0,0).InsertBefore(prefix+'\r')
                        doc.Repaginate();doc.Save()
                    finally:doc.Close(SaveChanges=False)
                    outputs.append({'role':role,'path':str(target)})
                    continue
                if role=='uzk' and supports_uzk_first(source):
                    if not facts.get('uzk_design'):
                        raise Stop('Не вычислены расчётные характеристики '
                                   'текущих марок для исходной формы УЗК.')
                    create_uzk_copy(
                        source,target,number=facts['number'],
                        uzk_date=facts['uzk_date'],
                        design_rows=facts['uzk_design'],
                    )
                    doc=open_word(app,target,readonly=False)
                    try:
                        doc.Range(0,0).InsertBefore(prefix+'\r')
                        doc.Repaginate();doc.Save()
                    finally:doc.Close(SaveChanges=False)
                    outputs.append({'role':role,'path':str(target)})
                    continue
            shutil.copyfile(ref['path'],target);doc=open_word(app,target,readonly=False)
            try:
                # Native tagged Word fields are allowed only when still blank.
                has_number_token = '{{NUMBER}}' in doc.Content.Text
                bound_controls = fill_word_controls(doc, values)
                if not has_number_token and 'NUMBER' not in bound_controls:
                    raise Stop('Нужна утверждённая форма с {{NUMBER}} или пустым полем PASSPORT_NO/NUMBER: '+ref['path'])
                if role=='uzk':
                    if '{{UZK_DETAILS_TABLE}}' in doc.Content.Text:
                        details=validate_uzk_details(facts,vom,required=True)
                        headers=UZK_HEADERS[:8]
                        rows=[entry[:8] for entry in uzk_details_to_table(details)]
                        word_table(doc,'{{UZK_DETAILS_TABLE}}',headers,rows)
                    else:
                        rows=[(m,totals[m],by_mark[m]['uzk_quantity'],by_mark[m]['uzk_result'],by_mark[m].get('basis','')) for m in totals]
                        word_table(doc,'{{UZK_TABLE}}',('Марка','ВОМ, шт','УЗК, шт','Результат','Основание'),rows)
                if role=='vik' and '{{VIK_TABLE}}' in doc.Content.Text:
                    rows=[(m,totals[m],by_mark[m].get('vik_result','УЧЕБНЫЕ ДАННЫЕ'),by_mark[m].get('vik_basis','УЧЕБНЫЕ ДАННЫЕ')) for m in totals]
                    word_table(doc,'{{VIK_TABLE}}',('Марка','ВОМ, шт','Результат ВИК','Первичный протокол'),rows)
                replace_word(doc,values)
                if TOKEN.search(doc.Content.Text):raise Stop('Незаполненные поля Word: '+role)
                doc.Range(0,0).InsertBefore(prefix+'\r')
                doc.Repaginate();doc.Save()
            finally:doc.Close(SaveChanges=False)
            outputs.append({'role':role,'path':str(target)})
    return outputs

def verify_documents(outputs,number,training=False,draft=True):
    report=[]
    forbidden=('SAMPLE_ONLY','RANDBETWEEN','НЕ ПОДТВЕРЖДЕНО')
    with excel_session() as app:
        for out in outputs:
            path=Path(out['path'])
            if path.suffix!='.xlsx':continue
            b=open_excel(app,path)
            try:
                app.CalculateFullRebuild()
                for s in b.Worksheets:
                    text=' '.join(str(v) for r in s.UsedRange.Value2 for v in r if v is not None)
                    formulas=' '.join(str(v) for r in s.UsedRange.Formula for v in r if v is not None)
                    if TOKEN.search(text) or any(t in (text+' '+formulas).upper() for t in forbidden):raise Stop('Неподтверждённые/случайные данные Excel: '+str(path))
                    if not training and ('ТЕСТ' in text.upper() or 'TEST' in text.upper()):raise Stop('TEST в рабочем документе: '+str(path))
                    if out['role']=='appendix' or out['role'].startswith('akz_appendix'):
                        verify_appendix_print_layout(s)
                    errors=None
                    try:errors=s.UsedRange.SpecialCells(-4123,16)
                    except Exception:pass
                    if errors is not None:raise Stop('Ошибка формулы Excel: '+str(path))
                report.append({'file':str(path),'engine':'Microsoft Excel','status':'PASS'})
            finally:b.Close(SaveChanges=False)
    with word_session() as app:
        for out in outputs:
            path=Path(out['path'])
            if path.suffix!='.docx':continue
            d=open_word(app,path,readonly=True)
            try:
                d.Repaginate();text=d.Content.Text
                if number not in text or TOKEN.search(text):raise Stop('Неверный номер/незаполненные поля Word: '+str(path))
                if any(t in text.upper() for t in forbidden) or not training and any(t in text.upper() for t in ('ТЕСТ','TEST')):raise Stop('TEST/неподтверждённые данные Word: '+str(path))
                report.append({'file':str(path),'engine':'Microsoft Word','pages':d.ComputeStatistics(2),'status':'PASS'})
            finally:d.Close(SaveChanges=False)
    return report

def release_copies(outputs,folder,training):
    result=[]
    for out in outputs:
        target=folder/Path(out['path']).name;shutil.copyfile(out['path'],target);result.append(out|{'path':str(target)})
    if training:return result
    with excel_session() as app:
        for out in result:
            if Path(out['path']).suffix!='.xlsx':continue
            b=open_excel(app,out['path'])
            try:
                for s in b.Worksheets:
                    if s.PageSetup.CenterHeader!='ЧЕРНОВИК — НЕ ДЛЯ ВЫПУСКА':raise Stop('Изменена маркировка черновика.')
                    s.PageSetup.CenterHeader=''
                b.Save()
            finally:b.Close(SaveChanges=False)
    with word_session() as app:
        for out in result:
            if Path(out['path']).suffix!='.docx':continue
            d=open_word(app,out['path'],readonly=False)
            try:
                p=d.Paragraphs(1).Range
                if p.Text.rstrip('\r')!='ЧЕРНОВИК — НЕ ДЛЯ ВЫПУСКА':raise Stop('Изменена маркировка черновика Word.')
                p.Delete();d.Repaginate();d.Save()
            finally:d.Close(SaveChanges=False)
    return result

def inspect_existing(registry_path,project,number,folder,vom=None):
    from passport_app.registry import read_registry
    from .storage import pins,verify_pins
    from .model import parse_number,visible_vom
    parse_number(number);registry=read_registry(registry_path)
    matches=[r for r in registry.projects.get(project,[]) if r['number']==number]
    if len(matches)!=1:raise Stop('Существующий паспорт не найден однозначно.')
    root=Path(folder);files=[p for p in root.iterdir() if p.suffix.lower() in ('.docx','.xlsx') and not p.name.startswith('~$')]
    if not files:raise Stop('Нет исходных DOCX/XLSX.')
    before=pins([registry_path,*files,*([vom] if vom else [])]);report=[]
    prefix=filename(number)
    with word_session() as app:
        for path in files:
            if path.suffix!='.docx' or not path.name.startswith(prefix):continue
            d=open_word(app,path,readonly=True)
            try:
                d.Repaginate();text=d.Content.Text
                if number not in text:raise Stop('В теле Word отсутствует исходный номер: '+str(path))
                report.append({'file':str(path),'pages':d.ComputeStatistics(2),'number_preserved':True})
            finally:d.Close(SaveChanges=False)
    with excel_session() as app:
        for path in files:
            if path.suffix!='.xlsx':continue
            b=app.Workbooks.Open(str(path.resolve()),UpdateLinks=0,ReadOnly=True,AddToMru=False)
            try:report.append({'file':str(path),'sheets':b.Worksheets.Count,'read_only':bool(b.ReadOnly)})
            finally:b.Close(SaveChanges=False)
    comparison='NOT_VERIFIED_REAL_VOM_MISSING'
    if vom:
        v=visible_vom(vom)
        if abs(v.mass-Decimal(str(matches[0]['mass'])))>Decimal('.005'):raise Stop('Реальная ВОМ не совпала по массе с существующим паспортом.')
        comparison={'marks':len(v.rows),'quantity':str(v.quantity),'mass_kg':str(v.mass),'status':'PASS'}
    verify_pins(before)
    return {'status':'PASS','scope':'READ_ONLY_EXISTING_OFFICE_DOCUMENTS','project':project,'number':number,'documents':report,'vom_comparison':comparison,'source_hashes_unchanged':True}
