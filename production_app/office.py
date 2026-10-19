"""Native Office contracts. No TEST climate or Kazakhstan bindings imported."""
from datetime import date,datetime
from decimal import Decimal
from pathlib import Path
import re,shutil
from passport_app.excel import excel_session,open_excel,load_book
from passport_app.vik import word_session,open_word
from passport_app.appendix import generate_appendix
from .storage import Stop,sha
from .model import filename,columns
from .word_controls import fill_word_controls
from .facts import validate_layer_ranges,validate_dry_averages

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
    r=doc.Content.Duplicate
    if not r.Find.Execute(FindText=token,MatchCase=True,Wrap=0):raise Stop('Шаблон УЗК не содержит '+token)
    r.Text='';table=doc.Tables.Add(r,len(rows)+1,len(headers))
    table.Borders.Enable=True
    for c,value in enumerate(headers,1):table.Cell(1,c).Range.Text=value
    table.Rows(1).HeadingFormat=True;table.Rows(1).Range.Font.Bold=True
    for n,row in enumerate(rows,2):
        for c,value in enumerate(row,1):table.Cell(n,c).Range.Text=str(value)

def tokens_excel(sheet,values):
    # Preserve number types in Excel when a cell contains a single numeric
    # placeholder, especially climate and coating-range bounds used by formulas.
    numeric_keys={'AIR','SURFACE','HUMIDITY','DEW','THICKNESS'}
    numeric_keys.update('LAYER_'+str(n)+'_'+edge
                        for n in (1,2,3) for edge in ('MIN','MAX','DRY_AVG','TOTAL_MIN','TOTAL_MAX'))
    used=sheet.UsedRange
    for row in range(used.Row,used.Row+used.Rows.Count):
        for col in range(used.Column,used.Column+used.Columns.Count):
            cell=sheet.Cells(row,col)
            if cell.HasFormula:continue
            v=cell.Value2
            if isinstance(v,str) and '{{' in v:
                key=v[2:-2] if v.startswith('{{') and v.endswith('}}') else ''
                if key in numeric_keys and key in values:
                    numeric_value=Decimal(str(values[key]).replace(',','.'))
                    if not numeric_value.is_finite():
                        raise Stop('Некорректное числовое поле Excel: '+key)
                    cell.Value2=float(numeric_value)
                    continue
                for key,value in values.items():v=v.replace('{{'+key+'}}',str(value))
                cell.Value2=v

def akz_values(climate):
    """One daily climate set shared by all three coating layers.

    Range tokens can be placed on an approved 3-layer Excel form:
    {{LAYER_1_MIN}}, {{LAYER_1_MAX}}, {{LAYER_1_RANGE}}, etc.
    A range is a specification, never a replacement for actual wet/dry
    measurements and their primary protocols.
    """
    values={str(key).upper():value for key,value in climate.items()
            if key not in ('variant','layers','dry_averages')}
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
    sheet.PageSetup.PrintArea=f'$A$1:$G${total}';sheet.PageSetup.PrintTitleRows=f'$1:${first-1}'
    sheet.PageSetup.FitToPagesWide=1;sheet.PageSetup.FitToPagesTall=False;sheet.PageSetup.Zoom=False

def render(profile,project,registry,row,vom,facts,folder,training=False):
    """Approved tokens prevent stale facts from another issued passport."""
    values=fields(registry,row);values.update(RESPONSIBLE=facts['responsible'],
        EVIDENCE='; '.join(Path(e['path']).name+' SHA256 '+e['sha256'] for e in facts['evidence']))
    for k in ('start_date','vik_date','uzk_date','akz_date'):
        actual=date.fromisoformat(facts[k]).strftime('%d.%m.%Y')
        if values[k.upper()]!=actual:raise Stop('Фактическая дата '+k+' ('+actual+') не совпадает с формулой/данными реестра ('+values[k.upper()]+'). Ничего не подставлено автоматически.')
    values['MASS']=str(vom.mass);values['CONNECTIONS']=vom.connections
    values['NUMBER']=facts['number']
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
                generate_appendix(app,target,project,vom,{'number':facts['number'],'date':date.fromisoformat(facts['date']),'object':values['OBJECT']})
            else:
                b=open_excel(app,target)
                try:
                    s=b.Worksheets(1)
                    if role.startswith('akz_appendix'):
                        selected=[r for r in vom.rows if by_mark[str(r[0])]['akz']==variant]
                        akz_rows(s,selected)
                    native_values=values|akz_values(climate)
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
        for role in ('vik','uzk','quality'):
            ref=profile['templates'][role];target=folder/(filename(facts['number'])+' - '+role+'.docx')
            shutil.copyfile(ref['path'],target);doc=open_word(app,target,readonly=False)
            try:
                # Native tagged Word fields are allowed only when still blank.
                has_number_token = '{{NUMBER}}' in doc.Content.Text
                bound_controls = fill_word_controls(doc, values)
                if not has_number_token and 'NUMBER' not in bound_controls:
                    raise Stop('Нужна утверждённая форма с {{NUMBER}} или пустым полем PASSPORT_NO/NUMBER: '+ref['path'])
                if role=='uzk':
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
                    if not s.PageSetup.PrintArea and out['role'].startswith('akz_appendix'):raise Stop('Нет области печати АКЗ.')
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
