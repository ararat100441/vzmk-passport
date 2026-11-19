"""Existing one/two/three-coat AKZ XLSX and appendix (read only until copied).

In new copies row 27 contains independently measured wet coat means;
row 28 contains cumulative dry means. Legacy one-coat calls without range
input retain the reference wet formula (never certify it).
Registry supplies number/date; current ranges supply thickness and climate.
Current visible VOM supplies all seven columns of every appendix body row.
Formatting, nonmeasurement formulas and other reference cells stay intact.
"""
from datetime import date,datetime
from decimal import Decimal
from zipfile import ZipFile
import re
from xml.etree import ElementTree as ET
from .storage import Stop

X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NUMBER=re.compile(r'^[1-9][0-9]*(?:/[1-9][0-9]*)?/20[0-9]{2}(?:-[1-3])?$')
TITLE=re.compile(r'^Приложение № (\d+/20\d\d(?:-[1-3])?) к акту АКЗ № (\d+/20\d\d(?:-[1-3])?)\s+от (\d\d\.\d\d\.20\d\d) г\.$')


def act_number(number,variant):
    if type(variant) is not int or variant not in range(4):raise Stop('Неверный вид АКЗ.')
    if '-' in str(number):raise Stop('Для нумерации АКЗ нужен базовый номер паспорта, без суффикса акта.')
    return _number(number)+('-'+str(variant) if variant else '')


def _cells(path):
    with ZipFile(path) as z:
        root=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            strings=ET.fromstring(z.read('xl/sharedStrings.xml'))
            shared=[''.join(t.text or '' for t in si.iter(X+'t'))
                    for si in strings.findall(X+'si')]
        result={}
        for cell in root.iter(X+'c'):
            v=cell.find(X+'v');f=cell.find(X+'f')
            text=v.text if v is not None else None
            if text is not None and cell.get('t')=='s':
                text=shared[int(text)]
            inline=cell.find(X+'is')
            if inline is not None and cell.get('t')=='inlineStr':
                text=''.join(t.text or '' for t in inline.iter(X+'t'))
            if text is not None or f is not None:
                result[cell.get('r')]={
                    'value':text,'formula':f.text if f is not None else None}
        return result


def is_first_akz(path):
    try:
        c=_cells(path)
        if _two_coat_reference(c) or _three_coat_reference(c):return True
        return (bool(re.fullmatch(r'\d+(?:/\d+)?/20\d\d(?:-[1-3])?\s+от\s*',
                      str(c.get('E4',{}).get('value',''))))
                and all(c.get(p,{}).get('formula')==f for p,f in (
                    ('G8','E4'),('I8','H4'),('E22','H4-1'),
                    ('E30','E22'),('D49','H4')))
                and (c.get('E27',{}).get('formula')=='E28/0.6' or
                     _numeric(c.get('E27',{}).get('value'),'thickness')>0)
                and all(c.get(p,{}).get('value') is not None for p in
                    ('H4','E28','E31','E32','E33','E34')))
    except Exception:return False


def _three_coat_reference(c):
    """Recognize the existing E/H/J act, never infer coats from an act suffix."""
    return (bool(re.fullmatch(r'\d+(?:/\d+)?/20\d\d(?:-[1-3])?\s+от\s*',
                             str(c.get('E4',{}).get('value',''))))
            and c.get('H4',{}).get('value') is not None
            and all(str(c.get(col+'17',{}).get('value'))==str(i)
                    for i,col in enumerate(('E','H','J'),1))
            and _numeric(c.get('E27',{}).get('value'),'thickness')>0
            and all(c.get(p,{}).get('formula')==f for p,f in (
                ('G8','E4'),('I8','H4'),('D49','H4'),
                ('E22','H22-1'),('H22','J22-1'),('J22','H4-1'),
                ('E30','E22'),('H30','H22'),('J30','J22')))
            and all(c.get(col+str(row),{}).get('value') is not None
                    for col in ('E','H','J') for row in (28,31,32,33,34))
            and all(c.get(col+'27',{}).get('formula')==f'{col}28*100/L27'
                    or (c.get(col+'27',{}).get('formula') is None
                        and _numeric(c.get(col+'27',{}).get('value'),'thickness')>0)
                    for col in ('H','J')))


def _phosphorit_layout(value, formula):
    return (value('E17') in (1,'1') and value('H17') in (None,'')
            and value('I17') in (2,'2') and value('J17') in (None,'')
            and 'ФОСФОРИТ' in str(value('E7')).upper()
            and formula('E22')=='I22-5' and formula('I22')=='H4-2')


def _two_coat_reference(c):
    """The supplied two-coat act has header H17 but second-coat data in I."""
    def value(cell):return c.get(cell,{}).get('value')
    def formula(cell):return c.get(cell,{}).get('formula')
    standard=(str(value('E17'))=='1' and str(value('H17'))=='2'
              and value('J17') is None
              and formula('E22')=='I22-1' and formula('I22')=='H4-1')
    return (bool(re.fullmatch(r'\d+(?:/\d+)?/20\d\d(?:-[1-3])?\s+от\s*',str(value('E4'))))
            and (standard or _phosphorit_layout(value,formula))
            and all(formula(p)==f for p,f in (
                ('G8','E4'),('I8','H4'),('D49','H4'),('E30','E22'),('I30','I22')))
            and all(c.get(col+'27',{}).get('formula')==f'{col}28*100/L27'
                    or (c.get(col+'27',{}).get('formula') is None
                        and _numeric(c.get(col+'27',{}).get('value'),'thickness')>0)
                    for col in ('E','I'))
            and all(c.get(col+str(row),{}).get('value') is not None
                    or c.get(col+str(row),{}).get('formula') is not None
                    for col in ('E','I') for row in (28,31,32,33,34)))


def reference_coats(path):
    """Supported original layouts; the variant number never selects the count."""
    try:
        cells=_cells(path)
        if _two_coat_reference(cells):return 2
        if _three_coat_reference(cells):return 3
    except (Stop,TypeError,ValueError):pass
    if is_first_akz(path):return 1
    raise Stop('Не распознано расположение ячеек исходного акта АКЗ.')


def is_first_akz_appendix(path):
    try:
        c=_cells(path)
        title=TITLE.fullmatch(str(c.get('A1',{}).get('value') or ''))
        return (bool(title and title.group(1)==title.group(2))
                and str(c.get('A2',{}).get('value','')).strip()=='Отправочная марка'
                and str(c.get('B2',{}).get('value','')).strip()=='Наименование марки'
                and c.get('A4',{}).get('value') not in (None,'','Итого')
                and any(v.get('value')=='Итого' for p,v in c.items()
                        if re.fullmatch(r'A(?:[4-9]|[1-9]\d{1,3})',p)))
    except Exception:return False


def _day(v):
    if isinstance(v,datetime):v=v.date()
    if isinstance(v,date):return v
    try:return date.fromisoformat(str(v))
    except (ValueError,TypeError) as exc:
        raise Stop('АКЗ: отсутствует дата из реестра ВЗМК.') from exc


def _number(s):
    if not isinstance(s,str) or not NUMBER.fullmatch(s):
        raise Stop('АКЗ: неверный номер из реестра ВЗМК.')
    return s


def _numeric(x,label):
    try:v=Decimal(str(x))
    except Exception as exc:
        raise Stop('АКЗ: нечисловое поле '+label) from exc
    # Air/surface/dew may legitimately be below 0 C.
    # Humidity and physical coating thickness must remain positive.
    if (not v.is_finite() or
            label in ('thickness','humidity') and v<=0):
        raise Stop('АКЗ: нет измеренного значения '+label)
    return float(v)


def _check_formula(s,cell,formula):
    actual=str(s.Range(cell).Formula or '').replace('$','')
    if actual.upper()!=formula.upper():
        raise Stop('АКЗ: изменена исходная формула '+cell+
                   ' (ожидается '+formula+')')


def fill_first_akz(s,*,number,akz_date,climate):
    number=_number(number);day=_day(akz_date)
    if not isinstance(climate,dict) or climate.get('variant') not in range(4):
        raise Stop('АКЗ: не определён вид покрытия.')
    number=act_number(number,climate['variant'])
    coats=3 if all(s.Range(col+'17').Value2 in (i,str(i))
                   for i,col in enumerate(('E','H','J'),1)) else 1
    phosphorit=coats==1 and _phosphorit_layout(
        lambda cell:s.Range(cell).Value2,
        lambda cell:str(s.Range(cell).Formula or '').replace('$','').upper().removeprefix('='))
    if phosphorit or (coats==1 and s.Range('H17').Value2 in (2,'2')
            and str(s.Range('E22').Formula).replace('$','').upper()=='=I22-1'):
        coats=2
    columns=('E','I') if coats==2 else ('E','H','J')[:coats]
    intervals=climate.get('measurement_intervals',{}).get('layers',[])
    actual_coats=len(intervals) if intervals else len(climate.get('layers',[])) or 1
    if actual_coats!=coats:
        raise Stop('АКЗ: число измеренных слоёв не совпадает с исходным актом.')
    if intervals and coats==1 and climate.get('wet_thickness') is None:
        raise Stop('АКЗ: нет текущего измерения мокрой толщины первого слоя.')
    if not re.fullmatch(r'\d+(?:/\d+)?/20\d\d(?:-[1-3])?\s+от\s*',
                        str(s.Range('E4').Value2 or '')):
        raise Stop('АКЗ: изменён исходный номер E4.')
    formulas=[('G8','=E4'),('I8','=H4'),('D49','=H4')]
    formulas+=([('E22','=H4-1'),('E30','=E22')] if coats==1 else [
        ('E22','=I22-5' if phosphorit else '=I22-1'),
        ('I22','=H4-2' if phosphorit else '=H4-1'),
        ('E30','=E22'),('I30','=I22')]
        if coats==2 else [
        ('E22','=H22-1'),('H22','=J22-1'),('J22','=H4-1'),
        ('E30','=E22'),('H30','=H22'),('J30','=J22')])
    for cell,expected in formulas:
        _check_formula(s,cell,expected)
    from .measurement_ranges import verify_midrange_provenance
    verify_midrange_provenance(climate)
    # Validate the complete update BEFORE changing any cells. The original
    # reference's personnel, materials, batches and specifications stay as-is;
    # they are not new confirmations or input requirements.
    updates={}
    if coats in (2,3):
        from .facts import validate_dry_averages
        from .akz_layer_totals import verify_wet_layers
        dry=validate_dry_averages(climate.get('layers'),climate.get('dry_averages'))
        verify_wet_layers(climate)
        wet=climate.get('wet_averages')
        if not isinstance(wet,list) or len(wet)!=coats:
            raise Stop('АКЗ: нужны текущие мокрые измерения всех слоёв.')
        for col in (columns if coats==2 else ('H','J')):
            formula=str(s.Range(col+'27').Formula or '')
            if formula.startswith('='):
                _check_formula(s,col+'27','='+col+'28*100/L27')
        for i,col in enumerate(columns):
            updates[col+'28']=_numeric(dry[i]['average'],'thickness')
            updates[col+'27']=_numeric(wet[i]['average'],'thickness')
    else:
        updates['E28']=_numeric(climate.get('thickness'),'thickness')
        if climate.get('wet_thickness') is not None:
            updates['E27']=_numeric(climate['wet_thickness'],'thickness')
    for col in columns:
        for row,key in ((31,'surface'),(32,'air'),(33,'humidity'),(34,'dew')):
            updates[col+str(row)]=_numeric(climate.get(key),key)
    if coats==2 and not phosphorit:
        # These are the original unprinted climate helpers, not control
        # results. Replace their random climate formulas in the COPY too;
        # the global prohibition of random data remains unchanged.
        for col in ('L','M'):
            updates[col+'31']=_numeric(climate.get('surface'),'surface')
            updates[col+'33']=_numeric(climate.get('humidity'),'humidity')
    before_formula=s.Range('E27').Formula
    cell=s.Range('H4')
    before_format=(cell.NumberFormat,cell.NumberFormatLocal)
    s.Range('E4').Value2=number+' от '
    cell.Value2=(day-date(1899,12,30)).days
    if before_format!=(cell.NumberFormat,cell.NumberFormatLocal):
        raise Stop('АКЗ: изменён исходный формат даты.')
    for position,value in updates.items():s.Range(position).Value2=value
    if coats==2 and not phosphorit:
        # The supplied I27 has a RUB currency format. It is a thickness,
        # so use the first coat's numeric format in the new copy only.
        s.Range('I27').NumberFormat=s.Range('E27').NumberFormat
    if 'E27' not in updates and s.Range('E27').Formula!=before_formula:
        raise Stop('АКЗ: затёрта формула расчёта мокрого слоя.')
    return {'number':number,'date':day.isoformat(),
            'coats':coats,'updated_cells':sorted(['E4','H4',*updates]),
            'wet_film_formula_preserved':'E27' not in updates,
            'unchanged_reference_fields_are_not_current_control':True}


def _cellvalue(x):
    if isinstance(x,Decimal):
        if not x.is_finite():
            raise Stop('АКЗ: нечисловая масса/площадь.')
        return float(x)
    if type(x) not in (str,int,float):
        raise Stop('АКЗ: недопустимый тип значения ВОМ.')
    return x


def fill_first_akz_appendix(s,*,number,akz_date,rows,undated_draft=False):
    number=_number(number)
    day=None if akz_date is None and undated_draft else _day(akz_date)
    if not isinstance(rows,(list,tuple)) or not 1<=len(rows)<=2000:
        raise Stop('АКЗ: нет строк текущих изделий.')
    title=TITLE.fullmatch(str(s.Range('A1').Value2 or ''))
    if (not title or title.group(1)!=title.group(2)
            or str(s.Range('A2').Value2 or '').strip()!='Отправочная марка'
            or str(s.Range('B2').Value2 or '').strip()!='Наименование марки'):
        raise Stop('АКЗ: не совпадает историческое приложение.')
    first=4
    limit=min(int(s.UsedRange.Row+s.UsedRange.Rows.Count)+1,10000)
    totals=[i for i in range(first,limit+1)
            if str(s.Cells(i,1).Value2 or '').strip()=='Итого']
    if len(totals)!=1 or totals[0]<=first:
        raise Stop('АКЗ: в первом приложении нет строки «Итого».')
    old_total=totals[0];old_count=old_total-first
    count=len(rows);last=first+count-1;total=last+1
    if count>old_count:
        s.Rows(f'{old_total}:{old_total+count-old_count-1}').Insert()
    elif count<old_count:
        s.Rows(f'{first+count}:{old_total-1}').Delete()
    s.Range(f'A{first}:G{last}').FillDown()
    for i,row in enumerate(rows,first):
        if not isinstance(row,(tuple,list)) or len(row)!=7:
            raise Stop('АКЗ: неполная строка ВОМ.')
        if (not isinstance(row[0],str) or not row[0].strip()
                or not isinstance(row[1],str) or not row[1].strip()):
            raise Stop('АКЗ: нет марки/наименования ВОМ.')
        try:s.Cells(i,1).Hyperlinks.Delete()
        except Exception:pass
        s.Range(f'A{i}:G{i}').Value2=(tuple(map(_cellvalue,row)),)
    day_text=day.strftime('%d.%m.%Y') if day is not None else '__________'
    s.Range('A1').Value2=(f'Приложение № {number} к акту АКЗ № {number}  '
                          f'от {day_text} г.')
    if undated_draft:s.PageSetup.CenterHeader='ЧЕРНОВИК — НЕ ДЛЯ ВЫПУСКА'
    s.Cells(total,1).Value2='Итого'
    for col in ('C','E','G'):
        s.Range(f'{col}{total}').Formula=f'=SUM({col}3:{col}{last})'
    # Preserve the agreed VOM area precision in the printed control total.
    # Historic two-decimal formatting hid 0.004 m2 in the 104-mark shipment.
    s.Range(f'G{total}').NumberFormatLocal='0'+str(s.Application.DecimalSeparator)+'000'
    s.Range(f'A{first}:G{total}').Borders.LineStyle=1
    s.Range(f'A{total}:G{total}').Font.Bold=True
    s.Rows('2:3').RowHeight=28
    s.PageSetup.PrintArea='$A$1:$G$'+str(total)
    s.PageSetup.PrintTitleRows='$2:$3'
    s.PageSetup.Zoom=False
    s.PageSetup.FitToPagesWide=1
    s.PageSetup.FitToPagesTall=False
    s.Range('A1:G1').Interior.Pattern=-4142
    return {'number':number,'date':day.isoformat() if day is not None else None,
            'rows':count,'total_row':total}
