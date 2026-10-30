"""Exact first AKZ #0 XLSX and appendix (read only until copied).

In new r4 copies E27 is the independently measured wet thickness.
Legacy calls without wet input retain the reference formula (never certify it).
Registry supplies number/date; approved current climate supplies E28/E31:E34.
Current visible VOM supplies all seven columns of every appendix body row.
Existing cell formatting, measurement formulas and personnel stay intact.
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
        return (bool(re.fullmatch(r'\d+/20\d\d(?:-[1-3])? от ',
                      str(c.get('E4',{}).get('value',''))))
                and all(c.get(p,{}).get('formula')==f for p,f in (
                    ('G8','E4'),('I8','H4'),('E22','H4-1'),
                    ('E30','E22'),('D49','H4')))
                and (c.get('E27',{}).get('formula')=='E28/0.6' or
                     _numeric(c.get('E27',{}).get('value'),'thickness')>0)
                and all(c.get(p,{}).get('value') is not None for p in
                    ('H4','E28','E31','E32','E33','E34')))
    except Exception:return False


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
    if climate.get('layers') or climate.get('layer_details'):
        raise Stop('АКЗ: первый бланк только для одного слоя.')
    if not re.fullmatch(r'\d+/20\d\d(?:-[1-3])? от ',
                        str(s.Range('E4').Value2 or '')):
        raise Stop('АКЗ: изменён исходный номер E4.')
    for cell,expected in (
        ('G8','=E4'),('I8','=H4'),('E22','=H4-1'),
        ('E30','=E22'),('D49','=H4')
    ):
        _check_formula(s,cell,expected)
    before_formula=s.Range('E27').Formula
    cell=s.Range('H4')
    before_format=(cell.NumberFormat,cell.NumberFormatLocal)
    s.Range('E4').Value2=number+' от '
    cell.Value2=(day-date(1899,12,30)).days
    if before_format!=(cell.NumberFormat,cell.NumberFormatLocal):
        raise Stop('АКЗ: изменён исходный формат даты.')
    for position,key in (
        ('E28','thickness'),('E31','surface'),('E32','air'),
        ('E33','humidity'),('E34','dew')
    ):
        s.Range(position).Value2=_numeric(climate.get(key),key)
    if climate.get('wet_thickness') is not None:
        s.Range('E27').Value2=_numeric(climate['wet_thickness'],'thickness')
    elif s.Range('E27').Formula!=before_formula:
        raise Stop('АКЗ: затёрта формула расчёта мокрого слоя.')
    for area in (
        'E4:G4','H4:J4','E27:K27','E28:K28',
        'E31:K31','E32:K32','E33:K33','E34:K34'
    ):
        s.Range(area).Interior.Pattern=-4142
    return {'number':number,'date':day.isoformat(),
            'wet_film_formula_preserved':climate.get('wet_thickness') is None}


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
