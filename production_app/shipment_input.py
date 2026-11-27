"""Seven-column shipment input. Unknown area remains unknown, never zero."""
import csv
import io
from dataclasses import dataclass
from decimal import Decimal,InvalidOperation
from pathlib import Path
from .storage import Stop

HEADERS=('Отправочная марка','Наименование марки','Количество, шт.','Масса 1 шт., кг','Масса всего, кг','Площадь 1 шт., м²','Площадь всего, м²')
MARKER='VZMK_CLIPBOARD_VOM_1'
EXCEL_MAX=Decimal('9.99999999999999E+307')
EXCEL_MIN=Decimal('2.2250738585072E-308')

def numeric(value,label,*,blank=False):
    text=str('' if value is None else value).strip().replace('\u00a0','').replace(' ','').replace(',','.')
    if blank and not text:return None
    try:n=Decimal(text)
    except InvalidOperation as exc:raise Stop(label+': требуется число.') from exc
    if not n.is_finite() or n<0:raise Stop(label+': требуется неотрицательное число.')
    if n and not EXCEL_MIN<=n<=EXCEL_MAX:
        raise Stop(label+': число вне диапазона, поддерживаемого Excel.')
    return n

@dataclass(frozen=True)
class Shipment:
    rows:tuple
    @property
    def quantity(self):return sum((r[2] for r in self.rows),Decimal(0))
    @property
    def mass(self):return sum((r[4] for r in self.rows),Decimal(0))
    @property
    def area(self):return None if any(r[6] is None for r in self.rows) else sum((r[6] for r in self.rows),Decimal(0))
    @property
    def unique_marks(self):return len({r[0] for r in self.rows})
    def summary(self):
        return f'{len(self.rows)} строк / {self.unique_marks} уникальных марок / {self.quantity} изделий / {self.mass:.2f} кг / '+('площадь не указана' if self.area is None else f'{self.area:.3f} м²')

def normalize(rows):
    result=[]
    for no,raw in enumerate(rows,1):
        if not any(str(x or '').strip() for x in raw):continue
        if len(raw) not in (5,6,7):raise Stop(f'Строка {no}: ожидается 5, 6 или 7 колонок.')
        raw=list(raw)+['']*(7-len(raw));mark,name=(str(x or '').strip() for x in raw[:2])
        if not mark or not name or any('\x00' in x or x.startswith('=') for x in (mark,name)):
            raise Stop(f'Строка {no}: нужны марка и наименование без формул.')
        q,mu,mt=(numeric(raw[i],f'Строка {no}, {HEADERS[i]}') for i in (2,3,4))
        if q<=0 or q!=q.to_integral_value():raise Stop(f'Строка {no}: количество должно быть положительным целым.')
        if Decimal(format(q,'.15g'))!=q:raise Stop(f'Строка {no}: Excel не может сохранить количество без округления.')
        if abs(q*mu-mt)>Decimal('.011'):raise Stop(f'Строка {no}: масса всего не равна количеству × массе 1 шт. Проверьте исходные данные.')
        au,at=(numeric(raw[i],f'Строка {no}, площадь',blank=True) for i in (5,6))
        if au is None and at is not None:au=at/q
        elif at is None and au is not None:at=q*au
        elif at is not None and abs(q*au-at)>Decimal('.011'):raise Stop(f'Строка {no}: площадь всего не соответствует количеству × площади 1 шт.')
        for value in (au,at):
            if value is not None:numeric(value,f'Строка {no}, площадь')
        result.append((mark,name,q,mu,mt,au,at))
    if not result:raise Stop('ВОМ пуста: вставьте строки из Excel.')
    return Shipment(tuple(result))


def manual_row(mark,name,quantity,mass,area):
    q=numeric(quantity,HEADERS[2])
    m=numeric(mass,HEADERS[3])
    a=numeric(area,HEADERS[5],blank=True)
    return normalize([(mark,name,q,m,q*m,a,None if a is None else q*a)]).rows[0]

def parse_tsv(text):
    if not isinstance(text,str) or len(text)>5_000_000:raise Stop('Некорректный буфер ВОМ.')
    rows=[r for r in csv.reader(io.StringIO(text),delimiter='\t') if any(c.strip() for c in r)]
    if not rows:raise Stop('В буфере нет строк ВОМ.')
    # A mark label by itself is also a possible value in a data row.
    def is_header(row):
        labels = [cell.strip().casefold() for cell in row]
        offset = int(bool(labels) and 'лист' in labels[0])
        return (len(labels) >= offset + 3
                and labels[offset] in ('марка', 'отправочная марка')
                and labels[offset + 1].startswith('наименование')
                and labels[offset + 2].startswith(('количество', 'кол-во')))

    header = next((i for i, row in enumerate(rows[:5]) if is_header(row)), None)
    sheet_column=False
    if header is not None:
        r=rows[header];mark=next(i for i,c in enumerate(r) if c.strip().casefold() in ('марка','отправочная марка'))
        if mark==1 and len(r)==8 and 'лист' in r[0].casefold():sheet_column=True
        elif mark!=0:raise Stop('Неоднозначная шапка ВОМ: скопируйте B:H либо пять обязательных колонок.')
        if any(sum(bool(cell.strip()) for cell in row) >= 5 for row in rows[:header]):
            raise Stop('Строки данных перед заголовком ВОМ. Скопируйте таблицу с одной шапкой.')
        rows=rows[header+1:]
        # Excel merges the identifying columns vertically in a two-line header.
        data_start = 4 if sheet_column else 3
        if (rows and not any(cell.strip() for cell in rows[0][:data_start])
                and any(cell.strip().casefold() in ('всего', 'всех') for cell in rows[0][data_start:])):
            rows = rows[1:]
    output=[]
    for r in rows:
        if len(r)==8:
            if not sheet_column:
                # Without header the exact numeric-sheet / text-mark / text-name
                # / integer-quantity / numeric-mass signature must hold for ALL rows.
                if not r[0].strip().isdigit() or not r[1].strip() or not r[2].strip() or r[2].strip().replace(',','.').replace('.','',1).isdigit():
                    raise Stop('8 колонок не распознаны: выделите B:H или включите заголовок «Лист №».')
            r=r[1:]
        elif sheet_column:raise Stop('В строках ВОМ разное число колонок после заголовка A:H.')
        output.append(r)
    return normalize(output)

def write_internal(shipment,path):
    """Exclusive immutable input snapshot used by the existing XLSX engine."""
    from openpyxl import Workbook
    path=Path(path)
    if path.exists():raise Stop('Внутренняя копия ВОМ уже существует.')
    path.parent.mkdir(parents=True,exist_ok=True)
    b=Workbook();s=b.active;s.title='Отгрузочная ВОМ'
    s.append(['Отправочная марка','Наименование марки','Количество','Масса, кг',None,'Площадь, м²',None])
    for column in ('A','B','C'):s.merge_cells(column+'1:'+column+'2')
    s.merge_cells('D1:E1');s.merge_cells('F1:G1')
    s.cell(2,4,'1 шт.');s.cell(2,5,'Всего');s.cell(2,6,'1 шт.');s.cell(2,7,'Всего')
    for row in shipment.rows:s.append([float(v) if isinstance(v,Decimal) else v for v in row])
    s['J1']=MARKER
    try:
        with path.open('xb') as f:b.save(f)
    finally:b.close()
    return path
