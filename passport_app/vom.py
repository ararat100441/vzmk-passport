from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from .config import PassportError
from .excel import load_book, sha256
from .tables import locate_layout


FIELDS = ('mark', 'name', 'quantity', 'mass_unit', 'mass_total', 'area_unit', 'area_total')


def number(value, coordinate):
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise PassportError(f'Нет числового значения в {coordinate}: {value!r}. Для формул сохраните пересчёт ВОМ в Excel.')
    result = Decimal(str(value))
    if not result.is_finite() or result < 0:
        raise PassportError(f'Недопустимое число в {coordinate}.')
    return result


def validate_flat_shipping_amounts(quantity,mass_unit,mass_total,area_unit,area_total,row):
    """Cross-check inferred E/F/G/H fields in a one-line shipping invoice.

    The layout is not labeled with second-row subheaders. Its meanings are
    accepted only when each total agrees with quantity × per-item value.
    """
    if quantity<=0:
        raise PassportError('Количество в отгрузочной ВОМ должно быть больше нуля: '
                            +str(row))
    if abs(quantity*mass_unit-mass_total)>Decimal('0.011'):
        raise PassportError('Масса всей строки не равна количество × масса 1 шт: '
                            +str(row))
    if abs(quantity*area_unit-area_total)>Decimal('0.011'):
        raise PassportError('Площадь всей строки не равна количество × площадь 1 шт: '
                            +str(row))


@dataclass(frozen=True)
class Vom:
    path: Path
    digest: str
    layout: object
    rows: tuple
    source_rows: tuple
    quantity: Decimal
    mass: Decimal
    area: Decimal
    connections: str


def read_vom(path):
    path = Path(path).resolve()
    before = sha256(path)
    formulas = load_book(path)
    values = load_book(path, values=True)
    try:
        layout = locate_layout(formulas)
        sheet = values[layout.sheet]
        rows, source_rows, names = [], [], []
        stop = layout.total_row or sheet.max_row + 1
        for r in range(layout.start, stop):
            cells = [sheet.cell(r, layout.columns[f]) for f in FIELDS]
            if all(c.value is None for c in cells):
                continue
            mark, name = cells[0].value, cells[1].value
            if not isinstance(mark, (str, int)) or not str(mark).strip():
                raise PassportError(f'Не найдена марка в строке {r} ВОМ.')
            if name is not None and not isinstance(name, str):
                raise PassportError(f'Наименование марки должно быть текстом: строка {r}.')
            nums = [number(c.value, c.coordinate) for c in cells[2:]]
            if nums[0] != nums[0].to_integral_value():
                raise PassportError(f'Количество должно быть целым неотрицательным: строка {r}.')
            if getattr(layout,'flat_shipping_header',False):
                validate_flat_shipping_amounts(*nums,r)
            rows.append((mark, name, *nums))
            source_rows.append(r)
            trimmed = (name or '').strip()
            if trimmed and trimmed not in names:
                names.append(trimmed)
        if not rows or not names:
            raise PassportError('ВОМ не содержит строк марок или непустых наименований для «Соединения».')
        sums = [sum((row[c] for row in rows), Decimal(0)) for c in (2, 4, 6)]
        if layout.total_row:
            for field, expected, tolerance in zip(('quantity', 'mass_total', 'area_total'), sums,
                                                   (Decimal(0), Decimal('0.01'), Decimal('0.001'))):
                cell = sheet.cell(layout.total_row, layout.columns[field])
                if abs(number(cell.value, cell.coordinate) - expected) > tolerance:
                    raise PassportError(f'Сумма строк не совпадает с «Итого» ВОМ: {field}.')
        if sha256(path) != before:
            raise PassportError('ВОМ изменился во время чтения. Повторите проверку.')
        return Vom(path, before, layout, tuple(rows), tuple(source_rows), *sums, ', '.join(names))
    finally:
        formulas.close()
        values.close()
