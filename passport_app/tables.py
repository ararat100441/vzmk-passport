from dataclasses import dataclass
from .config import PassportError


def norm(value):
    return ' '.join(str(value or '').lower().replace('ё', 'е').replace('\n', ' ').split()).strip(' .,')


@dataclass(frozen=True)
class Layout:
    sheet: str
    header_row: int
    start: int
    columns: dict
    total_row: int | None = None
    # True only for confirmed one-line shipping invoice form with grouped
    # headers E1:F1 and G1:H1, without a separate '1 шт / Всех' subrow.
    flat_shipping_header: bool = False


def _one(candidates, label):
    if len(candidates) != 1:
        raise PassportError(f'Неоднозначный или отсутствующий столбец «{label}»: {candidates}.')
    return candidates[0]



def _flat_shipping_columns(sheet,header,cells,cols,*,appendix):
    """Narrow compatibility for VZMK shipping VOM with one header row.

    E/F = one/total mass and G/H = one/total painting area.
    Allow ONLY the confirmed eight-column geometry; otherwise use the
    original two-tier-header parser and report its ordinary validation error.
    The strict appendix parser is intentionally unchanged.
    """
    if appendix or cols.get('mark')!=2 or cols.get('name')!=3 or cols.get('quantity')!=4:
        return None
    indexes={
        'mass':[c.column for c in cells
                if norm(c.value).startswith('масса (изделие)')],
        'area':[c.column for c in cells
                if norm(c.value).startswith('площадь окраски')],
    }
    if indexes != {'mass':[5],'area':[7]}:
        return None
    # In this approved invoice header both labels are merged HORIZONTALLY
    # on the same row. No subordinate heading row exists. Never guess
    # adjacent meanings from unrelated spreadsheet layouts.
    for key,start in (('mass',5),('area',7)):
        group=[m for m in sheet.merged_cells.ranges
               if (m.min_row==header==m.max_row
                   and m.min_col==start and m.max_col==start+1)]
        if len(group)!=1:
            raise PassportError('Отгрузочная ВОМ: неверная объединённая шапка '
                                +key+'; ожидается одна строка.')
        if start+1>len(cells) or norm(cells[start].value):
            raise PassportError('Отгрузочная ВОМ: рядом с заголовком '+key+
                                ' не должен быть другой столбец.')
    return {'mass_unit':5,'mass_total':6,'area_unit':7,'area_total':8}


def find_layout(sheet, *, appendix=False):
    aliases = {'mark': {'отправочная марка', 'марка', 'марка изделия'},
               'name': {'наименование марки', 'наименование', 'наименование изделия'}}
    candidates = []
    for row in sheet.iter_rows(max_row=min(sheet.max_row, 30)):
        if any(norm(c.value) in aliases['mark'] for c in row):
            candidates.append(row[0].row)
    header = _one(candidates, 'строка заголовков марок')
    cells = list(sheet[header])
    cols = {key: _one([c.column for c in cells if norm(c.value) in values], key)
            for key, values in aliases.items()}
    cols['quantity'] = _one([c.column for c in cells if norm(c.value).startswith('количество')], 'Количество')
    bottom = header
    for key in ('mark', 'name', 'quantity'):
        for merged in sheet.merged_cells.ranges:
            if merged.min_row <= header <= merged.max_row and merged.min_col <= cols[key] <= merged.max_col:
                bottom = max(bottom, merged.max_row)
    flat=_flat_shipping_columns(sheet,header,cells,cols,appendix=appendix)
    if flat is not None:
        cols.update(flat)
        totals = [row[0].row for row in sheet.iter_rows(min_row=header+1)
                  if any(norm(c.value) in {'итого','всего'} for c in row
                         if isinstance(c.value,str))]
        if len(totals)>1:
            raise PassportError('Неоднозначная строка «Итого» в отгрузочной ВОМ.')
        return Layout(sheet.title,header,header+1,cols,
                      totals[0] if totals else None,flat_shipping_header=True)

    for key, label in [('mass', 'масса'), ('area', 'площадь')]:
        group_col = _one([c.column for c in cells if norm(c.value).startswith(label)], label)
        group = [m for m in sheet.merged_cells.ranges
                 if m.min_row == header and m.min_col == group_col and m.max_col > group_col]
        if len(group) != 1:
            raise PassportError(f'Не найдена однозначная группа заголовков «{label}».')
        group = group[0]
        subrow = group.max_row + 1
        subcells = [sheet.cell(subrow, c) for c in range(group.min_col, group.max_col + 1)]
        cols[key + '_unit'] = _one([c.column for c in subcells if norm(c.value) in
                                   {'1 дет', '1 шт', '1 деталь', '1 изделие'}], label + ' 1 шт.')
        cols[key + '_total'] = _one([c.column for c in subcells if norm(c.value) in
                                    {'всех', 'всего', 'общая', 'общая масса', 'общая площадь'}], label + ' Всех')
        bottom = max(bottom, subrow)
    totals = [row[0].row for row in sheet.iter_rows(min_row=bottom + 1)
              if any(norm(c.value) in {'итого', 'всего'} for c in row if isinstance(c.value, str))]
    if len(totals) > 1 or (appendix and not totals):
        raise PassportError(f'Неоднозначная строка «Итого» на листе {sheet.title}.')
    return Layout(sheet.title, header, bottom + 1, cols, totals[0] if totals else None)


def locate_layout(book, *, appendix=False):
    matches, failures = [], []
    for sheet in book.worksheets:
        try:
            matches.append(find_layout(sheet, appendix=appendix))
        except PassportError as exc:
            failures.append(f'{sheet.title}: {exc}')
    if len(matches) != 1:
        raise PassportError('Не найдена единственная таблица марок. ' + '; '.join(failures))
    return matches[0]
