"""One shared, read-only VZMK ultrasonic inspection XLSM for all projects.

The shared workbook contains *remaining quantities*, NOT inspected weld
outcomes. It can populate the planned UZK quantity for current VOM marks,
but cannot assert GОДЕН, dates, actual weld length or personnel.

Contract from the approved Stage 5 project:
 sheet 'УЗК', first-row columns 'Заказ', 'Марка элемента', 'Кол-во (шт)'.
 Blank numeric cell -> 1; exact identical duplicate doesn't increase stock;
 conflicting duplicate blocks. Selected project's rows only.
"""
from decimal import Decimal, InvalidOperation
from pathlib import Path
from openpyxl import load_workbook
from .storage import Stop, sha, safe


HEADERS=('Заказ','Марка элемента','Кол-во (шт)')


def _key(raw):
    return ' '.join(str(raw or '').replace('ё','е').casefold().split())


def _qty(raw,row):
    if raw is None or isinstance(raw,str) and not raw.strip():
        return 1
    if isinstance(raw,bool) or isinstance(raw,str) and raw.startswith('='):
        raise Stop('Общий УЗК: неверное количество в строке '+str(row))
    try:
        q=Decimal(str(raw).strip().replace(',','.'))
    except (ValueError,InvalidOperation):
        raise Stop('Общий УЗК: нечисловое количество в строке '+str(row))
    if not q.is_finite() or q<=0 or q!=q.to_integral_value():
        raise Stop('Общий УЗК: количество должно быть положительным целым, '
                   'строка '+str(row))
    return int(q)


def find_shared_uzk(cfg,*,require=True):
    """A single explicitly selected file, or the unambiguous standard sibling.

    Its location does not depend on the currently selected project, and
    nothing is saved or written during discovery.
    """
    explicit=str(cfg.get('uzk_registry') or '').strip()
    if not explicit:
        registry=cfg.get('registry')
        if not registry:
            if require:raise Stop('Выберите реестр ВЗМК перед поиском общего УЗК.')
            return None
        candidate=Path(registry).parent/'УЗК.xlsm'
    else:
        candidate=Path(explicit)
    candidate=safe(candidate)
    if not candidate.is_file():
        if require:
            raise Stop('Общий реестр УЗК.xlsm не найден: '+str(candidate)+
                       '. Выберите его один раз в «Настройки путей».')
        return None
    if candidate.suffix.lower()!='.xlsm':
        raise Stop('Общий реестр УЗК должен быть файлом *.xlsm.')
    return candidate


def read_shared_uzk(path):
    """Read only values; reject formula quantities, duplicate contradictory rows."""
    path=safe(path)
    before=sha(path)
    book=load_workbook(path,read_only=True,data_only=False,keep_links=False)
    rows={}
    try:
        if 'УЗК' not in book.sheetnames:
            raise Stop('В общем XLSM отсутствует лист «УЗК».')
        sheet=book['УЗК']
        header=next(sheet.iter_rows(min_row=1,max_row=1,max_col=3,
                                    values_only=True),())
        if tuple(_key(h) for h in header)!=tuple(_key(h) for h in HEADERS):
            raise Stop('Неверные заголовки УЗК: ожидаются Заказ | Марка '
                       'элемента | Кол-во (шт).')
        for no,values in enumerate(
            sheet.iter_rows(min_row=2,max_col=3,values_only=True),start=2
        ):
            project,mark,qty=values
            if all(x is None or isinstance(x,str) and not x.strip()
                   for x in values):
                continue
            if (not isinstance(project,str) or not project.strip() or
                    not isinstance(mark,str) or not mark.strip() or
                    project.startswith('=') or mark.startswith('=')):
                raise Stop('Общий УЗК: заказ/марка отсутствует или неверен '
                           'в строке '+str(no))
            count=_qty(qty,no)
            key=(_key(project),_key(mark))
            if key in rows:
                if rows[key]['remaining']!=count:
                    raise Stop('Общий УЗК: противоречивые остатки у марки '+
                               mark+', строки '+str(rows[key]['rows'][0])+
                               ' и '+str(no))
                rows[key]['rows'].append(no)
            else:
                rows[key]={'project':project,'mark':mark,'remaining':count,
                           'rows':[no]}
    finally:
        book.close()
    if sha(path)!=before:
        raise Stop('Общий реестр УЗК изменился во время чтения.')
    return rows


def planned_control(path,project,vom):
    """Compute min(shipped, remaining) independently for every VOM mark.

    The first encounter determines display name; repeated same mark in VOM
    accumulates quantity. Never sum different marks or duplicated registry
    rows. Missing registry mark => no planned UZK, NOT a negative verdict.
    """
    rows=read_shared_uzk(path)
    shipped={}
    labels={}
    for row in vom.rows:
        mark=str(row[0])
        key=_key(mark)
        if not key:
            raise Stop('Марка ВОМ не определена.')
        quantity=_qty(row[2], 'ВОМ '+mark)
        shipped[key]=shipped.get(key,0)+quantity
        labels.setdefault(key,mark)
    output={}
    for key,quantity in shipped.items():
        rec=rows.get((_key(project),key))
        required=min(quantity,rec['remaining']) if rec else 0
        output[labels[key]]={
            'mark':labels[key],
            'vom_quantity':quantity,
            'registry_remaining':rec['remaining'] if rec else None,
            'uzk_quantity':required,
            'matched':bool(rec),
            'registry_rows':list(rec['rows']) if rec else [],
        }
    return output


def verify_planned_quantities(facts,plan):
    """Cannot change a computed quantity to skip selected current inspection."""
    expected={name:int(x['uzk_quantity']) for name,x in plan.items()}
    for row in facts.get('marks',[]):
        mark=str(row.get('mark',''))
        if mark not in expected:
            raise Stop('УЗК: неизвестная марка текущей ВОМ '+mark)
        observed=_qty(row.get('uzk_quantity'), 'ВОМ '+mark) if (
            row.get('uzk_quantity') not in (0,'0')
        ) else 0
        if observed!=expected[mark]:
            raise Stop('УЗК: количество по марке '+mark+' должно быть '+
                       str(expected[mark])+' согласно общему реестру. '
                       'Изменять его вручную для выпуска нельзя.')
    if {str(r.get('mark')) for r in facts.get('marks',[])}!=set(expected):
        raise Stop('УЗК: не все марки текущей ВОМ проверены.')
    return True
