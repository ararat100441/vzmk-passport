"""Verified wet-coat inputs; no automatic coating-quality verdict.

Wet inputs represent EACH applied coat. Their arithmetic sum is an informational
total, not an independently measured cumulative wet film. Existing 3-coat
Excel formulas based on dry thickness and solids content are not overwritten.
"""
from copy import deepcopy
from decimal import Decimal, InvalidOperation

from .storage import Stop


def _positive(raw, label):
    if raw is None or isinstance(raw, bool):
        raise Stop('Нет фактического измерения: '+label)
    try:
        value=Decimal(str(raw).strip().replace(',','.'))
    except (InvalidOperation, ValueError):
        raise Stop('Некорректная толщина: '+label)
    if not value.is_finite() or value<=0:
        raise Stop('Толщина должна быть положительной: '+label)
    return value


def calculate_wet_layers(intervals, coats):
    """Return canonical intervals, coat midranges, cumulative arithmetic sum."""
    if type(coats) is not int or coats not in (2,3):
        raise Stop('Нужно 2 или 3 слоя АКЗ.')
    if not isinstance(intervals, list) or len(intervals)!=coats:
        raise Stop('Укажите мокрую толщину «От/До» каждого слоя АКЗ.')
    indexed={}
    for row in intervals:
        if not isinstance(row, dict) or type(row.get('layer')) is not int:
            raise Stop('Некорректный номер мокрого слоя АКЗ.')
        i=row['layer']
        if i not in range(1,coats+1) or i in indexed:
            raise Stop('Повторён или отсутствует номер мокрого слоя АКЗ.')
        low=_positive(row.get('min'),'мокрый слой '+str(i)+' — От')
        high=_positive(row.get('max'),'мокрый слой '+str(i)+' — До')
        if low>high:
            raise Stop('Мокрый слой '+str(i)+': «От» превышает «До».')
        indexed[i]={'layer':i,'min':format(low,'f'),'max':format(high,'f')}
    if set(indexed)!=set(range(1,coats+1)):
        raise Stop('Не указаны мокрые измерения всех слоёв.')
    canonical=[];averages=[];total=Decimal(0)
    for i in range(1,coats+1):
        row=indexed[i]
        current=(Decimal(row['min'])+Decimal(row['max']))/2
        total+=current
        canonical.append(row)
        averages.append({'layer':i,'average':format(current,'f'),
                         'cumulative':format(total,'f')})
    return canonical,averages,format(total,'f')


def with_wet_layers(climate, intervals, coats):
    """Attach verified current wet-coat ranges without altering dry readings."""
    if not isinstance(climate,dict) or not isinstance(climate.get('measurement_intervals'),dict):
        raise Stop('Нет подтверждённых диапазонов АКЗ.')
    if len(climate['measurement_intervals'].get('layers',[]))!=coats:
        raise Stop('Количество мокрых и сухих слоёв АКЗ не совпадает.')
    canonical,averages,total=calculate_wet_layers(intervals,coats)
    result=deepcopy(climate)
    result['wet_averages']=averages
    result['wet_total']=total
    result['measurement_intervals']['wet_layers']=canonical
    return result


def verify_wet_layers(climate):
    """Reject changes to recorded wet values and their calculated totals."""
    if not isinstance(climate,dict):
        raise Stop('Повреждён набор измерений АКЗ.')
    intervals=climate.get('measurement_intervals')
    if not isinstance(intervals,dict) or 'wet_layers' not in intervals:
        if 'wet_averages' in climate or 'wet_total' in climate:
            raise Stop('Отсутствуют исходные мокрые измерения АКЗ.')
        return True  # Prior accepted one-coat or three-coat payload.
    coats=len(intervals.get('layers',[]))
    normalized,averages,total=calculate_wet_layers(intervals['wet_layers'],coats)
    observed=climate.get('wet_averages')
    if (intervals['wet_layers']!=normalized or
            not isinstance(observed,list) or len(observed)!=coats):
        raise Stop('Изменены диапазоны мокрой толщины АКЗ.')
    for old,expected in zip(observed,averages):
        if (not isinstance(old,dict) or
                type(old.get('layer')) is not int or
                old['layer']!=expected['layer'] or
                _positive(old.get('average'),'средняя мокрая толщина')!=
                Decimal(expected['average']) or
                _positive(old.get('cumulative'),'сумма мокрых толщин')!=
                Decimal(expected['cumulative'])):
            raise Stop('Изменена рассчитанная мокрая толщина АКЗ.')
    if _positive(climate.get('wet_total'),'итог мокрых толщин')!=Decimal(total):
        raise Stop('Изменена сумма мокрых толщин АКЗ.')
    return True
