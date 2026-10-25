"""Human-confirmed measured *range* => arithmetical midrange for VZMK forms.

The user explicitly chose (lower + upper) / 2 as the reported central value.
This is a MIDRANGE, not a statistical sample mean. Do not calculate it from
allowed specification limits or carry over last passport measurements.

A confirmed primary protocol still remains mandatory for issuance.
"""
from decimal import Decimal
from .facts import numeric
from .storage import Stop

FIELDS={'air':(Decimal('-60'),Decimal('80')),
        'surface':(Decimal('-60'),Decimal('100')),
        'humidity':(Decimal('0'),Decimal('100')),
        'dew':(Decimal('-80'),Decimal('80')),
        'thickness':(Decimal('0.001'),None)}


def midrange(low,high,label,minimum=None,maximum=None):
    a=numeric(low,label+' от',minimum,maximum)
    b=numeric(high,label+' до',minimum,maximum)
    if a>b:
        raise Stop('Начало диапазона превышает конец: '+label)
    return format((a+b)/Decimal('2'),'f')


def measurements_from_ranges(variant,inputs,*,coats=1):
    """All displayed climate values are based on current confirmed intervals."""
    if type(variant) is not int or variant not in range(4):
        raise Stop('Некорректный вариант АКЗ.')
    if type(coats) is not int or coats not in (1,2,3):
        raise Stop('Неподдерживаемое количество слоёв АКЗ.')
    result={'variant':variant,'midrange_method':'(от+до)/2',
            'midrange_source':'CURRENT_CONFIRMED_MEASURED_RANGE',
            'measurement_intervals':{}}
    for key in ('air','surface','humidity','dew'):
        raw=inputs.get(key)
        if not isinstance(raw,dict) or set(raw)!={'min','max'}:
            raise Stop('Введите текущий диапазон '+key+' от и до.')
        low,high=FIELDS[key]
        average=midrange(raw['min'],raw['max'],key,low,high)
        result[key]=average
        result['measurement_intervals'][key]=dict(raw)
    layer_intervals=inputs.get('layer_intervals')
    if not isinstance(layer_intervals,list) or len(layer_intervals)!=coats:
        raise Stop('Укажите текущие фактические диапазоны всех используемых слоёв.')
    sorted_layers={}
    total=Decimal(0)
    dry=[]
    for row in layer_intervals:
        if not isinstance(row,dict) or type(row.get('layer')) is not int:
            raise Stop('Неверный номер слоя в фактическом диапазоне.')
        i=row['layer']
        if i not in range(1,coats+1) or i in sorted_layers:
            raise Stop('Повторён или пропущен слой АКЗ.')
        a=numeric(row.get('min'),'фактическая толщина слоя от',Decimal('.001'))
        b=numeric(row.get('max'),'фактическая толщина слоя до',Decimal('.001'))
        if a>b:
            raise Stop('Толщина слоя от превышает до.')
        avg=(a+b)/2
        total+=avg
        sorted_layers[i]={'layer':i,'min':format(a,'f'),
                          'max':format(b,'f')}
        dry.append({'layer':i,'average':format(total,'f')})
    if set(sorted_layers)!=set(range(1,coats+1)):
        raise Stop('Не указаны все слои АКЗ.')
    result['thickness']=format(total,'f')
    result['measurement_intervals']['layers']=[
        sorted_layers[i] for i in range(1,coats+1)]
    if coats>1:
        # Legacy three-layer forms require 3 cumulative dry-coat averages.
        # Their measured intervals are distinguishable from allowed spec
        # limits and shall not be taken from reusable project presets.
        result['layers']=[sorted_layers[i] for i in range(1,coats+1)]
        result['dry_averages']=dry
    return result


def verify_midrange_provenance(climate):
    if 'measurement_intervals' not in climate:
        return True   # existing historical/older accepted form without range UI
    if climate.get('midrange_source')!='CURRENT_CONFIRMED_MEASURED_RANGE':
        raise Stop('Нет подтверждения источника диапазонов климата/толщины.')
    ints=climate['measurement_intervals']
    for key in ('air','surface','humidity','dew'):
        if key not in ints or key not in climate:
            raise Stop('Неполный фактический диапазон '+key)
        low,high=FIELDS[key]
        expected=midrange(ints[key]['min'],ints[key]['max'],key,low,high)
        if numeric(climate[key],key)!=numeric(expected,key):
            raise Stop('Середина диапазона изменена: '+key)
    intervals=ints.get('layers',[])
    if not intervals:
        raise Stop('Не указан фактический диапазон покрытия.')
    total=Decimal(0)
    for i,row in enumerate(intervals,1):
        if row['layer']!=i:
            raise Stop('Некорректная последовательность слоёв.')
        total+=Decimal(midrange(row['min'],row['max'],'слой '+str(i),
                                Decimal('.001')))
    if numeric(climate['thickness'],'толщина')!=total:
        raise Stop('Середина диапазона толщины была изменена.')
    return True
