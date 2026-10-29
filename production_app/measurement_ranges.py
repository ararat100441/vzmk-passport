"""Human-confirmed measured *range* => arithmetical midrange for VZMK forms.

The user explicitly chose (lower + upper) / 2 as the reported central value.
This is a MIDRANGE, not a statistical sample mean. Do not calculate it from
allowed specification limits or carry over last passport measurements.

A confirmed primary protocol still remains mandatory for issuance.
"""
from decimal import Decimal, ROUND_HALF_UP
from math import log, isfinite
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



def magnus_dew_point(air_c,relative_humidity):
    """Compute °C using Magnus formula 17.62 / 243.12 over water.

    The input humidity is a real CURRENT measured mean, never an estimated
    default. Round the displayed temperature to a tenth of a degree.
    """
    air=numeric(air_c,'температура воздуха',Decimal('-60'),Decimal('80'))
    rh=numeric(relative_humidity,'относительная влажность',Decimal('0'),Decimal('100'))
    if rh<=0:
        raise Stop('Влажность должна быть больше 0% для расчёта точки росы.')
    gamma=log(float(rh)/100)+17.62*float(air)/(243.12+float(air))
    result=243.12*gamma/(17.62-gamma)
    if not isfinite(result):
        raise Stop('Не удалось рассчитать точку росы.')
    return Decimal(str(result)).quantize(Decimal('0.1'),rounding=ROUND_HALF_UP)


def validate_temperature_order(climate):
    """User-approved policy air > surface; surface >= dew point + 3 °C."""
    air=numeric(climate.get('air'),'температура воздуха',Decimal('-60'),Decimal('80'))
    surface=numeric(climate.get('surface'),'температура поверхности',Decimal('-60'),Decimal('100'))
    dew=numeric(climate.get('dew'),'точка росы',Decimal('-100'),Decimal('80'))
    if air <= surface:
        raise Stop('АКЗ: температура воздуха должна быть строго выше '
                   'температуры поверхности. Проверьте диапазоны.')
    if surface < dew+Decimal('3'):
        raise Stop('АКЗ: температура поверхности должна быть не менее '
                   'чем на 3 °C выше точки росы.')
    return True


def measurements_from_ranges(variant,inputs,*,coats=1):
    """All displayed climate values are based on current confirmed intervals."""
    if type(variant) is not int or variant not in range(4):
        raise Stop('Некорректный вариант АКЗ.')
    if type(coats) is not int or coats not in (1,2,3):
        raise Stop('Неподдерживаемое количество слоёв АКЗ.')
    result={'variant':variant,'midrange_method':'(от+до)/2',
            'midrange_source':'CURRENT_CONFIRMED_MEASURED_RANGE',
            'measurement_intervals':{}}
    for key in ('air','surface','humidity'):
        raw=inputs.get(key)
        if not isinstance(raw,dict) or set(raw)!={'min','max'}:
            raise Stop('Введите текущий диапазон '+key+' от и до.')
        low,high=FIELDS[key]
        average=midrange(raw['min'],raw['max'],key,low,high)
        result[key]=average
        result['measurement_intervals'][key]=dict(raw)
    # The user no longer enters dew point: Magnus derives it from confirmed
    # mean air temperature and mean relative humidity.
    result['dew']=format(magnus_dew_point(result['air'],result['humidity']),'f')
    result['dew_method']='MAGNUS_17.62_243.12'
    validate_temperature_order(result)
    layer_intervals=inputs.get('layer_intervals')
    if not isinstance(layer_intervals,list) or len(layer_intervals)!=coats:
        raise Stop('Укажите текущие фактические диапазоны всех используемых слоёв.')
    sorted_layers={}
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
        sorted_layers[i]={'layer':i,'min':format(a,'f'),
                          'max':format(b,'f')}
    if set(sorted_layers)!=set(range(1,coats+1)):
        raise Stop('Не указаны все слои АКЗ.')
    total=Decimal(0)
    dry=[]
    for i in range(1,coats+1):
        row=sorted_layers[i]
        total+=(Decimal(row['min'])+Decimal(row['max']))/2
        dry.append({'layer':i,'average':format(total,'f')})
    result['thickness']=format(total,'f')
    if coats==1 and inputs.get('wet'):
        raw=inputs['wet']
        result['wet_thickness']=midrange(raw.get('min'),raw.get('max'),
                                        'мокрый слой',Decimal('.001'))
        result['measurement_intervals']['wet']=dict(raw)
    result['measurement_intervals']['layers']=[
        sorted_layers[i] for i in range(1,coats+1)]
    if coats==2:
        raise Stop('Двухслойный акт пока не поддерживает расчёт по диапазонам; '
                   'требуется утверждённая форма и методика для двух слоёв.')
    if coats==3:
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
    for key in ('air','surface','humidity'):
        if key not in ints or key not in climate:
            raise Stop('Неполный фактический диапазон '+key)
        low,high=FIELDS[key]
        expected=midrange(ints[key]['min'],ints[key]['max'],key,low,high)
        if numeric(climate[key],key)!=numeric(expected,key):
            raise Stop('Середина диапазона изменена: '+key)
    if 'dew' in ints or climate.get('dew_method')!='MAGNUS_17.62_243.12':
        raise Stop('Точка росы должна быть рассчитана по Магнусу, '
                   'а не введена вручную.')
    expected_dew=magnus_dew_point(climate['air'],climate['humidity'])
    if numeric(climate.get('dew'),'точка росы')!=expected_dew:
        raise Stop('Расчётная точка росы была изменена.')
    validate_temperature_order(climate)
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
    if 'wet' in ints:
        expected=midrange(ints['wet']['min'],ints['wet']['max'],'мокрый слой',Decimal('.001'))
        if numeric(climate.get('wet_thickness'),'мокрый слой')!=Decimal(expected):
            raise Stop('Середина диапазона мокрого слоя была изменена.')
    return True
