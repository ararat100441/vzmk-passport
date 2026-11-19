"""Generate and reuse numeric previews of the entered AKZ ranges."""
from decimal import Decimal, ROUND_CEILING, ROUND_FLOOR
from random import Random

from .facts import numeric
from .measurement_ranges import DRAFT_SOURCE, FIELDS, magnus_dew_point
from .storage import Stop, digest

def interval_fields(inputs):
    """Index air, surface, humidity and separate wet/dry intervals by UI field."""
    fields = {key: inputs[key] for key in ('air', 'surface', 'humidity', 'wet')
              if isinstance(inputs.get(key), dict)}
    for source, prefix in (('layer_intervals', 'layer_'), ('wet_intervals', 'wet_')):
        for row in inputs.get(source) or ():
            if isinstance(row, dict) and row.get('layer') in (1, 2, 3):
                fields[prefix + str(row['layer'])] = row
    return fields


def sample_interval(key, pair, previous, scope):
    if any(pair.get(edge) is None or not str(pair[edge]).strip() for edge in ('min', 'max')):
        return None
    bounds = FIELDS.get(key, FIELDS['thickness'])
    low, high = [numeric(pair[edge], key, *bounds) for edge in ('min', 'max')]
    if low > high:
        raise Stop('Начало диапазона превышает конец: ' + key)
    step = Decimal('0.1') if key in ('air', 'surface') else Decimal(1)
    first = int((low / step).to_integral_value(rounding=ROUND_CEILING))
    last = int((high / step).to_integral_value(rounding=ROUND_FLOOR))
    if first > last:
        # A narrow interval can contain no integer or tenth. Preserve the
        # precision of its entered endpoints in that case.
        step = Decimal(1).scaleb(min(low.as_tuple().exponent, high.as_tuple().exponent))
        first, last = int(low / step), int(high / step)
    bounds = {'min': format(low.normalize(), 'f'), 'max': format(high.normalize(), 'f')}
    if isinstance(previous, dict) and all(previous.get(k) == v for k, v in bounds.items()):
        try:
            old = numeric(previous.get('value'), key, low, high)
            if old / step == (old / step).to_integral_value():
                return previous
        except Stop:
            pass
    rng = Random(digest([scope, key, bounds]))
    value = format(step * rng.randint(first, last), 'f')
    return {**bounds, 'value': value}


def preview_values(inputs, previous, scope, *, strict=False):
    result = {}
    for key, pair in interval_fields(inputs).items():
        try:
            sample = sample_interval(key, pair, previous.get(key), scope)
        except Stop:
            if strict:
                raise
            continue
        if sample is not None:
            result[key] = sample
    return result


def draft_climate(samples, coats):
    values = {key: sample['value'] for key, sample in samples.items()}
    climate = {k: values[k] for k in ('air', 'surface', 'humidity') if k in values}
    if 'air' in climate and 'humidity' in climate and Decimal(climate['humidity']) > 0:
        climate['dew'] = format(magnus_dew_point(climate['air'], climate['humidity']), 'f')
    if coats == 1 and 'wet' in values:
        climate['wet_thickness'] = values['wet']
    if coats in (1, 2, 3):
        dry, wet = Decimal(0), Decimal(0)
        dry_rows, wet_rows = [], []
        for i in range(1, coats + 1):
            key = f'layer_{i}'
            if key in values and len(dry_rows) == i - 1:
                dry += Decimal(values[key])
                dry_rows.append({'layer': i, 'average': format(dry, 'f')})
            key = f'wet_{i}'
            if key in values and len(wet_rows) == i - 1:
                wet += Decimal(values[key])
                wet_rows.append({'layer': i, 'average': values[key], 'cumulative': format(wet, 'f')})
        if len(dry_rows) == coats:
            climate['thickness'] = format(dry, 'f')
            if coats > 1:
                climate['dry_averages'] = dry_rows
        if coats > 1 and len(wet_rows) == coats:
            climate.update(wet_averages=wet_rows, wet_total=format(wet, 'f'))
    if climate:
        climate['midrange_source'] = DRAFT_SOURCE
    return climate

