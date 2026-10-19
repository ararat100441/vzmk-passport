"""Last-used АКЗ UI defaults, isolated by project and certificate variant.

This is a convenience cache, NEVER primary evidence of quality control.
Only the GUI writes it after a successfully completed ACTUAL passport.
Do not store dates/measurements of new work as confirmed by this cache.
"""
from copy import deepcopy
from datetime import date, datetime, timezone
from pathlib import Path

from .facts import numeric, validate_layer_ranges
from .storage import Stop, local_root, read, save

SCHEMA = 'VZMK_AKZ_PROJECT_PRESETS_1'
CLIMATE_FIELDS = ('air', 'surface', 'humidity', 'dew')
LIMITS = {'air': (-60, 80), 'surface': (-60, 100),
          'humidity': (0, 100), 'dew': (-80, 80)}


def _project(name):
    if not isinstance(name, str) or not name.strip():
        raise Stop('Не выбран объект для хранения параметров АКЗ.')
    return name


def _file(directory=None):
    return Path(directory) / 'akz_presets.json' if directory is not None else local_root() / 'akz_presets.json'


def _read_all(directory=None):
    path = _file(directory)
    if not path.exists():
        return {'schema': SCHEMA, 'projects': {}}
    try:
        data = read(path)
    except (OSError, ValueError, TypeError) as exc:
        raise Stop('Повреждены настройки АКЗ; автоматическая подстановка отключена: ' + str(exc)) from exc
    if (not isinstance(data, dict) or data.get('schema') != SCHEMA
            or not isinstance(data.get('projects'), dict)):
        raise Stop('Неизвестный формат настроек АКЗ. Перезапись запрещена.')
    return data


def _normalize_climate(values):
    if not isinstance(values, dict):
        raise Stop('Некорректные сохранённые показатели климата.')
    result = {}
    for field in CLIMATE_FIELDS:
        minimum, maximum = LIMITS[field]
        if field not in values:
            raise Stop('Нет сохранённого показателя: ' + field)
        result[field] = format(numeric(values[field], field, minimum, maximum), 'f')
    return result


def last_for_project(project, directory=None):
    """Return defaults only for the selected object; no writes."""
    project = _project(project)
    data = _read_all(directory)['projects'].get(project)
    if data is None:
        return {'saved_at': '', 'variants': {}}
    if not isinstance(data, dict) or not isinstance(data.get('variants'), dict):
        raise Stop('Повреждены сохранённые показатели объекта: ' + project)
    result = {'saved_at': str(data.get('saved_at', '')), 'variants': {}}
    for key, values in data['variants'].items():
        if key not in ('0', '1', '2', '3') or not isinstance(values, dict):
            raise Stop('Повреждена запись варианта АКЗ в параметрах объекта.')
        item = {'climate': _normalize_climate(values.get('climate'))}
        if 'layers' in values:
            item['layers'] = validate_layer_ranges(values['layers'])
        item['last_akz_date'] = str(values.get('last_akz_date', ''))
        result['variants'][key] = item
    return deepcopy(result)


def remember(project, climates, akz_date, directory=None):
    """Save defaults from a successfully issued passport only.

    Caller must verify completed ACTUAL issuance before calling.
    Thickness of a prior measurement, protocol references, people, paint
    batches, and other factual fields are intentionally never persisted.
    """
    project = _project(project)
    if not isinstance(climates, list):
        raise Stop('Некорректный набор климата АКЗ.')
    if not climates:
        return False
    try:
        date.fromisoformat(akz_date)
    except (TypeError, ValueError):
        raise Stop('Некорректная дата последних параметров АКЗ.')
    candidate = {}
    for record in climates:
        if not isinstance(record, dict) or type(record.get('variant')) is not int:
            raise Stop('Некорректный вариант АКЗ.')
        variant = record['variant']
        if variant not in range(4) or str(variant) in candidate:
            raise Stop('Дублирующийся или недопустимый вариант АКЗ.')
        item = {'climate': _normalize_climate(record), 'last_akz_date': akz_date}
        if record.get('layers'):
            item['layers'] = validate_layer_ranges(record['layers'])
        candidate[str(variant)] = item

    data = _read_all(directory)
    project_data = data['projects'].get(project, {'variants': {}})
    if not isinstance(project_data, dict) or not isinstance(project_data.get('variants'), dict):
        raise Stop('Повреждена история АКЗ объекта. Перезапись запрещена.')
    project_data = deepcopy(project_data)
    for variant, item in candidate.items():
        # Replace the selected variant's previous defaults, including any old
        # layer ranges. Do not silently retain an older coating specification
        # if the latest issued document had none.
        project_data['variants'][variant] = item
    project_data['saved_at'] = datetime.now(timezone.utc).isoformat(timespec='seconds')
    data['projects'][project] = project_data
    save(_file(directory), data)
    return True
