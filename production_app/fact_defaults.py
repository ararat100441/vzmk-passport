"""Last-used engineering inputs and saved previews, local to one project."""
from copy import deepcopy
from pathlib import Path

from .storage import Stop, digest, local_root, mutex, read, save

SCHEMA = 'VZMK_FACTS_DEFAULTS_1'


def _path(project, training, directory):
    mode = 'training' if training else 'actual'
    root = Path(directory) if directory is not None else local_root() / 'fact-defaults'
    return root / mode / (digest(project) + '.json')


def load_defaults(project, training, directory=None):
    path = _path(project, training, directory)
    try:
        data = read(path)
    except FileNotFoundError:
        return {'schema': SCHEMA, 'project': project, 'marks': {}, 'ranges': {}, 'drafts': {}}
    except (ValueError, OSError) as exc:
        raise Stop('Не удалось загрузить настройки фактических результатов объекта.') from exc
    if (not isinstance(data, dict) or data.get('schema') != SCHEMA
            or data.get('project') != project
            or not all(isinstance(data.get(key), dict) for key in ('marks', 'ranges', 'drafts'))
            or not all(isinstance(row, dict) for row in data['marks'].values())
            or not all(isinstance(row, dict) for row in data['ranges'].values())
            or not all(isinstance(row, dict) and isinstance(row.get('ranges'), dict)
                       and isinstance(row.get('preview'), dict) for row in data['drafts'].values())):
        raise Stop('Неверный формат настроек фактических результатов объекта.')
    return data


def seed_facts(defaults, scope):
    draft = defaults['drafts'].get(scope, {})
    ranges = deepcopy(draft.get('ranges', defaults['ranges']))
    return {'facts_scope': scope,
            'marks': [{'mark': mark, **deepcopy(values)} for mark, values in defaults['marks'].items()],
            'akz_drafts': ranges,
            'akz_preview': deepcopy(draft.get('preview', {}))}


def remember_defaults(project, training, scope, facts, directory=None):
    """Merge under a local lock. Only whitelisted fields may cross shipments."""
    path = _path(project, training, directory)
    with mutex(path):
        defaults = load_defaults(project, training, directory)
        for row in facts['marks']:
            defaults['marks'][row['mark']] = {
                key: row.get(key, '') for key in ('profile', 'akz', 'standard')
            }
        ranges = {}
        for variant, inputs in facts['akz_drafts'].items():
            ranges[variant] = {key: value for key, value in inputs.items()
                               if key in ('coats', 'air', 'surface', 'humidity', 'wet',
                                          'layer_intervals', 'wet_intervals')}
        defaults['ranges'].update(ranges)
        defaults['drafts'][scope] = {'ranges': ranges, 'preview': facts.get('akz_preview', {})}
        save(path, defaults)

