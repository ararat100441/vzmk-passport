"""Preview is pure read-only; generation checks the same source fingerprints."""
from copy import deepcopy
from dataclasses import dataclass
from decimal import Decimal
import hashlib
import json
from pathlib import Path

from profiles_stage4.model import ProfileError, calculate_uzk, normalize_profile
from profiles_stage4.offline import no_network
from profiles_stage4.storage import load_catalog
from .inputs import (ROOT, REGISTRY, UZK, PROFILES, CHECKPOINT, Stage5Error, key,
                     mark_profiles, positive_integer, project_sources, registry_rows,
                     select_passport, sha, text, verify_pins, visible_vom)


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode('utf8')).hexdigest()


@dataclass(frozen=True)
class Preview:
    data: dict
    seal: str

    def verify(self):
        if digest(self.data) != self.seal:
            raise Stage5Error('PREVIEW_CHANGED: повторите предварительный расчёт')
        for path, expected in self.data['fingerprints'].items():
            if not Path(path).is_file() or sha(path) != expected:
                raise Stage5Error(f'INPUT_CHANGED: {path}. Повторите preview после проверки checkpoint.')


def calculate_marks(project, vom_rows, registry, mappings, catalog):
    quantities = {}
    for row in vom_rows:
        identity = key(row['mark'])
        if identity not in quantities:
            quantities[identity] = {'mark': row['mark'], 'quantity': 0, 'rows': []}
        quantities[identity]['quantity'] += positive_integer(row['quantity'], f'ВОМ строка {row["row"]}')
        quantities[identity]['rows'].append(row['row'])
    marks, excluded = [], []
    for identity, shipment in quantities.items():
        stock = registry.get((key(project), identity))
        if stock is None:
            excluded.append(shipment['mark'])
            continue
        mapping = mappings.get(identity)
        if mapping is None:
            raise Stage5Error(f'PROFILE_MISSING: {shipment["mark"]}. Заполните профиль через принятый Stage 3.')
        raw_profile = text(mapping['profile'], f'профиль марки {shipment["mark"]}')
        if key(raw_profile) == 'проверить':
            raise Stage5Error(f'PROFILE_UNRESOLVED: {shipment["mark"]} / ПРОВЕРИТЬ')
        lookup = catalog.lookup(raw_profile)
        if lookup['status'] not in ('FOUND', 'FOUND_PREFERRED'):
            raise Stage5Error(f'PROFILE_{lookup["status"]}: {shipment["mark"]} / {raw_profile}. Разрешите lookup в Stage 4.')
        try:
            single = calculate_uzk(lookup)
        except ProfileError as exc:
            raise Stage5Error(f'UZK_RULE_INVALID: {shipment["mark"]}: {exc}') from exc
        if single['status'] != 'CALCULATED' or not single['rows']:
            raise Stage5Error(f'UZK_RULE_NOT_CONFIGURED: {shipment["mark"]} / {raw_profile}. Требуется принятое правило Stage 4.')
        remaining = positive_integer(stock['remaining'], f'УЗК {shipment["mark"]}')
        quantity = min(shipment['quantity'], remaining)
        segments = [{**deepcopy(r), 'L': str(Decimal(str(r['L'])) * quantity)} for r in single['rows']]
        marks.append({'mark': shipment['mark'], 'catalog_mark': mapping['mark'], 'raw_profile': raw_profile,
                      'catalog_row': mapping['row'], 'normalized_lookup': normalize_profile(raw_profile),
                      'lookup': lookup, 'single': single, 'segments': segments,
                      'vom_quantity': shipment['quantity'], 'vom_rows': shipment['rows'],
                      'remaining': remaining, 'registry_rows': stock['rows'], 'Q': quantity})
    return marks, excluded


def product_paths():
    return [ROOT / 'stage5.py', ROOT / 'Запустить Этап 5 DEV.cmd', ROOT / 'requirements_stage5.txt',
            *sorted((ROOT / 'uzk_stage5').glob('*.py'))]


def preview(project, number, vom_path):
    with no_network() as attempts:
        verify_pins()
        passport = select_passport(project, number)
        project_root, catalog_path, template = project_sources(passport)
        vom_path = Path(vom_path).resolve()
        # The selected VOM must belong to this project or an isolated validation fixture.
        if not (vom_path.is_relative_to(project_root) or vom_path.is_relative_to(ROOT / 'validation')):
            raise Stage5Error('VOM_PATH_OUTSIDE_PROJECT: выберите ВОМ выбранного проекта')
        paths = [REGISTRY, UZK, PROFILES, CHECKPOINT, ROOT / 'settings.json', catalog_path, template, vom_path,
                 *(PROFILES.parent / name for name in ('profiles.user.json', 'profiles.catalog.json', 'profiles.build.json')),
                 *product_paths(), *sorted((ROOT / 'profiles_stage4').glob('*.py')), ROOT / 'passport_app/catalog.py']
        fingerprints = {str(p.resolve()): sha(p) for p in paths}
        vom = visible_vom(vom_path)
        expected_mass = Decimal(str(passport['mass']))
        if not vom['rows'] or abs(Decimal(vom['mass_kg']) - expected_mass) > Decimal('0.005'):
            raise Stage5Error(f'VOM_SHIPMENT_MISMATCH: видимая масса {vom["mass_kg"]}, DEV паспорт {expected_mass}')
        registry, registry_evidence = registry_rows()
        required_marks = {key(row['mark']) for row in vom['rows']
                          if (key(project), key(row['mark'])) in registry}
        try:
            # A valid empty intersection needs no profile mappings or rules.
            mappings = mark_profiles(catalog_path, required_marks=required_marks) if required_marks else {}
            catalog = load_catalog(PROFILES) if required_marks else None
            marks, excluded = calculate_marks(project, vom['rows'], registry, mappings, catalog)
        except ProfileError as exc:
            raise Stage5Error(f'STAGE4_CHECKPOINT_INVALID: {exc}') from exc
        data = {'kind': 'TEST/DRAFT', 'status': 'UZK_REQUIRED' if marks else 'NO_UZK_REQUIRED',
                'passport': passport, 'project_root': str(project_root),
                'template': str(template), 'mark_catalog': str(catalog_path), 'vom': vom,
                'registry_evidence': registry_evidence, 'marks': marks, 'excluded_without_registry': excluded,
                'controlled_units': sum(m['Q'] for m in marks), 'fingerprints': fingerprints,
                'network_attempts': list(attempts), 'registry_mutation': 'FORBIDDEN'}
        result = Preview(data, digest(data))
        result.verify()
        if attempts:
            raise Stage5Error('NETWORK_ATTEMPTED')
        return result


def generate(plan, output_dir=None):
    from .word import write_test_document
    from datetime import datetime
    from uuid import uuid4
    with no_network() as attempts:
        plan.verify()
        verify_pins()
        if not plan.data['marks']:
            return None  # no directory and no document when no UZK is needed
        if output_dir is None:
            output_dir = ROOT / 'DEV/Stage5_Выход' / (datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid4().hex[:8])
        output_dir = Path(output_dir).resolve()
        allowed = (ROOT / 'DEV/Stage5_Выход', ROOT / 'validation')
        if not any(output_dir.is_relative_to(p) and output_dir != p for p in allowed):
            raise Stage5Error('OUTPUT_PATH_FORBIDDEN: разрешён DEV/Stage5_Выход или отдельная validation папка')
        name = plan.data['passport']['number'].replace('/', '_') + ' - УЗК (ТЕСТ).docx'
        output = output_dir / name
        if output.exists():
            raise Stage5Error(f'OUTPUT_EXISTS: {output}. Выберите новую папку.')
        output_dir.mkdir(parents=True, exist_ok=True)
        created = False
        try:
            write_test_document(plan.data, output)
            created = True
            plan.verify()
            if attempts:
                raise Stage5Error('NETWORK_ATTEMPTED')
        except Exception:
            # Only our newly created draft; never a source or an existing document.
            if created:
                output.unlink(missing_ok=True)
            raise
        return output
