"""Explicit facts for historical VZMK Word forms; never copy issued controls.

Word forms can use the additional {{FIELD}} markers below while retaining their
original layout. The operator confirms values for the *current* shipment.
"""
import csv
import re
from datetime import date
from decimal import Decimal
from pathlib import Path

from .storage import Stop
from .facts import numeric

DOC_FIELD_LABELS = {
    # The existing two-act VIK template.
    'VIK_WELDERS': 'ВИК — сварщики, удостоверения и клейма',
    'VIK_CONNECTION_TYPES': 'ВИК — типы и категории соединений',
    'VIK_WELD_METHOD': 'ВИК — способ и положение сварки',
    'VIK_REGULATIONS': 'ВИК — нормативная документация',
    'VIK_DEFECTS': 'ВИК — выявленные дефекты',
    'VIK_CONCLUSION': 'ВИК — подтверждённое заключение',
    'VIK_INSPECTOR': 'ВИК — исполнитель, квалификация и удостоверение',
    'VIK_HEAD': 'ВИК — руководитель, квалификация и удостоверение',
    'VIK_PROCESS_CONNECTIONS': 'ВИК (в процессе сварки) — соединения',
    'VIK_PROCESS_REGULATIONS': 'ВИК (в процессе сварки) — НД и технологические карты',
    'VIK_PROCESS_CONCLUSION': 'ВИК (в процессе сварки) — заключение',
    'UZK_EQUIPMENT': 'УЗК — дефектоскоп и поверка',
    'UZK_REFERENCE_BLOCKS': 'УЗК — контрольные образцы и поверка',
    'UZK_INSPECTOR': 'УЗК — исполнитель, квалификация и удостоверение',
    'QUALITY_ORDER': 'Документ о качестве — номер заказа',
    'QUALITY_MANUFACTURER': 'Документ о качестве — предприятие-изготовитель',
    'QUALITY_ADDRESS': 'Документ о качестве — адрес предприятия',
    'QUALITY_STANDARDS': 'Документ о качестве — нормы изготовления',
    'QUALITY_WELD_STANDARDS': 'Документ о качестве — нормы сварки',
    'QUALITY_ELECTRODES': 'Документ о качестве — электроды',
    'QUALITY_WIRE': 'Документ о качестве — сварочная проволока',
    'QUALITY_FLUX': 'Документ о качестве — флюс',
    'QUALITY_GAS': 'Документ о качестве — защитный газ',
    'QUALITY_ZINC': 'Документ о качестве — горячее цинкование',
    'QUALITY_PRIMER': 'Документ о качестве — грунтовка',
    'QUALITY_TOPCOAT': 'Документ о качестве — финишное покрытие',
    'QUALITY_PREPARATION': 'Документ о качестве — подготовка поверхности',
    'QUALITY_WELDER_CERTIFICATES': 'Документ о качестве — сведения об аттестации сварщиков',
    'QUALITY_SIGNER': 'Документ о качестве — подписант и должность',
    'QUALITY_CONCLUSION': 'Документ о качестве — подтверждённое заключение',
}

UZK_COLUMNS = (
    'mark', 'joint', 'welders', 'date', 'frequency_mhz',
    'angle_deg', 'defects', 'conclusion', 'basis'
)
UZK_HEADERS = (
    'Марка', 'Характеристика сварного соединения',
    'Ф.И.О. и клеймо сварщика', 'Дата контроля',
    'Частота, МГц', 'Угол ввода, °', 'Дефекты и оценка',
    'Заключение по стыку', 'Первичный протокол',
)
UZK_CSV_HEADER = ';'.join(UZK_COLUMNS)
UZK_ALIASES = {
    'марка': 'mark', 'отправочная марка': 'mark',
    'характеристика сварного соединения': 'joint',
    'сварное соединение': 'joint',
    'фио и клеймо сварщика': 'welders',
    'ф.и.о. и клеймо сварщика': 'welders',
    'дата контроля': 'date',
    'частота мгц': 'frequency_mhz',
    'частота, мгц': 'frequency_mhz',
    'угол ввода луча': 'angle_deg',
    'угол ввода луча градусы': 'angle_deg',
    'дефекты': 'defects', 'заключение': 'conclusion',
    'первичный протокол': 'basis',
}


def normalize_doc_fields(values):
    if values is None:
        return {}
    if not isinstance(values, dict):
        raise Stop('Дополнительные реквизиты документов должны быть словарём.')
    result = {}
    for key, value in values.items():
        if key not in DOC_FIELD_LABELS:
            raise Stop('Неподдержанное дополнительное поле: ' + str(key))
        if not isinstance(value, str):
            raise Stop('Не текстовое значение реквизита: ' + key)
        value = value.strip()
        if len(value) > 5000 or '{{' in value or '}}' in value or '\x00' in value:
            raise Stop('Недопустимое значение реквизита: ' + key)
        if value:
            result[key] = value
    return result


def _frequency_parts(value, label):
    tokens = re.split(r'[;|/]+', str(value).strip())
    if not tokens or any(not token.strip() for token in tokens):
        raise Stop('Неверная частота или угол УЗК: ' + label)
    return [numeric(token.strip(), label, Decimal('0.1'), Decimal('90'))
            for token in tokens]


def validate_uzk_details(facts, vom, *, required=False):
    """Check observed weld records, preserving repeats and sequence from input.

    Does not infer weld length, depth, probe frequency, angle, personnel,
    defect verdict or primary test reports from VOM or former passports.
    """
    details = facts.get('uzk_details')
    if details in (None, []):
        if required:
            raise Stop('Для выбранной подробной формы УЗК укажите реальные строки соединений.')
        return []
    if not isinstance(details, list) or len(details) > 10000:
        raise Stop('Некорректный список строк УЗК.')
    available = {str(row[0]) for row in vom.rows}
    start = date.fromisoformat(facts['start_date'])
    end = date.fromisoformat(facts['date'])
    confirmed = date.fromisoformat(facts['uzk_date'])
    output = []
    for index, row in enumerate(details, 1):
        if not isinstance(row, dict) or set(row) - set(UZK_COLUMNS):
            raise Stop('Неверные поля строки УЗК №' + str(index))
        record = {}
        for key in UZK_COLUMNS:
            value = row.get(key, '')
            if not isinstance(value, str) or not value.strip() or '{{' in value or '}}' in value:
                raise Stop('Не заполнено поле ' + key + ' в УЗК, строка ' + str(index))
            record[key] = value.strip()
        if record['mark'] not in available:
            raise Stop('В подробном УЗК марка отсутствует в видимой ВОМ: ' + record['mark'])
        try:
            control_date = date.fromisoformat(record['date'])
        except ValueError as exc:
            raise Stop('Дата в подробном УЗК должна иметь формат YYYY-MM-DD: строка ' + str(index)) from exc
        if not start <= control_date <= end or control_date > confirmed:
            raise Stop('Дата подробного УЗК не соответствует периоду контроля.')
        # Multiple joint parameters may be separated by semicolons, but every
        # supplied value must be a finite, physically plausible measurement.
        freq = _frequency_parts(record['frequency_mhz'], 'частота')
        angle = _frequency_parts(record['angle_deg'], 'угол')
        if any(n > 20 for n in freq) or len(freq) != len(angle):
            raise Stop('Частоты/углы УЗК должны иметь одинаковое количество значений.')
        record['date'] = control_date.strftime('%d.%m.%Y')
        output.append(record)
    return output


def require_form_inputs(profile, facts):
    """Stop before preparation if a Word form has unconfirmed extra fields.

    No unrecognized placeholders are silently left behind. Standard registry
    and form markers are produced by the existing approved document workflow.
    """
    from .word_controls import inspect_word_fields, BINDABLE_FIELDS
    known = set(BINDABLE_FIELDS)|set(DOC_FIELD_LABELS)|{
        'VIK_TABLE','UZK_TABLE','UZK_DETAILS_TABLE','ACCEPTANCE_ACTS',
        'ATTACHMENTS',
    }
    provided = normalize_doc_fields(facts.get('doc_fields'))
    for role in ('vik','uzk','quality'):
        tokens, controls = inspect_word_fields(profile['templates'][role]['path'])
        actual = tokens | controls
        unknown = actual - known
        if unknown:
            raise Stop('Неизвестные поля шаблона '+role+': '+', '.join(sorted(unknown)))
        missing = (actual & set(DOC_FIELD_LABELS)) - set(provided)
        if missing:
            labels = '; '.join(DOC_FIELD_LABELS[key] for key in sorted(missing))
            raise Stop('Заполните обязательные реквизиты '+role+' из первичных данных: '+labels)
    return True


def uzk_details_to_table(details):
    return [
        (r['mark'], r['joint'], r['welders'], r['date'],
         r['frequency_mhz'], r['angle_deg'], r['defects'],
         r['conclusion'], r['basis'])
        for r in details
    ]


def _normalize_header(value):
    return re.sub(r'\s+', ' ', str(value or '')).strip().lower()


def load_uzk_csv(path):
    """Import operator-maintained UTF-8 semicolon CSV; duplicate marks allowed."""
    path = Path(path)
    if path.suffix.lower() != '.csv':
        raise Stop('Выберите файл УЗК .csv (UTF-8, разделитель точка с запятой).')
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f, delimiter=';')
        if not reader.fieldnames:
            raise Stop('В таблице УЗК отсутствует строка заголовков.')
        keys = [UZK_ALIASES.get(_normalize_header(k), _normalize_header(k)) for k in reader.fieldnames]
        if len(keys) != len(UZK_COLUMNS) or set(keys) != set(UZK_COLUMNS):
            raise Stop('В УЗК CSV требуются девять столбцов: '+UZK_CSV_HEADER)
        if len(set(keys)) != len(keys):
            raise Stop('Повторяющиеся столбцы УЗК CSV.')
        result = []
        for no, record in enumerate(reader, 2):
            if None in record:
                raise Stop('Лишний разделитель в УЗК CSV, строка ' + str(no))
            result.append({key: record.get(header) or ''
                           for key, header in zip(keys, reader.fieldnames)})
            if len(result) > 10000:
                raise Stop('Слишком много строк подробного УЗК.')
        return result


def acceptance_acts(facts, variants):
    """Act references from current passport's confirmed dates, never historic."""
    number = facts['number']
    vik = date.fromisoformat(facts['vik_date']).strftime('%d.%m.%Y')
    uzk = date.fromisoformat(facts['uzk_date']).strftime('%d.%m.%Y')
    akz = date.fromisoformat(facts['akz_date']).strftime('%d.%m.%Y')
    names = [
        f'Акт ВИК № {number} от {vik} г.',
        f'Акт ВИК № {number}-1 от {vik} г.',
        f'Акт УЗК № {number} от {uzk} г.',
    ]
    for variant in sorted(set(variants)):
        if variant not in range(4):
            raise Stop('Неверный вариант АКЗ для документа о качестве.')
        qualifier = '' if variant == 0 else f' (вариант {variant})'
        names.append(f'Сертификат качества АКЗ № {number}{qualifier} от {akz} г.')
    return ', '.join(names)


def values_for_documents(facts, variants):
    values = normalize_doc_fields(facts.get('doc_fields'))
    values['ACCEPTANCE_ACTS'] = acceptance_acts(facts, variants)
    values['ATTACHMENTS'] = values['ACCEPTANCE_ACTS']
    return values
