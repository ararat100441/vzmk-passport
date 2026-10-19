"""Safe, opt-in Word content-control bindings for approved blank forms.

A historical issued document is not a template: populated controls are refused.
Table insertion markers remain explicit {{VIK_TABLE}}/{{UZK_TABLE}} tokens.
"""
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .storage import Stop

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
TOKEN = re.compile(r'\{\{([A-Z_0-9]+)\}\}')
TAG_ALIASES = {'PASSPORT_NO': 'NUMBER'}
BINDABLE_FIELDS = frozenset({
    'NUMBER', 'CONTRACT', 'CUSTOMER', 'OBJECT', 'CIPHER', 'MASS',
    'START_DATE', 'END_DATE', 'KM', 'KMD_COMPANY', 'STEEL',
    'VIK_DATE', 'UZK_DATE', 'AKZ_DATE', 'CONNECTIONS',
    'RESPONSIBLE', 'EVIDENCE',
})


def normalized_tag(tag):
    tag = str(tag or '').strip()
    return TAG_ALIASES.get(tag, tag)


def inspect_word_fields(path):
    """Return (literal placeholders, safe empty/placeholder SDT fields).

    Read-only OOXML inspection intentionally handles only document body SDTs.
    Headers/footers cannot count as bound fields until explicitly supported.
    """
    with zipfile.ZipFile(Path(path)) as package:
        root = ET.fromstring(package.read('word/document.xml'))
    text = ''.join(n.text or '' for n in root.iter(W + 't'))
    tokens = set(TOKEN.findall(text))
    controls = set()
    for sdt in root.iter(W + 'sdt'):
        props = sdt.find(W + 'sdtPr')
        if props is None:
            continue
        tag_element = props.find(W + 'tag')
        if tag_element is None:
            continue
        raw_tag = tag_element.get(W + 'val', '')
        key = normalized_tag(raw_tag)
        if key not in BINDABLE_FIELDS:
            continue
        body = sdt.find(W + 'sdtContent')
        if body is None:
            raise Stop('Поле Word без содержимого: ' + str(raw_tag))
        current = ''.join(n.text or '' for n in body.iter(W + 't')).strip()
        placeholders = {'{{' + raw_tag + '}}', '{{' + key + '}}'}
        showing_placeholder = props.find(W + 'showingPlcHdr') is not None
        if current and current not in placeholders and not showing_placeholder:
            raise Stop('Поле Word уже заполнено историческими данными: ' + str(raw_tag)
                       + '. Используйте отдельную очищенную и утверждённую форму.')
        controls.add(key)
    return tokens, controls


def fill_word_controls(document, values):
    """Fill only recognized blank fields; never overwrite issued facts."""
    bound = set()
    collection = document.ContentControls
    for index in range(1, collection.Count + 1):
        control = collection.Item(index)
        raw_tag = str(control.Tag or '').strip()
        key = normalized_tag(raw_tag)
        if key not in BINDABLE_FIELDS or key not in values:
            continue
        current = str(control.Range.Text or '').strip()
        placeholders = {'{{' + raw_tag + '}}', '{{' + key + '}}'}
        showing_placeholder = bool(control.ShowingPlaceholderText)
        if current and current not in placeholders and not showing_placeholder:
            raise Stop('Поле Word уже содержит данные предыдущего паспорта: ' + raw_tag)
        if control.LockContents:
            raise Stop('Поле Word заблокировано от записи: ' + raw_tag)
        control.Range.Text = str(values[key])
        bound.add(key)
    return bound
