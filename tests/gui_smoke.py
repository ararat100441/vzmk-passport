"""Exercise actual Tk controls and worker callbacks on real workbook copies."""
import json
import sys
import time
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from validate_real import copy_case, OUT
from passport_app.gui import Application
from passport_app.registry import read_registry


def settle(app):
    deadline = time.monotonic() + 180
    while app.busy and time.monotonic() < deadline:
        app.update()
        time.sleep(.04)
    if app.busy:
        raise AssertionError('GUI worker exceeded 180 seconds')
    app.update()


settings, vom, template = copy_case('gui_32')
app = Application(settings)
app.withdraw()
try:
    settle(app)
    assert len(app.combo['values']) == 53
    app.project_var.set('Гараж Казахстан')
    app.combo.event_generate('<<ComboboxSelected>>')
    settle(app)
    assert '1/2026' in app.passports_var.get() and '8/2026' in app.passports_var.get() and '9/2026' in app.passports_var.get()
    with patch('passport_app.gui.filedialog.askopenfilename', return_value=str(vom)):
        app.choose_button.invoke()
    settle(app)
    assert 'Строк: 32' in app.vom_var.get() and 'Балка' in app.vom_var.get()
    assert str(app.generate_button['state']) == 'disabled', repr(app.generate_button['state'])
    app.check_button.invoke()
    settle(app)
    assert app.plan is not None and 'Ошибок: 0' in app.output.get('1.0', 'end')
    assert str(app.generate_button['state']) == 'normal', repr(app.generate_button['state'])
    app.generate_button.invoke()
    settle(app)
    text = app.output.get('1.0', 'end')
    assert 'Блок 1 сформирован' in text, text
    assert str(app.generate_button['state']) == 'disabled', repr(app.generate_button['state'])
    records = read_registry(settings.registry_path).projects['Гараж Казахстан']
    assert records[-1]['number'] == '9/2026' and records[-1]['connections'] == 'Балка'
    result = {'pass': True, 'projects': 53, 'flow': 'project → choose VOM → check → generate',
              'row': records[-1]['row'], 'number': records[-1]['number'], 'root': str(OUT),
              'generation_summary': text, 'interactive_user_run': 'NOT_RUN'}
    (OUT / 'gui_results.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(result, ensure_ascii=False))
finally:
    app.destroy()
