"""Observe actual Win32 file picker and error dialog in our own Tk process only."""
import ctypes
import hashlib
import json
import os
from pathlib import Path
import sys
import threading
import time

sys.stdout.reconfigure(encoding='utf-8')
E2E = Path(__file__).resolve().parent
ROOT = E2E.parent
OUT = E2E / '.state' / 'gui_input_preview_fix'
sys.path.insert(0, str(E2E / 'runtime'))
from passport_app.config import Settings
from passport_app.catalog_gui import CatalogApplication

user = ctypes.WinDLL('user32', use_last_error=True)
user.GetWindowTextW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int)
user.GetDlgItem.argtypes = (ctypes.c_void_p, ctypes.c_int)
user.GetDlgItem.restype = ctypes.c_void_p
user.SetWindowTextW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p)
user.PostMessageW.argtypes = (ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t)
CALLBACK = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_ssize_t)
observed = []


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def controller(title, filename=None):
    deadline = time.monotonic() + 15
    handled = set()
    while time.monotonic() < deadline:
        windows = []
        @CALLBACK
        def scan(hwnd, data):
            pid = ctypes.c_ulong()
            user.GetWindowThreadProcessId(ctypes.c_void_p(hwnd), ctypes.byref(pid))
            name = ctypes.create_unicode_buffer(512)
            user.GetWindowTextW(hwnd, name, len(name))
            if pid.value == os.getpid() and name.value == title:
                windows.append(hwnd)
            return True
        user.EnumWindows(scan, 0)
        for hwnd in windows:
            if hwnd in handled:
                continue
            if not user.IsWindowVisible(ctypes.c_void_p(hwnd)):
                continue  # Wait for the real dialog to appear before selecting/closing.
            handled.add(hwnd)
            entry = {'title': title, 'own_process': True, 'window_visible': bool(user.IsWindowVisible(ctypes.c_void_p(hwnd)))}
            if filename is not None:
                edit = user.GetDlgItem(hwnd, 1152) or user.GetDlgItem(hwnd, 1148)
                if not edit:
                    edits = []
                    @CALLBACK
                    def child(chwnd, data):
                        classname = ctypes.create_unicode_buffer(128)
                        user.GetClassNameW(ctypes.c_void_p(chwnd), classname, len(classname))
                        if classname.value == 'Edit':
                            edits.append(chwnd)
                        return True
                    user.EnumChildWindows(ctypes.c_void_p(hwnd), child, 0)
                    edit = edits[0] if edits else None
                if not edit:
                    entry['error'] = 'file-name control unavailable'
                    observed.append(entry)
                    user.PostMessageW(hwnd, 0x0111, 2, 0)
                    return
                user.SetWindowTextW(edit, str(filename))
                entry['selected_file'] = str(filename)
            observed.append(entry)
            user.PostMessageW(hwnd, 0x0111 if filename is not None else 0x0010, 1 if filename is not None else 0, 0)
            return
        time.sleep(0.02)


app = CatalogApplication(Settings.load(E2E / 'settings.json'))
app.update()
catalog = E2E / 'Документы' / 'Гараж казахстан' / 'Справочник марок.xlsx'
dev = ROOT / 'DEV' / 'ВЗМК_DEV.xlsx'
before = {str(p): sha(p) for p in (catalog, dev, ROOT / 'passport_app' / 'catalog_gui.py')}
try:
    app.project_var.set('Гараж Казахстан')
    app.combo.event_generate('<<ComboboxSelected>>')
    app.update()
    worker = threading.Thread(target=controller, args=('Открыть файл импорта', E2E / 'E2E_ИМПОРТ.txt'), daemon=True)
    worker.start()
    app.open_file_button.invoke()
    app.update()
    worker.join(timeout=1)
    loaded = app.input.get('1.0', 'end-1c') == (E2E / 'E2E_ИМПОРТ.txt').read_bytes().decode('utf-8-sig')
    app.check_button.invoke()
    app.update()
    preview = app.plan is not None and len(app.tree.get_children()) == 5
    app.input.delete('1.0', 'end')
    app.update()
    worker = threading.Thread(target=controller, args=('Ошибка импорта',), daemon=True)
    worker.start()
    app.check_button.invoke()
    app.update()
    worker.join(timeout=1)
    error = app.plan is None and str(app.apply_button['state']) == 'disabled' and 'Вставьте таблицу' in app.result_var.get()
finally:
    app.destroy()
after = {p: sha(p) for p in before}
result = {'status': 'PASS' if loaded and preview and error and before == after and len(observed) == 2 and all(x['window_visible'] for x in observed) else 'FAIL',
          'native_file_picker_auto_fill': loaded, 'native_file_loaded_preview': preview,
          'actual_error_dialog_visible': error, 'observed': observed,
          'hashes_before': before, 'hashes_after': after, 'Apply': 'NOT_RUN', 'production_access': 'NOT_PERFORMED'}
(OUT / 'native_dialog_result.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(result, ensure_ascii=False, indent=2))
if result['status'] != 'PASS':
    raise SystemExit(1)
