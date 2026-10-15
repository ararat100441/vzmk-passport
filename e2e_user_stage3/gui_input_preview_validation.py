"""Targeted real Tk GUI validation. No Apply, no production workbook access."""
import argparse
import ctypes
from contextlib import redirect_stderr
from datetime import datetime, timezone
import hashlib
from io import StringIO
import json
import os
from pathlib import Path
import sys
import tkinter as tk
import traceback
import time
from unittest.mock import patch

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')
E2E = Path(__file__).resolve().parent
ROOT = E2E.parent
OUT = E2E / '.state' / 'gui_input_preview_fix'
PROJECT = 'Гараж Казахстан'


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def product_hashes():
    return {str(p): sha(p) for p in sorted((ROOT / 'passport_app').glob('*.py')) + [ROOT / 'stage3.py']}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('phase', choices=('before', 'after'))
    args = parser.parse_args()
    assert sys.flags.isolated and sys.dont_write_bytecode
    sys.path.insert(0, str(E2E / 'runtime'))
    from passport_app.config import Settings
    from passport_app.catalog_gui import CatalogApplication
    import passport_app.catalog as catalog_module

    baseline = json.loads((E2E / 'baseline.json').read_text(encoding='utf-8'))
    dev = ROOT / 'DEV' / 'ВЗМК_DEV.xlsx'
    catalog = E2E / 'Документы' / 'Гараж казахстан' / 'Справочник марок.xlsx'
    assert sha(dev) == baseline['development_registry_sha256'], 'STOP: DEV changed'
    assert sha(catalog) == baseline['catalog_initial_sha256'], 'STOP: E2E catalog changed'
    assert sha(E2E / 'ВЗМК_DEV.xlsx') == sha(dev)
    OUT.mkdir(exist_ok=True)
    hashes_before = product_hashes()
    for source, digest in hashes_before.items():
        assert sha(E2E / 'runtime' / Path(source).relative_to(ROOT)) == digest, 'runtime/source differ'
    (OUT / (args.phase + '_hashes_before.json')).write_text(json.dumps(hashes_before, indent=2), encoding='utf-8')
    if args.phase == 'before':
        if (OUT / 'catalog_gui.before.py').exists():
            assert sha(OUT / 'catalog_gui.before.py') == sha(ROOT / 'passport_app' / 'catalog_gui.py')
        else:
            (OUT / 'catalog_gui.before.py').write_bytes((ROOT / 'passport_app' / 'catalog_gui.py').read_bytes())
    inventory_before = {str(p): sha(p) for p in catalog.parent.rglob('*') if p.is_file()}
    app = CatalogApplication(Settings.load(E2E / 'settings.json'))
    app.update()
    app.input.focus_force()
    app.update()
    user = ctypes.WinDLL('user32', use_last_error=True)
    user.GetForegroundWindow.restype = ctypes.c_void_p
    user.GetKeyboardLayout.restype = ctypes.c_void_p
    events, callback_errors, dialogs, checks = [], [], [], {}
    app.input.bind('<KeyPress>', lambda e: events.append({'keycode': e.keycode, 'keysym': e.keysym, 'state': e.state}), add='+')

    def callback_error(kind, value, tb):
        callback_errors.append(''.join(traceback.format_exception(kind, value, tb)))
    app.report_callback_exception = callback_error
    def visible_error(title, message, **kwargs):
        dialogs.append({'title': title, 'message': message})

    def content():
        return app.input.get('1.0', 'end-1c')
    def replace(text, update=True):
        app.input.delete('1.0', 'end')
        app.input.insert('1.0', text)
        if update:
            app.update()
    def select(name=PROJECT):
        app.project_var.set(name)
        app.combo.event_generate('<<ComboboxSelected>>')
        app.update()
    def clipboard(text):
        app.clipboard_clear()
        app.clipboard_append(text)
        app.update()
    def key(vk, modifier=17):
        app.lift()
        app.attributes('-topmost', True)
        app.input.focus_force()
        app.update()
        user.GetAncestor.restype = ctypes.c_void_p
        hwnd = user.GetAncestor(ctypes.c_void_p(app.winfo_id()), 2)
        user.SetForegroundWindow(ctypes.c_void_p(hwnd))
        app.update()
        app.attributes('-topmost', False)
        pid = ctypes.c_ulong()
        foreground_thread = user.GetWindowThreadProcessId(ctypes.c_void_p(user.GetForegroundWindow()), ctypes.byref(pid))
        if pid.value != os.getpid():
            own_thread = ctypes.WinDLL('kernel32').GetCurrentThreadId()
            attached = user.AttachThreadInput(own_thread, foreground_thread, True)
            try:
                user.SetForegroundWindow(ctypes.c_void_p(hwnd))
                user.SetFocus(ctypes.c_void_p(app.input.winfo_id()))
            finally:
                if attached:
                    user.AttachThreadInput(own_thread, foreground_thread, False)
            app.update()
            user.GetWindowThreadProcessId(ctypes.c_void_p(user.GetForegroundWindow()), ctypes.byref(pid))
        assert pid.value == os.getpid(), f'STOP: keyboard focus belongs to another process (foreground={pid.value}, own={os.getpid()}, hwnd={hwnd}, focus={app.focus_get()})'
        try:
            modifier_vk = 160 if modifier == 16 else modifier
            user.keybd_event(modifier_vk, user.MapVirtualKeyW(modifier_vk, 0), 0, 0)
            time.sleep(0.03)
            app.update()
            extended = 1 if vk == 45 else 0  # Dedicated Insert uses an extended scan code.
            user.keybd_event(vk, user.MapVirtualKeyW(vk, 0), extended, 0)
            time.sleep(0.03)
            app.update()
        finally:
            user.keybd_event(vk, user.MapVirtualKeyW(vk, 0), 2 | extended, 0)
            user.keybd_event(modifier_vk, user.MapVirtualKeyW(modifier_vk, 0), 2, 0)
            time.sleep(0.03)
            app.update()
    def check():
        app.check_button.invoke()
        app.update()
    def blocked():
        return app.plan is None and str(app.apply_button['state']) == 'disabled' and not app.tree.get_children()
    def error_case(name, text, project=PROJECT, effect=None):
        select(project)
        replace(text)
        previous = len(dialogs)
        with patch('passport_app.catalog_gui.messagebox.showerror', side_effect=visible_error):
            if effect is None:
                check()
            else:
                with patch.object(app.service, 'preview', side_effect=effect):
                    check()
        checks[name] = blocked() and len(dialogs) == previous + 1 and bool(app.result_var.get())

    old_clipboard = None
    try:
        try:
            old_clipboard = app.clipboard_get()
        except tk.TclError:
            pass
        small = '| Марка | Профиль |\n|---|---|\n| К1 | 35Б1 |'
        full = (E2E / 'E2E_ИМПОРТ.txt').read_text(encoding='utf-8-sig')
        file_text = (E2E / 'E2E_ИМПОРТ.txt').read_bytes().decode('utf-8-sig')
        original_info = {'tk_version': app.tk.call('info', 'patchlevel'),
                         'window_system': app.tk.call('tk', 'windowingsystem'),
                         'keyboard_layout': hex(user.GetKeyboardLayout(0)),
                         'widget_state': str(app.input['state']), 'focus': str(app.focus_get()),
                         'bindtags': app.input.bindtags(), 'widget_bindings': app.input.bind(),
                         'paste_events': app.event_info('<<Paste>>'),
                         'open_file_button_exists': hasattr(app, 'open_file_button')}
        clipboard(small)
        replace('')
        key(86)
        checks['native_ctrl_v'] = content() == small
        replace('')
        key(45, modifier=16)
        checks['native_shift_insert'] = content() == small
        select()
        replace(full)
        chain = []
        real_parse, real_preview = catalog_module.parse_input, app.service.preview
        def parse(text):
            chain.append({'step': 'parser', 'input_matches': text == full})
            return real_parse(text)
        def preview(project, text):
            chain.append({'step': 'service.preview', 'project': project, 'input_matches': text == full})
            result = real_preview(project, text)
            chain.append({'step': 'preview_model', 'rows': len(result.rows)})
            return result
        with patch.object(catalog_module, 'parse_input', side_effect=parse), patch.object(app.service, 'preview', side_effect=preview):
            check()
        statuses = [row.status for row in app.plan.rows] if app.plan else []
        checks['correct_markdown_preview'] = app.plan is not None and len(app.tree.get_children()) == 5
        checks['five_statuses'] = statuses == ['SAME', 'FILL', 'NEW', 'CONFLICT', 'CHECK_REQUIRED']
        chain.append({'step': 'render', 'tree_rows': len(app.tree.get_children()), 'result': app.result_var.get()})
        replace(full, update=False)
        app.check_button.invoke()
        immediate = app.plan is not None
        app.update()
        checks['pending_modified_does_not_erase_preview'] = immediate and app.plan is not None
        if args.phase == 'before':
            previous = app.result_var.get()
            with patch.object(app.service, 'preview', side_effect=RuntimeError('TARGETED_PREVIEW_EXCEPTION')):
                check()
            checks['unexpected_exception_visible'] = app.result_var.get() != previous
        else:
            replace('')
            clipboard(full)
            key(86)
            checks['paste_reaches_preview_handler'] = content() == full
            check()
            checks['paste_reaches_preview_handler'] &= app.plan is not None and app.plan.text == full
            key(65)
            checks['native_ctrl_a'] = app.input.get('sel.first', 'sel.last') == full
            key(67)
            checks['native_ctrl_c'] = app.clipboard_get() == full
            key(88)
            checks['native_ctrl_x'] = content() == '' and blocked()
            clipboard(full)
            key(86)
            check()
            with patch('passport_app.catalog_gui.filedialog.askopenfilename', return_value=str(E2E / 'E2E_ИМПОРТ.txt')) as picker:
                app.open_file_button.invoke()
                app.update()
            checks['open_file_autofill_utf8_bom'] = content() == file_text and picker.call_count == 1
            checks['file_load_invalidates_plan'] = blocked()
            checks['loaded_file_path_visible'] = str(E2E / 'E2E_ИМПОРТ.txt') in app.result_var.get()
            check()
            checks['file_loaded_text_previews'] = app.plan is not None and app.plan.text == file_text
            no_bom = OUT / 'input_utf8_no_bom.txt'
            no_bom.write_bytes(full.encode('utf-8'))
            with patch('passport_app.catalog_gui.filedialog.askopenfilename', return_value=str(no_bom)):
                app.open_file_button.invoke()
                app.update()
            checks['open_file_utf8_without_bom'] = content() == full and blocked()
            bad_utf8 = OUT / 'input_invalid_utf8.txt'
            bad_utf8.write_bytes(b'\xff\xfe\xff')
            previous = len(dialogs)
            with patch('passport_app.catalog_gui.messagebox.showerror', side_effect=visible_error), patch('passport_app.catalog_gui.filedialog.askopenfilename', return_value=str(bad_utf8)):
                app.open_file_button.invoke()
                app.update()
            checks['invalid_utf8_visible_error'] = blocked() and len(dialogs) == previous + 1 and content() == full
            app.input.insert('end', '\n')
            app.update()
            checks['edit_invalidates_plan'] = blocked()
            check()
            select(next(name for name in app.combo['values'] if name != PROJECT))
            checks['project_invalidates_plan'] = blocked()
            error_case('empty_visible_error', '')
            error_case('malformed_visible_error', 'не таблица')
            error_case('unselected_project_visible_error', full, project='')
            error_case('unexpected_preview_visible_error', full, effect=RuntimeError('TARGETED_PREVIEW_EXCEPTION'))
            with patch('passport_app.catalog_gui.messagebox.showerror', side_effect=visible_error), patch.object(catalog_module, 'parse_input', side_effect=ValueError('TARGETED_PARSER_EXCEPTION')):
                check()
            checks['unexpected_parser_visible_error'] = blocked() and 'внутренней ошибки' in app.result_var.get()
            select()
            replace(full)
            check()
            previous = len(dialogs)
            with patch('passport_app.catalog_gui.messagebox.showerror', side_effect=visible_error), patch('passport_app.catalog_gui.filedialog.askopenfilename', return_value=str(OUT / 'missing.txt')):
                app.open_file_button.invoke()
                app.update()
            checks['unreadable_file_visible_error'] = blocked() and len(dialogs) == previous + 1 and content() == full
            check()
            plan = app.plan
            with patch('passport_app.catalog_gui.filedialog.askopenfilename', return_value=''):
                app.open_file_button.invoke()
                app.update()
            checks['file_cancel_keeps_plan'] = app.plan is plan
            replace(full)
            check()
            checks['rendered_status_labels'] = [app.tree.set(str(i), 'status') for i in range(5)] == [catalog_module.LABELS[s] for s in statuses]
            checks['no_callback_exceptions'] = not callback_errors
    finally:
        if old_clipboard is not None:
            app.clipboard_clear()
            app.clipboard_append(old_clipboard)
            app.update()
        app.destroy()
    hashes_after = product_hashes()
    (OUT / (args.phase + '_hashes_after.json')).write_text(json.dumps(hashes_after, indent=2), encoding='utf-8')
    checks['product_hashes_before_equals_after'] = hashes_before == hashes_after
    checks['preview_no_write'] = inventory_before == {str(p): sha(p) for p in catalog.parent.rglob('*') if p.is_file()}
    checks['e2e_catalog_unchanged'] = sha(catalog) == baseline['catalog_initial_sha256']
    checks['dev_unchanged'] = sha(dev) == baseline['development_registry_sha256']
    result = {'phase': args.phase, 'timestamp_utc': datetime.now(timezone.utc).isoformat(),
              'checks': checks, 'status': 'PASS' if all(checks.values()) else 'REPRODUCED' if args.phase == 'before' else 'FAIL',
              'environment': original_info, 'key_events': events, 'chain': chain, 'preview_statuses': statuses,
              'visible_error_dialog_calls': dialogs, 'callback_tracebacks': callback_errors,
              'product_hashes_before': hashes_before, 'product_hashes_after': hashes_after,
              'catalog_sha_before_after': sha(catalog), 'dev_sha_before_after': sha(dev),
              'file_picker_selection': 'AUTOMATED_RETURN_VALUE; REAL_BUTTON_AND_FILE_READER',
              'error_dialogs': 'CALLS_CAPTURED; REAL_VISIBLE_INLINE_STATUS', 'production_access': 'NOT_PERFORMED',
              'Apply': 'NOT_RUN', 'USER_E2E': 'NOT_RUN'}
    (OUT / (args.phase + '_result.json')).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    (OUT / (args.phase + '_callback_traceback.txt')).write_text('\n'.join(callback_errors), encoding='utf-8')
    print(json.dumps({'phase': args.phase, 'status': result['status'], 'checks': checks}, ensure_ascii=False, indent=2))
    if args.phase == 'after' and not all(checks.values()):
        raise SystemExit(1)


if __name__ == '__main__':
    console = StringIO()
    try:
        with redirect_stderr(console):
            main()
    finally:
        OUT.mkdir(parents=True, exist_ok=True)
        (OUT / (sys.argv[1] + '_console_tracebacks.txt')).write_text(console.getvalue(), encoding='utf-8')
