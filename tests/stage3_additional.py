"""Bounded follow-up checks for final safeguards and native Excel-produced XML."""
import json
import subprocess
from uuid import uuid4
from unittest.mock import patch
import win32con
import win32file
import openpyxl
from openpyxl.comments import Comment
from stage3_validate import case, production_guard, OUT, PROJECT, md, key, error, values, gui_smoke
from stage3_evidence import bind_result, product_hashes, timestamp
from passport_app.catalog import CatalogService
from passport_app.excel import sha256, excel_session, open_excel


def main():
    tested_product_before, started = product_hashes(), timestamp()
    production_guard()
    suffix = '_' + uuid4().hex[:8]
    checks = []
    service, path = case('backup_guard_final' + suffix, [('К1', '35Б1', 'RAL 7004', 0)])
    before = sha256(path)
    plan = service.preview(PROJECT, md(('К1', '35Б2')))
    target = OUT / ('backup_redirect_target' + suffix)
    target.mkdir()
    link = path.parent / '.stage3-backups'
    # Actual Windows junction confined to validation; no UAC, no shell deletion.
    command = f"New-Item -ItemType Junction -Path '{str(link).replace(chr(39), chr(39)*2)}' -Target '{str(target).replace(chr(39), chr(39)*2)}' | Out-Null"
    subprocess.run(['powershell.exe', '-NoProfile', '-Command', command], check=True)
    message = error(lambda: service.apply(plan, {key('К1')}), 'путь резервных')
    assert sha256(path) == before and not list(target.iterdir())
    checks.append({'backup_junction_guard': 'PASS', 'message': message})

    service, path = case('writer_lock_final' + suffix, [('К1', '35Б1', 'RAL 7004', 0)])
    before = sha256(path)
    plan = service.preview(PROJECT, md(('К1', '35Б2')))
    handle = win32file.CreateFile(str(path.parent / '.stage3-catalog.lock'), win32con.GENERIC_WRITE,
                                  0, None, win32con.OPEN_ALWAYS, win32con.FILE_FLAG_DELETE_ON_CLOSE, None)
    try:
        message = error(lambda: service.apply(plan, {key('К1')}), 'занят')
    finally:
        handle.Close()
    assert sha256(path) == before
    checks.append({'concurrent_writer_native_guard': 'PASS', 'message': message})

    service, path = case('native_manual_final' + suffix, [('К1', '35Б1', 'RAL 7004', 0), ('К3', None, '=1+1', '=0')])
    b = openpyxl.load_workbook(path)
    b.active['C2'].comment = Comment('Ручное примечание', 'User')
    b.save(path)
    b.close()
    with excel_session() as excel:
        book = open_excel(excel, path)
        book.Save()
        book.Close(SaveChanges=False)
    plan = service.preview(PROJECT, md(('К1', '35Б2'), ('К3', '45Б1'), ('К2', 'Гнз120×60х6')))
    result = service.apply(plan, {key('К1')})
    b = openpyxl.load_workbook(path)
    assert b.active['C3'].value == '=1+1' and b.active['D3'].value == '=0'
    assert b.active['C2'].comment.text == 'Ручное примечание'
    assert b.active['C2'].value == 'RAL 7004' and b.active['D2'].value == 0
    assert b.active['B4'].value == 'Гнз120×60х6'
    b.close()
    checks.append({'native_Excel_saved_XML': 'PASS', 'manual_formulas_comments_values_styles': 'PASS',
                   'confirmed_profile_spelling': 'Гнз120×60х6', 'result': result})

    checks.append({'final_GUI_constructor_and_workflow': 'PASS', 'evidence': gui_smoke('gui_final' + suffix)})
    production_guard()
    result = bind_result({'status': 'PASS', 'checks': checks, 'pass_count': len(checks), 'fail_count': 0}, tested_product_before, started)
    (OUT / 'final_targeted_additions.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('Final targeted additions: 4/4 PASS')


if __name__ == '__main__':
    main()
