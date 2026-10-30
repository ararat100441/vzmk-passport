import hashlib
import ctypes
from contextlib import contextmanager
from pathlib import Path

import openpyxl
from .config import PassportError


def sha256(path):
    try:
        return hashlib.sha256(Path(path).read_bytes()).hexdigest()
    except OSError as exc:
        raise PassportError(f'Не удалось прочитать файл: {path}. Он отсутствует или занят другим процессом.') from exc


def atomic_replace(source, target, *, backup=None):
    """Native Windows replacement preserves target ACLs/metadata and supports delete sharing."""
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    function = kernel.ReplaceFileW
    function.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_wchar_p,
                         ctypes.c_uint32, ctypes.c_void_p, ctypes.c_void_p]
    function.restype = ctypes.c_int
    def extended(path):
        return '\\\\?\\' + str(Path(path).resolve())
    if not function(extended(target), extended(source), extended(backup) if backup else None, 0, None, None):
        raise ctypes.WinError(ctypes.get_last_error())


def load_book(path, *, values=False):
    path = Path(path)
    if not path.is_file():
        raise PassportError(f'Не найден файл: {path}')
    if path.suffix.lower() != '.xlsx':
        raise PassportError(f'Требуется книга .xlsx: {path.name}. Исходный файл не преобразуется автоматически.')
    try:
        return openpyxl.load_workbook(path, data_only=values, keep_links=True)
    except Exception as exc:
        raise PassportError(f'Не удалось прочитать {path.name}: {exc}') from exc


@contextmanager
def locked_file(path, *, writable=False):
    """Deny readers/writers; allow our atomic replacement while the old handle lives."""
    import win32con
    import win32file
    try:
        handle = win32file.CreateFile(str(path), win32con.GENERIC_READ | (win32con.GENERIC_WRITE | win32con.DELETE if writable else 0),
                                      win32con.FILE_SHARE_DELETE, None, win32con.OPEN_EXISTING, 0, None)
    except Exception as exc:
        raise PassportError(f'Файл занят другим процессом или недоступен для записи: {path}. Закройте его в Excel.') from exc
    try:
        chunks = []
        while True:
            _, chunk = win32file.ReadFile(handle, 1024 * 1024)
            if not chunk:
                break
            chunks.append(chunk)
        yield b''.join(chunks)
    finally:
        handle.Close()


@contextmanager
def excel_session():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3
        excel.ScreenUpdating = False
        yield excel
    except PassportError:
        raise
    except Exception as exc:
        raise PassportError(f'Ошибка Microsoft Excel: {exc}') from exc
    finally:
        if excel is not None:
            try:
                excel.CutCopyMode = False
                excel.Workbooks.Close()
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def open_excel(excel, path):
    book = excel.Workbooks.Open(Filename=str(Path(path).resolve()), UpdateLinks=0, ReadOnly=False,
                                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False)
    if book.ReadOnly:
        book.Close(SaveChanges=False)
        raise PassportError(f'Excel открыл временную книгу только для чтения: {path}')
    return book


def assert_no_formula_errors(sheet, row, columns):
    # SpecialCells(xlCellTypeFormulas, xlErrors) checks actual native calculation.
    region = sheet.Range(sheet.Cells(row, 1), sheet.Cells(row, columns))
    try:
        errors = region.SpecialCells(-4123, 16)
    except Exception:
        return
    if errors is not None:
        raise PassportError(f'В новой строке Excel есть ошибка формулы: {errors.Address}.')
