"""Prepare an EMPTY АКЗ appendix template using native Microsoft Excel.

Developer-only utility: NEVER writes source, registry or existing output.
The historical issued appendix is NOT itself an approved blank form.
This utility creates a separate candidate that requires human approval.
"""
import argparse
import hashlib
import os
import shutil
from pathlib import Path
from uuid import uuid4


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def validate_converted_appendix(sheet):
    """Explicit title/object/headers/dynamic-row contract of the new layout."""
    expected={'A2':'Объект: {{OBJECT}}',
              'A3':'Отправочная марка',
              'A5':'{{AKZ_MARKS}}','A6':'Итого'}
    for address,value in expected.items():
        if str(sheet.Range(address).Value2 or '').strip()!=value:
            raise ValueError('Неверная структура приложения АКЗ: '+address)
    if '{{NUMBER}}' not in str(sheet.Range('A1').Value2 or ''):
        raise ValueError('Нет номера паспорта в шапке АКЗ.')
    if '{{AKZ_DATE}}' not in str(sheet.Range('A1').Value2 or ''):
        raise ValueError('Нет даты акта в шапке АКЗ.')
    merged=str(sheet.Range('A2').MergeArea.Address).upper().replace('
    source = Path(source).resolve(strict=True)
    dest = Path(destination).absolute()
    if source.suffix.lower() != ".xlsx" or dest.suffix.lower() != ".xlsx":
        raise ValueError("Expected .xlsx source and destination.")
    if source == dest or dest.exists():
        raise FileExistsError("Output must be a NEW file, distinct from the source.")
    if not dest.parent.is_dir():
        raise FileNotFoundError("Create an explicit output directory first.")

    before = digest(source)
    scratch = dest.parent / (".vzmk-akz-template-" + uuid4().hex + ".xlsx")
    shutil.copyfile(source, scratch)
    app = None
    book = None
    try:
        import win32com.client
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        try:
            app.AutomationSecurity = 3  # disabled macros / external automation
        except Exception:
            pass
        book = app.Workbooks.Open(str(scratch), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        if book.Worksheets.Count != 1:
            raise ValueError("Expected exactly one appendix worksheet.")
        sheet = book.Worksheets(1)
        if sheet.ProtectContents:
            raise ValueError("Protected sheets cannot be converted automatically.")
        normalized = lambda v: str(v or "").strip().casefold()
        if (normalized(sheet.Range("A2").Value2) != "отправочная марка"
                or not normalized(sheet.Range("B2").Value2).startswith("наименование марки")
                or not normalized(sheet.Range("C2").Value2).startswith("количество")):
            raise ValueError("Unsupported АКЗ appendix header. NO changes committed.")

        last_used = int(sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1)
        totals = [row for row in range(4, last_used + 1)
                  if normalized(sheet.Cells(row, 1).Value2) == "итого"]
        if len(totals) != 1 or totals[0] <= 4:
            raise ValueError("Expected exactly one historical total row after mark rows.")

        # Insert one explicitly approved, merged OBJECT row. The historic
        # appendix has a title in A1:G1, headers in rows 2/3 and no place for
        # an object name. Adding the separate row preserves both header tiers.
        # All historical marks/results are cleared ONLY in the temporary copy.
        original_total = totals[0]
        sheet.Rows("2:2").Insert()
        sheet.Range("A2:G2").Merge()
        sheet.Range("A2").Value2 = "Объект: {{OBJECT}}"
        sheet.Range("A2").WrapText = True
        sheet.Rows("2:2").RowHeight = 75
        sheet.Range("A5:G" + str(original_total + 1)).ClearContents()
        sheet.Range("A1").Value2 = (
            "Приложение № {{NUMBER}} к акту АКЗ № {{NUMBER}} от {{AKZ_DATE}} г."
        )
        sheet.Range("A5").Value2 = "{{AKZ_MARKS}}"
        sheet.Range("A6").Value2 = "Итого"
        sheet.Range("A6:G6").Font.Bold = True
        sheet.Range("A5:G6").Borders.LineStyle = 1
        sheet.PageSetup.PrintArea = "$A$1:$G$6"
        sheet.PageSetup.PrintTitleRows = "$1:$4"
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
        sheet.PageSetup.Zoom = False
        validate_converted_appendix(sheet)
        book.Save()
        book.Close(SaveChanges=False)
        book = None
        proof = app.Workbooks.Open(str(scratch), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        try:
            check = proof.Worksheets(1)
            validate_converted_appendix(check)
        finally:
            proof.Close(SaveChanges=False)
        if digest(source) != before:
            raise RuntimeError("Source SHA changed unexpectedly — STOP.")
        if dest.exists():
            raise FileExistsError("Output appeared during preparation — STOP.")
        os.rename(scratch, dest)
        return {"destination": str(dest), "source_unchanged": True,
                "template_sha256": digest(dest), "approved": False}
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        # On failure retain a uniquely named scratch COPY for diagnostics.
        # Never touch the issued source.


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", help="historical issued appendix, READ ONLY")
    parser.add_argument("destination", help="NEW explicitly chosen copy")
    args = parser.parse_args()
    print(convert(args.source, args.destination))
,'')
    if merged!='A2:G2':
        raise ValueError('Объект АКЗ должен занимать объединённую строку A2:G2.')
    if not str(sheet.PageSetup.PrintArea or '').strip():
        raise ValueError('Утеряна область печати приложения АКЗ.')
    return True


def convert(source, destination):
    source = Path(source).resolve(strict=True)
    dest = Path(destination).absolute()
    if source.suffix.lower() != ".xlsx" or dest.suffix.lower() != ".xlsx":
        raise ValueError("Expected .xlsx source and destination.")
    if source == dest or dest.exists():
        raise FileExistsError("Output must be a NEW file, distinct from the source.")
    if not dest.parent.is_dir():
        raise FileNotFoundError("Create an explicit output directory first.")

    before = digest(source)
    scratch = dest.parent / (".vzmk-akz-template-" + uuid4().hex + ".xlsx")
    shutil.copyfile(source, scratch)
    app = None
    book = None
    try:
        import win32com.client
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        try:
            app.AutomationSecurity = 3  # disabled macros / external automation
        except Exception:
            pass
        book = app.Workbooks.Open(str(scratch), UpdateLinks=0, ReadOnly=False, AddToMru=False)
        if book.Worksheets.Count != 1:
            raise ValueError("Expected exactly one appendix worksheet.")
        sheet = book.Worksheets(1)
        if sheet.ProtectContents:
            raise ValueError("Protected sheets cannot be converted automatically.")
        normalized = lambda v: str(v or "").strip().casefold()
        if (normalized(sheet.Range("A2").Value2) != "отправочная марка"
                or not normalized(sheet.Range("B2").Value2).startswith("наименование марки")
                or not normalized(sheet.Range("C2").Value2).startswith("количество")):
            raise ValueError("Unsupported АКЗ appendix header. NO changes committed.")

        last_used = int(sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1)
        totals = [row for row in range(4, last_used + 1)
                  if normalized(sheet.Cells(row, 1).Value2) == "итого"]
        if len(totals) != 1 or totals[0] <= 4:
            raise ValueError("Expected exactly one historical total row after mark rows.")

        # Insert one explicitly approved, merged OBJECT row. The historic
        # appendix has a title in A1:G1, headers in rows 2/3 and no place for
        # an object name. Adding the separate row preserves both header tiers.
        # All historical marks/results are cleared ONLY in the temporary copy.
        original_total = totals[0]
        sheet.Rows("2:2").Insert()
        sheet.Range("A2:G2").Merge()
        sheet.Range("A2").Value2 = "Объект: {{OBJECT}}"
        sheet.Range("A2").WrapText = True
        sheet.Rows("2:2").RowHeight = 75
        sheet.Range("A5:G" + str(original_total + 1)).ClearContents()
        sheet.Range("A1").Value2 = (
            "Приложение № {{NUMBER}} к акту АКЗ № {{NUMBER}} от {{AKZ_DATE}} г."
        )
        sheet.Range("A5").Value2 = "{{AKZ_MARKS}}"
        sheet.Range("A6").Value2 = "Итого"
        sheet.Range("A6:G6").Font.Bold = True
        sheet.Range("A5:G6").Borders.LineStyle = 1
        sheet.PageSetup.PrintArea = "$A$1:$G$6"
        sheet.PageSetup.PrintTitleRows = "$1:$4"
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
        sheet.PageSetup.Zoom = False
        if (sheet.Range("A2").Value2 != "Объект: {{OBJECT}}" or
                sheet.Range("A3").Value2 != "Отправочная марка" or
                sheet.Range("A5").Value2 != "{{AKZ_MARKS}}" or
                sheet.Range("A6").Value2 != "Итого"):
            raise ValueError("Неверная структура нового приложения АКЗ.")
        book.Save()
        book.Close(SaveChanges=False)
        book = None
        proof = app.Workbooks.Open(str(scratch), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        try:
            check = proof.Worksheets(1)
            if (check.Range("A2").Value2 != "Объект: {{OBJECT}}" or
                    check.Range("A3").Value2 != "Отправочная марка" or
                    check.Range("A5").Value2 != "{{AKZ_MARKS}}" or
                    check.Range("A6").Value2 != "Итого" or
                    not str(check.PageSetup.PrintArea or "").strip()):
                raise ValueError("Excel Save/Open изменил шаблон приложения АКЗ.")
        finally:
            proof.Close(SaveChanges=False)
        if digest(source) != before:
            raise RuntimeError("Source SHA changed unexpectedly — STOP.")
        if dest.exists():
            raise FileExistsError("Output appeared during preparation — STOP.")
        os.rename(scratch, dest)
        return {"destination": str(dest), "source_unchanged": True,
                "template_sha256": digest(dest), "approved": False}
    finally:
        if book is not None:
            try:
                book.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        # On failure retain a uniquely named scratch COPY for diagnostics.
        # Never touch the issued source.


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", help="historical issued appendix, READ ONLY")
    parser.add_argument("destination", help="NEW explicitly chosen copy")
    args = parser.parse_args()
    print(convert(args.source, args.destination))
