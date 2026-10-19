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

        # Remove old marks, quantities, weights, areas and historical total values
        # from the COPY only. Keep all header cells, merges, borders and native
        # printer settings. Runtime inserts dynamic rows and recalculates totals.
        sheet.Range("A4:G" + str(totals[0])).ClearContents()
        sheet.Range("A1").Value2 = (
            "Приложение № {{NUMBER}} к акту АКЗ № {{NUMBER}} от {{AKZ_DATE}} г."
        )
        sheet.Range("A4").Value2 = "{{AKZ_MARKS}}"
        sheet.Range("A5").Value2 = "Итого"
        sheet.Range("A5:G5").Font.Bold = True
        sheet.Range("A4:G5").Borders.LineStyle = 1
        sheet.PageSetup.PrintArea = "$A$1:$G$5"
        sheet.PageSetup.PrintTitleRows = "$1:$3"
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
        sheet.PageSetup.Zoom = False
        book.Save()
        book.Close(SaveChanges=False)
        book = None
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
