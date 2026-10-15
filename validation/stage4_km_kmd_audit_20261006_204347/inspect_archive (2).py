"""Read-only source inventory. Never starts Excel or evaluates workbook code."""
import hashlib
import json
import stat
import sys
import zipfile
from pathlib import Path, PurePosixPath

import olefile
import xlrd

sys.stdout.reconfigure(encoding="utf-8")
BASE = Path(__file__).resolve().parent
TARGET = BASE / "extracted"
TARGET.mkdir(exist_ok=True)


def save(name, data):
    (BASE / name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


archive = zipfile.ZipFile(BASE / "sortament.zip")
bad = archive.testzip()
if bad is not None:
    raise ValueError(f"Corrupt archive member: {bad}")
members = []
names = set()
for info in archive.infolist():
    # This ZIP has no UTF-8 flag; its raw filenames are DOS Cyrillic CP866.
    raw_name = info.filename.encode("cp437") if not info.flag_bits & 0x800 else None
    name = raw_name.decode("cp866") if raw_name else info.filename
    normalized = name.replace("\\", "/")
    parts = PurePosixPath(normalized)
    if parts.is_absolute() or ".." in parts.parts or ":" in normalized:
        raise ValueError(f"Unsafe path: {name}")
    if stat.S_ISLNK(info.external_attr >> 16):
        raise ValueError(f"Symlink forbidden: {name}")
    target = (TARGET / normalized).resolve()
    if TARGET.resolve() not in target.parents or str(target).casefold() in names:
        raise ValueError(f"Unsafe or duplicate target: {name}")
    names.add(str(target).casefold())
    if info.flag_bits & 1:
        raise ValueError(f"Encrypted entry: {name}")
    if target.suffix.lower() not in {".xls", ".xlsx", ".csv"}:
        raise ValueError(f"Non-table member: {name}")
    data = archive.read(info)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and target.read_bytes() != data:
        raise ValueError(f"Existing extracted file differs: {name}")
    if not target.exists():
        target.write_bytes(data)
    members.append({"name": name, "zip_python_name": info.filename,
                    "raw_filename_hex": raw_name.hex() if raw_name else None,
                    "filename_encoding": "cp866" if raw_name else "utf-8",
                    "size": len(data), "compressed_size": info.compress_size,
                    "sha256": hashlib.sha256(data).hexdigest()})
save("archive_inventory.json", {"integrity": "PASS", "members": members})

inventory = []
for member in members:
    path = TARGET / member["name"]
    book = xlrd.open_workbook(str(path), on_demand=True, formatting_info=True)
    with olefile.OleFileIO(str(path)) as ole:
        streams = ["/".join(s) for s in ole.listdir()]
    item = {"filename": path.name, "format": "OLE/BIFF XLS", "biff_version": book.biff_version,
            "sha256": member["sha256"], "ole_streams": streams,
            "macro_streams": [s for s in streams if "vba" in s.lower() or "macro" in s.lower()],
            "sheets": []}
    for sheet in book.sheets():
        rows = []
        used = []
        for r in range(sheet.nrows):
            cells = []
            for c in range(sheet.ncols):
                cell = sheet.cell(r, c)
                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    continue
                xf = book.xf_list[cell.xf_index]
                fmt = book.format_map[xf.format_key].format_str
                cells.append({"col": c + 1, "value": cell.value, "ctype": cell.ctype, "format": fmt})
                used.append((r + 1, c + 1))
            if cells:
                rows.append({"row": r + 1, "cells": cells})
        item["sheets"].append({"name": sheet.name, "allocated_rows": sheet.nrows,
                               "allocated_columns": sheet.ncols,
                               "nonempty_row_count": len(rows),
                               "used_bounds": [min(r for r, c in used), min(c for r, c in used),
                                               max(r for r, c in used), max(c for r, c in used)] if used else None,
                               "merged_ranges_zero_based": sheet.merged_cells,
                               "rows": rows})
    inventory.append(item)
    book.release_resources()
save("raw_workbook_inventory.json", inventory)
for item in inventory:
    print("\nWORKBOOK", item["filename"], "BIFF", item["biff_version"], "macros", item["macro_streams"])
    for s in item["sheets"]:
        print("SHEET", s["name"], "allocated", s["allocated_rows"], s["allocated_columns"],
              "bounds", s["used_bounds"], "nonempty_rows", s["nonempty_row_count"],
              "merged", s["merged_ranges_zero_based"])
        for row in s["rows"][:22] + s["rows"][-6:]:
            print(row["row"], [(c["col"], c["value"]) for c in row["cells"]])
