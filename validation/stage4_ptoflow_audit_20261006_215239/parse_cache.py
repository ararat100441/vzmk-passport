"""OFFLINE only. Parse saved HTML and safely decode Next Flight JSON strings."""
import collections
import hashlib
import json
import re
import sys
from pathlib import Path

from lxml import html

BASE = Path(__file__).resolve().parent
sys.stdout.reconfigure(encoding="utf-8")


def text(node):
    return " ".join(" ".join(node.itertext()).split())


def number(raw):
    cleaned = re.sub(r"[\s\u00a0]", "", str(raw)).replace(",", ".")
    return float(cleaned) if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", cleaned) else None


def flight_data(tree):
    decoder = json.JSONDecoder()
    payload = ""
    for script in tree.xpath("//script[not(@src)]"):
        value = script.text or ""
        for hit in re.finditer(re.escape("self.__next_f.push("), value):
            array, _ = decoder.raw_decode(value[hit.end():])
            if len(array) > 1 and isinstance(array[1], str):
                payload += array[1]
    datasets = []
    for hit in re.finditer(r'"rows":', payload):
        try:
            rows, _ = decoder.raw_decode(payload[hit.end():])
        except ValueError:
            continue
        if isinstance(rows, list) and rows and isinstance(rows[0], dict) and "size" in rows[0]:
            datasets.append(rows)
    return payload, datasets


def designation(raw, category, standard):
    value = raw.strip().removeprefix("L")
    value = re.sub(r"[xхXХ]", "×", value).replace(",", ".")
    result = {"raw": raw, "normalized": value, "fields": None, "mapping": None}
    if category == "truba" and re.fullmatch(r"\d+(?:\.\d+)?×\d+(?:\.\d+)?", value):
        d, s = map(float, value.split("×"))
        result.update(fields={"D": d, "s": s}, mapping="category explicitly defines diameter × wall")
    elif category == "profil":
        if re.fullmatch(r"\d+(?:\.\d+)?×\d+(?:\.\d+)?×\d+(?:\.\d+)?", value):
            h, b, t = map(float, value.split("×"))
            result.update(fields={"h": h, "b": b, "t": t}, mapping="A × B × wall; validated on detail/group samples")
        elif re.fullmatch(r"\d+(?:\.\d+)?×\d+(?:\.\d+)?-\d+(?:\.\d+)?", value) and "ГОСТ 8645-68" in standard:
            size, thickness = value.rsplit("-", 1)
            h, b = map(float, size.split("×"))
            result.update(apparent_fields={"h": h, "b": b, "t": float(thickness)}, mapping="h × b - t apparent; detail has no explicit dimension section; NOT_PROVED")
        elif re.fullmatch(r"\d+(?:\.\d+)?×\d+(?:\.\d+)?", value) and "ГОСТ 8639-82" in standard:
            h, t = map(float, value.split("×"))
            result.update(apparent_fields={"h": h, "b": h, "t": t}, mapping="square side × wall apparent; detail has no explicit dimension section; NOT_PROVED")
        elif re.fullmatch(r"\d+(?:\.\d+)?×\d+(?:\.\d+)?", value) and standard == "ТУ 36-2287-80":
            result["mapping"] = "two-number TU-only designation: shape mapping not independently proved"
    elif category == "ugolok":
        parts = value.split("×")
        if len(parts) in (2, 3) and all(number(x) is not None for x in parts):
            p = list(map(float, parts))
            result.update(fields={"b": p[0], "b2": p[0] if len(p) == 2 else p[1], "t": p[-1]},
                          mapping="angle side(s) × thickness; category and standard-group proof")
    elif category == "shveller":
        parts = value.split("×")
        if len(parts) in (3, 4) and all(number(x) is not None for x in parts):
            p = list(map(float, parts))
            result.update(fields={"h": p[0], "b": p[1], "t": p[-1], **({"b2": p[2]} if len(p) == 4 else {})},
                          mapping="bent channel designation: apparent mapping, not full-section proof")
    return result


entries = json.loads((BASE / "pages.json").read_text(encoding="utf-8"))
pages = []
for entry in entries:
    local = BASE / entry["local_path"]
    data = local.read_bytes()
    assert hashlib.sha256(data).hexdigest() == entry["sha256"]
    tree = html.fromstring(data)
    main = tree.xpath("//main")[0]
    tables = []
    for i, table in enumerate(main.xpath(".//table"), 1):
        headers = [text(x).removesuffix(" ⇅") for x in table.xpath(".//thead/tr[1]/th")]
        rows = []
        for n, row in enumerate(table.xpath(".//tbody/tr"), 1):
            cells = [text(c) for c in row.xpath("./td|./th")]
            rows.append({"tbody_row": n, "cells": cells,
                         "links": [a.get("href") for a in row.xpath(".//a[@href]")]})
        tables.append({"table_number": i, "headers": headers, "rows": rows, "row_count": len(rows)})
    payload, datasets = flight_data(tree)
    summary = {**entry, "h1": [text(x) for x in main.xpath(".//h1")],
               "headings": [text(x) for x in main.xpath(".//h2|.//h3")],
               "tables": tables, "embedded_row_datasets": datasets,
               "embedded_row_counts": [len(x) for x in datasets],
               "embedded_row_keys": [list(x[0]) for x in datasets],
               "pagination_links": [{"href":a.get("href"), "text":text(a)} for a in main.xpath(".//a[@href]") if re.search(r"[?&]page=", a.get("href"))],
               "linked_standards": [{"href":a.get("href"), "text":text(a)} for a in main.xpath(".//a[@href]") if a.get("href").startswith("/gost/")],
               "geometry_sections": []}
    for heading in main.xpath(".//h2"):
        title = text(heading)
        if title.startswith("Размеры сечения"):
            parent = heading.getparent()
            summary["geometry_sections"].append({"heading": title,
                "tables": [[[text(c) for c in r.xpath("./td|./th")] for r in tb.xpath(".//tr")] for tb in parent.xpath(".//table")]})
    if entry["kind"] == "category" and entry["url"] != "https://ptoflow.ru/sortament":
        category = entry["url"].rsplit("/", 1)[1]
        summary["standard_label_counts"] = dict(collections.Counter(x["cells"][1] for x in tables[0]["rows"]))
        summary["rows"] = [{"source":"PTOflow", "url":entry["url"], "retrieval_date":entry["timestamp_utc"],
                            "table_number":1, **row, "designation_mapping":designation(row["cells"][0], category, row["cells"][1])} for row in tables[0]["rows"]]
        summary["category_geometry_columns"] = []
        summary["html_count_equals_embedded"] = len(tables[0]["rows"]) == (len(datasets[0]) if datasets else -1)
    pages.append(summary)

(BASE / "parsed_pages.json").write_text(json.dumps(pages, ensure_ascii=False, indent=2), encoding="utf-8")
for page in pages:
    print(json.dumps({"url":page["url"],"kind":page["kind"],"table_counts":[t["row_count"] for t in page["tables"]],
                      "embedded_counts":page["embedded_row_counts"],"geometry_sections":page["geometry_sections"]}, ensure_ascii=False))
