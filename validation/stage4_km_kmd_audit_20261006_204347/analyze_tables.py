"""Audit only: explicit observed table blocks, provenance, no source correction."""
import collections
import hashlib
import json
import math
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
BASE = Path(__file__).resolve().parent
books = json.loads((BASE / "raw_workbook_inventory.json").read_text(encoding="utf-8"))
records, inventories = [], []


def number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if math.isfinite(value) else None
    text = str(value).strip().replace("\u00a0", "")
    # The source footnote explicitly explains asterisks; preserve them in raw data.
    if re.fullmatch(r"[+-]?\d+(?:[,.]\d+)?\*?", text):
        return float(text.rstrip("*").replace(",", "."))
    return None


def compact(value):
    return re.sub(r"\s+", "", str(value)).replace("\u00ad", "")


for book in books:
    sheet = book["sheets"][0]
    cells = {(r["row"], c["col"]): c for r in sheet["rows"] for c in r["cells"]}
    title = cells[(1, 1)]["value"]
    standard = re.search(r"(?:ГОСТ|СТО\s+АСЧМ)\s+\d+-\d+", title).group(0)
    standard = re.sub(r"\s+", " ", standard)
    filename = book["filename"]
    if "неравнополочные" in filename:
        family = "Уголок неравнополочный"
        blocks = [("левый", 18, 51, 1, ["source_profile", "b", "b2", "t", "R", "r", "area_cm2", "x0_cm", "y0_cm", "mass_kg_m"]),
                  ("правый", 18, 45, 12, ["source_profile", "b", "b2", "t", "R", "r", "area_cm2", "x0_cm", "y0_cm", "mass_kg_m"])]
        header_rows = [15, 16, 17]
    elif "равнополочные" in filename:
        family = "Уголок равнополочный"
        blocks = [("левый", 21, 57, 1, ["source_profile", "b", "t", "R", "r", "area_cm2", "x0_cm", "mass_kg_m"]),
                  ("правый", 6, 57, 10, ["source_profile", "b", "t", "R", "r", "area_cm2", "x0_cm", "mass_kg_m"])]
        header_rows = [3, 4, 5, 18, 19, 20]
    elif "швеллера" in filename:
        family = "Швеллер"
        fields = ["source_profile", "h", "b", "s", "t", "R", "r", "area_cm2", "mass_kg_m", "x0_cm"]
        blocks = [("У", 6, 23, 1, fields), ("П", 29, 46, 1, fields)]
        header_rows = [3, 4, 5, 26, 27, 28]
    else:
        family = "Двутавр"
        fields = ["source_profile", "h", "b", "s", "t", "R", "area_cm2", "mass_kg_m"]
        blocks = [("К", 6, 21, 1, fields), ("Б", 26, 55, 1, fields), ("Ш", 26, 55, 10, fields)]
        header_rows = [3, 4, 5, 23, 24, 25]

    def resolve(r, c):
        raw = cells.get((r, c), {}).get("value")
        if raw is not None:
            return raw, [r, c], "cell"
        for rlo, rhi, clo, chi in sheet["merged_ranges_zero_based"]:
            if rlo <= r - 1 < rhi and clo <= c - 1 < chi:
                anchor = [rlo + 1, clo + 1]
                return cells.get(tuple(anchor), {}).get("value"), anchor, "merged_anchor"
        return None, [r, c], "empty"

    count_before = len(records)
    for subtype, first, last, start, fields in blocks:
        previous_group = {}
        for row in range(first, last + 1):
            raw, effective, provenance = {}, {}, {}
            for offset, field in enumerate(fields):
                col = start + offset
                raw[field] = cells.get((row, col), {}).get("value")
                val, origin, method = resolve(row, col)
                # Equal angles use unmerged blank labels within explicit printed groups.
                # Do not propagate radii or any other field through arbitrary blank cells.
                if family == "Уголок равнополочный" and field in {"source_profile", "b"}:
                    if val is not None:
                        previous_group[field] = (val, origin)
                    elif field in previous_group:
                        val, origin = previous_group[field]
                        method = "same_printed_profile_group"
                effective[field] = val
                provenance[field] = {"cell": [row, col], "origin": origin, "method": method}
            if effective["source_profile"] is None or number(effective["t"]) is None:
                raise ValueError(f"Profile block contains service/invalid row: {filename}:{row}:{start}")
            parsed = {k: number(v) for k, v in effective.items() if k != "source_profile"}
            parsed["source_profile"] = compact(effective["source_profile"])
            # Audit identity includes angle thickness because source numbers designate sizes only.
            key = parsed["source_profile"]
            if "Уголок" in family:
                key = f"{parsed['b']:g}x" + (f"{parsed['b2']:g}x" if "b2" in parsed else "") + f"{parsed['t']:g}"
            records.append({"family": family, "standard": standard, "subtype": subtype,
                            "workbook": filename, "sheet": sheet["name"], "row": row,
                            "start_column": start, "raw": raw, "effective": effective,
                            "parsed": parsed, "provenance": provenance,
                            "identity": [key, standard]})
    count = len(records) - count_before
    inventories.append({"workbook": filename, "format": book["format"], "biff_version": book["biff_version"],
                        "family": family, "standard": standard, "standard_source_cell": "Лист1!A1",
                        "headers": [r for r in sheet["rows"] if r["row"] in header_rows],
                        "profile_records": count,
                        "sheets": [{k: s[k] for k in ["name", "allocated_rows", "allocated_columns", "used_bounds", "nonempty_row_count", "merged_ranges_zero_based"]} for s in book["sheets"]]})

records.sort(key=lambda x: (x["family"], x["standard"], x["workbook"], x["sheet"], x["row"], x["start_column"]))
groups = collections.defaultdict(list)
for item in records:
    groups[(item["family"], item["standard"])].append(item)
coverage, samples, anomalies = [], [], []
for (family, standard), group in groups.items():
    required = ["h", "b", "s", "t", "R"] if family == "Двутавр" else (
        ["h", "b", "s", "t", "R", "r"] if family == "Швеллер" else (
        ["b", "b2", "t", "R", "r"] if family == "Уголок неравнополочный" else ["b", "t", "R", "r"]))
    missing = {f: [{"workbook": x["workbook"], "row": x["row"], "start_column": x["start_column"]} for x in group if x["parsed"][f] is None] for f in required}
    invalid = [{"row": x["row"], "start_column": x["start_column"], "field": k, "raw": v} for x in group for k, v in x["effective"].items() if k != "source_profile" and v is not None and number(v) is None]
    text_numbers = [{"row": x["row"], "start_column": x["start_column"], "field": k, "raw": v, "parsed": number(v)} for x in group for k, v in x["effective"].items() if k != "source_profile" and isinstance(v, str)]
    coverage.append({"family": family, "standard": standard, "rows": len(group), "geometry_fields": required,
                     "complete_geometry_rows": sum(all(x["parsed"][f] is not None for f in required) for x in group),
                     "missing": missing, "invalid_numeric": invalid, "text_numeric_count": len(text_numbers),
                     "text_numeric_examples": text_numbers[:6],
                     "subtypes": dict(collections.Counter(x["subtype"] for x in group))})
    for label, index in [("first", 0), ("middle", (len(group) - 1) // 2), ("last", len(group) - 1)]:
        samples.append({"selection": label, "ordinal": index + 1, "group_size": len(group), **group[index]})
    for x in group:
        p = x["parsed"]
        if "Уголок" in family:
            pieces = re.findall(r"\d+(?:[,.]\d+)?", p["source_profile"])
            for i, field in enumerate(["b", "b2"] if "b2" in p else ["b"]):
                if i < len(pieces) and number(pieces[i]) * 10 != p[field]:
                    anomalies.append({"kind": "profile_number_vs_dimension", "workbook": x["workbook"], "row": x["row"],
                                      "start_column": x["start_column"], "source_profile": p["source_profile"], "field": field,
                                      "actual": p[field], "number_times_10": number(pieces[i]) * 10,
                                      "action": "retain source value; do not correct"})
        # A plausibility check only. Source states steel density 7.85 g/cm3.
        area, mass = p["area_cm2"], p["mass_kg_m"]
        if area and mass and abs(mass / (area * 0.785) - 1) > 0.03:
            anomalies.append({"kind": "area_mass_disagreement_over_3_percent", "workbook": x["workbook"], "row": x["row"],
                              "start_column": x["start_column"], "source_profile": p["source_profile"],
                              "area_cm2": area, "mass_kg_m": mass, "mass_check_only": area * 0.785,
                              "action": "retain source values; do not correct"})

controls = {profile: [x for x in records if x["family"] == "Двутавр" and x["parsed"]["source_profile"] == profile]
            for profile in ["35Б1", "20Ш1", "25Б1", "30Б2"]}
expected_comparison = []
for x in controls["35Б1"]:
    actual = {f: x["parsed"][f] for f in ["h", "b", "s", "t"]}
    expected_comparison.append({"actual": actual, "matches_geometry_group": actual == dict(h=346, b=174, s=6, t=9),
                                "matches_gost_26020_83": actual == dict(h=346, b=155, s=6.2, t=8.5),
                                "expected_used_only_for_comparison": True})
duplicate_groups = collections.defaultdict(list)
for x in records:
    duplicate_groups[tuple(x["identity"])].append({"workbook": x["workbook"], "row": x["row"], "start_column": x["start_column"]})
results = {"inventory": inventories, "coverage": coverage, "total_profile_records": len(records),
           "sample_rule": "Sort by source row then start column within each family/standard; lower middle ordinal floor((n-1)/2)+1",
           "samples": samples, "controls": controls, "comparison_35B1": expected_comparison,
           "anomalies": anomalies,
           "duplicate_identity_records": [{"identity": list(k), "sources": v} for k, v in duplicate_groups.items() if len(v) > 1],
           "decision": "KM-KMD CAN BE PARTIAL BASE"}
(BASE / "audit_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
(BASE / "parsed_source_records.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"coverage": coverage, "total": len(records), "controls": {k: [{f: x[f] for f in ["workbook", "sheet", "row", "start_column", "parsed"]} for x in v] for k, v in controls.items()},
                  "anomalies": anomalies, "duplicates": results["duplicate_identity_records"],
                  "samples": [{"selection": x["selection"], "family": x["family"], "standard": x["standard"], "workbook": x["workbook"], "row": x["row"], "start_column": x["start_column"], "parsed": x["parsed"]} for x in samples]}, ensure_ascii=False, indent=2))
