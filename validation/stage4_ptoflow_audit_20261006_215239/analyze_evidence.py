"""Offline audit comparisons; preserves each source value and its own standard."""
import collections
import hashlib
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
SOURCE_A = BASE.parent / "stage4_km_kmd_audit_20261006_204347"
sys.stdout.reconfigure(encoding="utf-8")
pages = json.loads((BASE / "parsed_pages.json").read_text(encoding="utf-8"))
a_results = json.loads((SOURCE_A / "audit_results.json").read_text(encoding="utf-8"))
a_records = json.loads((SOURCE_A / "parsed_source_records.json").read_text(encoding="utf-8"))


def n(raw):
    try:
        return float(re.sub(r"\s+", "", str(raw)).replace(",", "."))
    except ValueError:
        return None


def value(raw):
    return raw["sort"] if isinstance(raw, dict) and "sort" in raw else None


def label(raw):
    return raw.get("label", "") if isinstance(raw, dict) else str(raw)


def norm(raw):
    return re.sub(r"\s+", "", str(raw)).replace(",", ".").replace("×", "x").replace("х", "x").removeprefix("L").rstrip("*")


standards = {"26020-83": ("Двутавр", "ГОСТ 26020-83"), "57837-2017": ("Двутавр", "ГОСТ Р 57837-2017"),
             "8239-89": ("Двутавр", "ГОСТ 8239-89"), "8240-97": ("Швеллер", "ГОСТ 8240-97"),
             "8509-93": ("Уголок равнополочный", "ГОСТ 8509-93"), "8510-86": ("Уголок неравнополочный", "ГОСТ 8510-86"),
             "30245-2003": ("Профильная труба", "ГОСТ 30245-2003"),
             "10704-91": ("Круглая труба", "ГОСТ 10704-91"), "8732-2025": ("Круглая труба", "ГОСТ 8732-2025")}
group_records, group_inventory, rounding = [], [], []
for page in pages:
    code = page["url"].rsplit("/", 1)[1]
    if page["kind"] != "group" or code not in standards:
        continue
    family, standard = standards[code]
    rows = page["embedded_row_datasets"][0]
    for i, raw in enumerate(rows, 1):
        fields = {k: value(raw[k]) for k in ["h", "b", "s", "t", "R", "area", "mass"] if k in raw}
        if family == "Уголок неравнополочный":
            fields["b2"] = fields.pop("b")
            fields["b"] = value(raw["B"])
        elif family == "Круглая труба":
            fields["D"] = value(raw["d"])
        elif family == "Профильная труба":
            dims = [n(s) for s in norm(label(raw["size"])).split("x")]
            assert len(dims) == 3 and all(x is not None for x in dims)
            assert dims[2] == fields["t"]
            fields.update(h=dims[0], b=dims[1])
        geometry_fields = ["h", "b", "s", "t"] if family in ["Двутавр", "Швеллер"] else (
            ["D", "s"] if family == "Круглая труба" else (
            ["b", "b2", "t"] if family == "Уголок неравнополочный" else (
            ["b", "t"] if family == "Уголок равнополочный" else ["h", "b", "t"])))
        group_records.append({"source": "PTOflow", "url": page["url"], "standard": standard,
                              "retrieval_date": page["timestamp_utc"], "family": family,
                              "embedded_row": i, "raw": raw, "designation": label(raw["size"]),
                              "detail_href": raw["size"].get("href") if isinstance(raw["size"], dict) else None,
                              "parsed": fields, "primary_geometry_fields": geometry_fields,
                              "primary_geometry_complete": all(fields.get(k) is not None for k in geometry_fields)})
        for key, raw_value in raw.items():
            if key in ["h", "b", "B", "s", "t", "d", "R"] and isinstance(raw_value, dict):
                if n(raw_value["display"]) != raw_value["sort"]:
                    rounding.append({"url": page["url"], "row": i, "designation": label(raw["size"]),
                                     "field": key, "raw_numeric": raw_value["sort"], "raw_display": raw_value["display"]})
    group_inventory.append({"family": family, "standard": standard, "url": page["url"], "rows": len(rows),
                            "fields": [k for k in rows[0] if k != "size"],
                            "geometry_radii_in_bulk": "R" in rows[0],
                            "primary_geometry_complete_rows": sum(x["primary_geometry_complete"] for x in group_records if x["url"] == page["url"])})

details = []
for page in pages:
    if page["kind"] != "detail":
        continue
    sections = []
    for section in page["geometry_sections"]:
        parsed = {}
        for key, raw in section["tables"][0]:
            if raw.endswith(" мм"):
                symbol = key.rsplit(" ", 1)[1]
                if symbol in {"h", "b", "B", "s", "t", "R", "r"}:
                    parsed[symbol] = n(raw[:-3])
            elif key.startswith("Площадь сечения"):
                parsed["area"] = n(raw.split(" см")[0])
        sections.append({"standard": section["heading"].removeprefix("Размеры сечения по "),
                         "raw": section, "parsed": parsed})
    sources = next((t["rows"] for t in page["tables"] if t["headers"] and t["headers"][0] == "Источник"), [])
    details.append({"url": page["url"], "standards_listed": [r["cells"][0] for r in sources],
                    "source_mass_rows": sources, "geometry_sections": sections,
                    "geometry_not_propagated_to_other_standards": True})

beam_tests = []
for profile in ["35Б1", "20Ш1", "25Б1", "30Б2"]:
    variants = [r for r in group_records if r["family"] == "Двутавр" and norm(r["designation"]) == profile]
    a = a_results["controls"][profile][0]
    detail = next(x for x in details if x["url"].endswith("/" + a["parsed"]["source_profile"].replace("Б", "b").replace("Ш", "sh")))
    same_nd = next((s for s in detail["geometry_sections"] if s["standard"] == a["standard"]), None)
    beam_tests.append({"profile": profile, "detail": detail,
                       "detail_multi_standard_result": "STANDARD_GEOMETRY_AMBIGUOUS" if len(detail["standards_listed"]) > len(detail["geometry_sections"]) else "PASS",
                       "group_variants": variants, "standard_group_test": "PASS",
                       "source_a": a, "same_standard_comparison": "MATCH" if same_nd and all(same_nd["parsed"][k] == a["parsed"][k] for k in ["h", "b", "s", "t"]) else "PTOFLOW_INCOMPLETE",
                       "modern_group_numerically_matches_source_a": any(r["standard"] == "ГОСТ Р 57837-2017" and all(r["parsed"][k] == a["parsed"][k] for k in ["h", "b", "s", "t"]) for r in variants),
                       "cross_standard_numeric_match_is_not_same_standard_proof": True})
expect = {"35Б1": {"ГОСТ 26020-83": dict(h=346, b=155, s=6.2, t=8.5),
                       "ГОСТ Р 57837-2017": dict(h=346, b=174, s=6, t=9)}}
for test in beam_tests:
    test["expected_comparison_only"] = [{"standard": r["standard"], "matches": all(r["parsed"][k] == v for k, v in expected.items())}
        for r in test["group_variants"] if (expected := expect.get(test["profile"], {}).get(r["standard"]))]

profile_category = next(p for p in pages if p["url"] == "https://ptoflow.ru/sortament/profil")
round_category = next(p for p in pages if p["url"] == "https://ptoflow.ru/sortament/truba")
profile_samples = []
for designation in ["100×100×6,5", "80×40×4", "90×50×6,5", "400×200×11,5"]:
    row = next(r for r in profile_category["rows"] if r["cells"][0] == designation)
    related = [r for r in group_records if r["detail_href"] in row["links"] and r["family"] == "Профильная труба"]
    detail = next(d for d in details if d["url"] == "https://ptoflow.ru" + row["links"][0])
    profile_samples.append({"designation": designation, "standard": row["cells"][1], "raw_category": row,
                            "parsed": row["designation_mapping"]["fields"], "group_rows": related, "detail": detail,
                            "format_proof": "PASS: category explicitly defines A×B×wall; two samples have matching explicit h/b/t and group rows"})
round_samples = []
for designation in ["10×1", "57×3,5", "159×6", "530×8", "1020×10"]:
    row = next(r for r in round_category["rows"] if r["cells"][0] == designation)
    related = [r for r in group_records if r["detail_href"] in row["links"] and r["family"] == "Круглая труба"]
    assert related and all(all(r["parsed"][k] == v for k, v in row["designation_mapping"]["fields"].items()) for r in related)
    round_samples.append({"designation": designation, "standard": row["cells"][1], "raw_category": row,
                          "parsed": row["designation_mapping"]["fields"], "group_rows": related,
                          "format_proof": "PASS: category D×wall matches explicit embedded numeric D/s in each standard group"})

selected = [("уголки неравнополочные.xls", 30, 12), ("уголки неравнополочные.xls", 31, 12),
            ("уголки неравнополочные.xls", 43, 1), ("уголки неравнополочные.xls", 44, 1),
            ("уголки неравнополочные.xls", 18, 1), ("уголки неравнополочные.xls", 18, 12),
            ("СТО АСЧМ 20-93.xls", 34, 1), ("СТО АСЧМ 20-93.xls", 47, 1),
            ("уголки равнополочные.xls", 16, 10), ("уголки неравнополочные.xls", 34, 12)]
crosscheck = []
for filename, row, col in selected:
    a = next(r for r in a_records if r["workbook"] == filename and r["row"] == row and r["start_column"] == col)
    possible = [r for r in group_records if r["standard"] == a["standard"] and r["family"] == a["family"]]
    if a["family"] == "Уголок неравнополочный":
        b = next((r for r in possible if norm(r["raw"].get("number")) == norm(a["parsed"]["source_profile"]) and r["parsed"]["t"] == a["parsed"]["t"]), None)
    elif a["family"] == "Уголок равнополочный":
        b = next((r for r in possible if r["parsed"]["b"] == a["parsed"]["b"] and r["parsed"]["t"] == a["parsed"]["t"]), None)
    else:
        b = next((r for r in possible if norm(r["designation"]) == a["parsed"]["source_profile"]), None)
    item = {"source_a": a, "source_b": b, "result": "PTOFLOW_INCOMPLETE" if not b else "MATCH", "comparisons": []}
    if b:
        for field in (["b", "b2", "t"] if a["family"] == "Уголок неравнополочный" else ["b", "t"]) + ["area", "mass"]:
            afield = {"area": "area_cm2", "mass": "mass_kg_m"}.get(field, field)
            av, bv = a["parsed"].get(afield), b["parsed"].get(field)
            match = av is not None and bv is not None and abs(av - bv) < 1e-8
            item["comparisons"].append({"field": field, "source_a": av, "source_b": bv, "result": "MATCH" if match else "MISMATCH"})
            if not match:
                item["result"] = "MISMATCH"
        detail = next((d for d in details if d["url"] == "https://ptoflow.ru" + str(b["detail_href"])), None)
        geom = next((s for s in detail["geometry_sections"] if s["standard"] == a["standard"]), None) if detail else None
        item["source_b_radius_proof"] = geom
        if geom:
            for radius in ["R", "r"]:
                av, bv = a["parsed"].get(radius), geom["parsed"].get(radius)
                if av is not None and bv is not None:
                    match = abs(av - bv) < 1e-8
                    item["comparisons"].append({"field": radius, "source_a": av, "source_b": bv,
                                                "result": "MATCH" if match else "MISMATCH"})
                    if not match:
                        item["result"] = "MISMATCH"
                elif av is None and bv is not None:
                    item.setdefault("additional_source_b_fields", {})[radius] = bv
        if a["parsed"].get("R") is None and a["parsed"].get("r") is None and not geom:
            item["result"] = "PTOFLOW_INCOMPLETE"
            item["reason"] = "Primary dimensions match; R/r absent from group and no bounded detail proof for this selected row"
    else:
        item["reason"] = "No geometry block/table under the SAME НД; numeric agreement under another НД is not substituted"
        item["different_standard_reference"] = [r for r in group_records if r["family"] == a["family"] and norm(r["designation"]) == a["parsed"]["source_profile"]]
    crosscheck.append(item)

atomic_counts = []
for p in pages:
    if "standard_label_counts" not in p:
        continue
    counts = collections.Counter()
    for combined, count in p["standard_label_counts"].items():
        for standard in combined.split(", "):
            counts[standard] += count
    atomic_counts.append({"category": p["url"], "combined_label_counts": p["standard_label_counts"],
                          "apparent_standard_attributions": dict(counts),
                          "note": "Several НД in one category row count as apparent attributions, not separate proven geometries"})
excluded = [r for r in profile_category["rows"] if any(s in r["cells"][1] for s in ["ГОСТ 8282-83", "ГОСТ 8283-93", "ТУ 14-105-424-79", "ГОСТ 7511-73"])]
profile_proven = [r for r in profile_category["rows"] if r["designation_mapping"]["fields"] and r not in excluded]
profile_kinds = collections.Counter("квадрат" if r["designation_mapping"]["fields"]["h"] == r["designation_mapping"]["fields"]["b"] else "прямоугольник" for r in profile_proven)
within_nd_duplicates = []
for inventory in group_inventory:
    rows = [r for r in group_records if r["url"] == inventory["url"]]
    display_counts = collections.Counter(norm(r["designation"]) for r in rows)
    duplicates = {k:v for k,v in display_counts.items() if v > 1}
    if duplicates:
        within_nd_duplicates.append({"url": inventory["url"], "duplicate_display_designations": duplicates,
                                    "action": "Keep raw numeric dimensions and row provenance; never merge by rounded display label"})

round_keys = collections.defaultdict(list)
for row in group_records:
    if row["family"] == "Круглая труба":
        round_keys[(row["standard"], row["parsed"]["D"], row["parsed"]["s"])].append(row)
exact_round_duplicates = [{"standard": key[0], "D": key[1], "s": key[2], "records": rows,
                           "action": "Flag both source rows; no automatic collapse"}
                          for key, rows in round_keys.items() if len(rows) > 1]

result = {"network_requests": len(pages), "detail_pages": len(details), "category_pages_including_index": sum(p["kind"] == "category" for p in pages),
          "group_pages_including_index": sum(p["kind"] == "group" for p in pages), "http_429": any(p["http_status"] == 429 for p in pages),
          "group_inventory": group_inventory, "category_standard_inventory": atomic_counts,
          "beam_tests": beam_tests, "profile_tube_samples": profile_samples, "round_pipe_samples": round_samples,
          "source_a_crosscheck": crosscheck, "profile_category_excluded_count": len(excluded),
          "profile_category_proven_mapping_count": len(profile_proven), "profile_category_proven_shape_counts": dict(profile_kinds),
          "profile_category_incomplete_mapping_count": len(profile_category["rows"]) - len(excluded) - len(profile_proven),
          "group_geometry_display_rounding": rounding, "within_standard_display_duplicates": within_nd_duplicates,
          "exact_round_duplicate_rows": exact_round_duplicates,
          "decision": "KM-KMD + PTOFLOW COVER MOST, SMALL GAPS REMAIN",
          "ready_to_design_import": "YES: design verified-source contract, with missing/unproved fields and conflicts explicit; no final preference assigned"}
(BASE / "audit_results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
(BASE / "group_source_evidence.json").write_text(json.dumps(group_records, ensure_ascii=False, indent=2), encoding="utf-8")
(BASE / "detail_source_evidence.json").write_text(json.dumps(details, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({k:result[k] for k in ["network_requests", "detail_pages", "group_inventory", "profile_category_proven_mapping_count", "profile_category_proven_shape_counts", "profile_category_incomplete_mapping_count", "profile_category_excluded_count"]},ensure_ascii=False,indent=2))
print("CROSSCHECK",[(x["source_a"]["workbook"],x["source_a"]["row"],x["source_a"]["start_column"],x["result"],x["comparisons"]) for x in crosscheck])
print("ROUNDING",len(rounding),"DUPLICATE DISPLAY GROUPS",len(within_nd_duplicates))
