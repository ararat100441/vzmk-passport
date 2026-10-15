"""Offline source-readiness diagnostics. No network, product code or XLSX writes."""
from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import re
import sys

from pypdf import PdfReader
import pypdfium2

sys.stdout.reconfigure(encoding="utf-8")
OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
assert OUT.parent == ROOT / "validation"
CACHE = ROOT / "DEV/Справочники/source_cache"


def hash_file(path):
    path = Path(path).resolve()
    assert path.is_relative_to(ROOT)
    return sha256(path.read_bytes()).hexdigest()


def write(name, data):
    (OUT / name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


previous = ROOT / "validation/stage4_source_audit_20261006_161355/STAGE4_STOP_MANIFEST.json"
previous_manifest = json.loads(previous.read_text(encoding="utf-8"))
for path, expected in previous_manifest["audit_evidence_sha256"].items():
    assert hash_file(path) == expected, "Previous audit evidence changed"
accepted = json.loads((previous.parent / "hashes_after.json").read_text(encoding="utf-8"))
for path, expected in accepted.items():
    assert hash_file(path) == expected, "Accepted product/DEV checkpoint changed"
before = {path: hash_file(path) for path in accepted}
write("accepted_hashes_before.json", before)
fetch_records = json.loads((OUT / "bulk_fetch_results.json").read_text(encoding="utf-8-sig"))
diagnostics = []
for source in fetch_records:
    if source["status"] != "FETCHED_NOT_PARSED":
        diagnostics.append(source | {"parser_status": "STOP_FETCH_429", "eligible_for_mass_import": False})
        continue
    path = Path(source["snapshot_path"])
    assert path.resolve().is_relative_to(CACHE) and hash_file(path) == source["sha256"]
    reader = PdfReader(path)
    texts = [page.extract_text() for page in reader.pages]
    glyph_count = sum(len(re.findall(r"/G[0-9a-fA-F]+", t)) for t in texts)
    font_diagnostics = []
    for number, page in enumerate(reader.pages):
        fonts = page.get("/Resources", {}).get("/Font", {})
        if hasattr(fonts, "get_object"):
            fonts = fonts.get_object()
        for name, ref in fonts.items():
            font = ref.get_object()
            encoding = font.get("/Encoding")
            if hasattr(encoding, "get_object"):
                encoding = encoding.get_object()
            if isinstance(encoding, dict) and any(str(v).startswith("/G") for v in encoding.get("/Differences", [])):
                font_diagnostics.append({"page": number + 1, "font_resource": name,
                    "base_font": str(font.get("/BaseFont")), "type": str(font.get("/Subtype")),
                    "has_ToUnicode": bool(font.get("/ToUnicode")),
                    "custom_glyph_names": True})
    if sum(map(len, texts)) == 0:
        status = "STOP_IMAGE_ONLY_NO_VERIFIED_TABLE_PARSER"
    elif glyph_count:
        status = "STOP_CUSTOM_FONT_NO_VERIFIED_UNICODE_TABLE_PARSER"
    else:
        status = "TEXT_LAYER_AVAILABLE_NOT_YET_APPROVED_FOR_MASS_IMPORT"
    diagnostic = source | {"family": "Прокатные двутавры" if source["standard"] in (
        "ГОСТ Р 57837-2017", "СТО АСЧМ 20-93", "ГОСТ 26020-83", "ГОСТ 35087-2024") else
        {"ГОСТ 8240-97": "Швеллеры", "ГОСТ 8509-93": "Уголки равнополочные",
         "ГОСТ 8510-86": "Уголки неравнополочные"}.get(source["standard"], "Не определено"),
        "page_count": len(reader.pages), "extractable_text_characters": sum(map(len, texts)),
        "raw_custom_glyph_token_count": glyph_count, "font_diagnostics": font_diagnostics,
        "parser_status": status, "eligible_for_mass_import": False,
        "expected_fields": ["profile", "h", "b", "s", "t", "r"] if source["standard"] in (
          "ГОСТ Р 57837-2017", "СТО АСЧМ 20-93", "ГОСТ 26020-83", "ГОСТ 35087-2024", "ГОСТ 8240-97")
          else ["profile", "b", "b2", "s", "r"]}
    # Keep a bounded raw decoder example. Never translate it using guessed glyph mappings.
    diagnostic["raw_text_sample"] = next((t[:600] for t in texts if t), "")
    if source["standard"] in ("ГОСТ Р 57837-2017", "СТО АСЧМ 20-93"):
        hits = []
        for page_number, text in enumerate(texts):
            for line in text.splitlines():
                if re.match(r"^35\s*Б1\s", line):
                    hits.append({"page": page_number + 1, "raw_source_line": line})
        diagnostic["35B1_raw_source_rows"] = hits
    diagnostics.append(diagnostic)
    if source["standard"] in ("ГОСТ 26020-83", "ГОСТ 8240-97", "ГОСТ 35087-2024"):
        document = pypdfium2.PdfDocument(path)
        page_index = 1 if len(document) > 1 else 0
        image = document[page_index].render(scale=1.7).to_pil()
        preview = OUT / (path.stem + ".page" + str(page_index + 1) + ".png")
        image.save(preview)
        diagnostic["visual_source_preview"] = str(preview)
        document.close()

decision = {"source_identity": ["normalized_profile_key", "normalized_standard"],
    "preserve_all_source_variants": True,
    "35B1_preference": {"Ключ": "35Б1", "Предпочтительный НД": "ГОСТ Р 57837-2017",
        "Основание": "USER_CONFIRMED_UZK_328x6_348x9", "Статус": "ПОДТВЕРЖДЕНО ПОЛЬЗОВАТЕЛЕМ",
        "Дата": "2026-10-06"},
    "lookup_contract": {"explicit_standard": "Exact normalized profile+standard; FOUND/NOT_FOUND; no fallback",
        "one_geometry_no_standard": "FOUND; preserve all equivalent provenance",
        "multiple_geometries_valid_preference": "FOUND_PREFERRED; resolved_by_user_preference=true",
        "multiple_geometries_no_preference": "AMBIGUOUS; return standards and geometry",
        "missing_or_conflicting_preference_target": "ERROR"},
    "workbook_sheets": ["Профили", "Правила УЗК", "Источники", "Предпочтения"],
    "applied_to_product_or_workbook": False,
    "previous_audit_observations": str(previous.parent / "source_observations.json")}
write("standard_aware_decision.json", decision)
write("source_manifest.audit.json", {"scope": "RESUMED_SOURCE_AUDIT_ONLY", "sources": diagnostics,
    "network_requests_stopped_after_429": True, "recursive_crawl": False, "throttle_seconds": 4,
    "not_fetched_after_429": ["ГОСТ 8645-68", "ГОСТ 10704-91", "ГОСТ 30245-2003",
        "ГОСТ 2879-2006", "ГОСТ 8239-89"],
    "mass_import_ready": False})
after = {path: hash_file(path) for path in accepted}
assert before == after
write("accepted_hashes_after.json", after)
write("audit_results.json", {"status": "SOURCE_IMPORT_STOP", "completed_at_utc": datetime.now(timezone.utc).isoformat(),
    "standard_aware_decision": "USER_AUTHORIZED_RECORDED_NOT_IMPLEMENTED",
    "downloaded_bulk_snapshots": sum(s["status"] == "FETCHED_NOT_PARSED" for s in fetch_records),
    "source_snapshots_sha256": {s["snapshot_path"]: s["sha256"] for s in fetch_records if s["status"] == "FETCHED_NOT_PARSED"},
    "network_stop": "HTTP 429 on ГОСТ 8639-82 reference page; no subsequent live requests",
    "deterministic_parser_stop": "Four PDFs expose custom /G glyph names without Unicode mapping; ГОСТ 35087-2024 has zero text layer. No verified deterministic dimension parser for those snapshots; do not guess or import them.",
    "all_current_protected_hashes_before_after_equal": True, "protected_entries": len(before),
    "product_code_changes": [], "catalog_created": False,
    "stage4_targeted_tests": "NOT_RUN", "production_access": "NOT_PERFORMED", "stage5": "NOT_RUN",
    "ready_for_independent_review": False})
print(json.dumps({"status": "SOURCE_IMPORT_STOP", "bulk_snapshots": 7,
    "sources": [{"standard": d["standard"], "parser_status": d["parser_status"]} for d in diagnostics],
    "protected_hashes_equal": before == after}, ensure_ascii=False, indent=2))
