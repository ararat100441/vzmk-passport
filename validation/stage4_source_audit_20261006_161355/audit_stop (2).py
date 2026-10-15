"""Bounded Stage 4 prerequisite audit. Writes evidence only; no product loader."""
from datetime import datetime, timezone
from hashlib import sha256
import importlib
import json
from pathlib import Path
import sys
from zipfile import ZipFile

sys.dont_write_bytecode = True
sys.stdout.reconfigure(encoding="utf-8")
OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
assert OUT.parent == ROOT / "validation" and OUT.name.startswith("stage4_")
sys.path.insert(0, str(ROOT))


def digest(path):
    path = Path(path).resolve()
    assert path.is_relative_to(ROOT), "Evidence audit cannot hash outside workspace"
    return sha256(path.read_bytes()).hexdigest()


def save(name, value):
    (OUT / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def stamp():
    return datetime.now(timezone.utc).isoformat()


started = stamp()
manifests = ("REVIEW_MANIFEST.json", "STAGE2_REVIEW_MANIFEST.json", "STAGE3_REVIEW_MANIFEST.json")
baseline = {}
historic_checks = []
for name in manifests:
    data = json.loads((ROOT / name).read_text(encoding="utf-8-sig"))
    entries = data["code_sha256"]
    mismatches = []
    for filename, expected in entries.items():
        path = Path(filename)
        if not path.is_absolute():
            path = ROOT / path
        actual = digest(path)
        baseline[str(path.resolve())] = expected
        if actual != expected:
            mismatches.append({"path": str(path), "expected": expected, "current": actual})
    historic_checks.append({"manifest": name, "checked": len(entries), "mismatches": mismatches})

gui_manifest_path = ROOT / "e2e_user_stage3/.state/gui_input_preview_fix/GUI_INPUT_PREVIEW_REVIEW_MANIFEST.json"
gui = json.loads(gui_manifest_path.read_text(encoding="utf-8"))
assert gui["product_hashes_before_tests"] == gui["product_hashes_after_tests"]
assert gui["targeted_validation"] == "PASS"
assert gui["actual_native_file_picker_and_error_dialog"] == "PASS"
for filename, expected in gui["evidence_sha256"].items():
    assert digest(filename) == expected
for filename, expected in gui["product_hashes_after_tests"].items():
    assert digest(filename) == expected
gui_path = str((ROOT / "passport_app/catalog_gui.py").resolve())
assert baseline[gui_path] == gui["original_gui_sha256"]
baseline[gui_path] = gui["current_gui_sha256"]
assert all(digest(p) == h for p, h in baseline.items()), "Unexpected accepted file change"

dev_checkpoint = ROOT / "DEV/DEV_REGISTRY_CHECKPOINT.json"
checkpoint = json.loads(dev_checkpoint.read_text(encoding="utf-8"))
dev = ROOT / "DEV/ВЗМК_DEV.xlsx"
assert Path(checkpoint["path"]).resolve() == dev.resolve()
assert digest(dev) == checkpoint["sha256"]
tracked = set(baseline) | {str(ROOT / n) for n in manifests} | {
    str(dev), str(dev_checkpoint), str(gui_manifest_path), str(Path(__file__).resolve()),
    str(OUT / "source_observations.json"), str(OUT / "gost_26020_83.observed_fragment.html"),
    str(OUT / "attempt1_429.py"), str(OUT / "attempt1_failure.json")}
before = {p: digest(p) for p in sorted(tracked)}
save("hashes_before.json", before)

smoke = {}
for module in ("passport_app.config", "passport_app.service", "passport_app.vik",
               "passport_app.catalog", "passport_app.catalog_gui", "main", "stage2", "stage3"):
    importlib.import_module(module)
    smoke[module] = "PASS"
with ZipFile(dev) as archive:
    assert archive.testzip() is None
    assert "xl/workbook.xml" in archive.namelist()
smoke["DEV registry checkpoint and XLSX ZIP integrity"] = "PASS"

observations = json.loads((OUT / "source_observations.json").read_text(encoding="utf-8"))
sources = observations["records"]
for source in sources:
    geometry = source["geometry_mm"]
    assert all(isinstance(v, (int, float)) and v > 0 for v in geometry.values())
    assert geometry["h"] > 2 * geometry["t"]
    wall = [geometry["h"] - 2 * geometry["t"], geometry["s"]]
    flanges = [2 * geometry["b"], geometry["t"]]
    source["audit_only_user_rule_arithmetic"] = {"wall_L_S": wall, "flanges_L_S": flanges}
    source["matches_confirmed_35B1_result"] = wall == [328, 6] and flanges == [348, 9]
assert sum(s["matches_confirmed_35B1_result"] for s in sources) == 3
assert len({tuple(s["geometry_mm"].values()) for s in sources}) == 2
extracted_evidence_hashes = {str(OUT / name): digest(OUT / name) for name in (
    "source_observations.json", "gost_26020_83.observed_fragment.html")}

source_manifest = {
    "scope": "SOURCE_AUDIT_ONLY_STOP_NO_MASS_IMPORT_AUTHORIZATION",
    "allowed_host": "23met.ru", "allowed_scheme": "https", "recursive_crawl": False,
    "audited_family": "Прокатные двутавры", "audited_designation": "35Б1",
    "exact_allowed_reference_pages": [s["url"] for s in sources],
    "source_records": sources,
    "source_extract_provenance": observations["provenance"],
    "retained_full_http_snapshots": "NOT_AVAILABLE: capture attempt HTTP 429; no further requests",
    "discovered_family_routes_not_fetched_after_STOP": [
        "https://23met.ru/spravka/tryba_vgp", "https://23met.ru/spravka/tryba_es_kvadr",
        "https://23met.ru/spravka/tryba_es_pr", "https://23met.ru/spravka/ygolok",
        "https://23met.ru/spravka/shveller", "https://23met.ru/spravka/ygolok_gnyt",
        "https://23met.ru/spravka/shveller_gnyt", "https://23met.ru/spravka/tryba_es_ploskooval",
        "https://23met.ru/spravka/shestigrannik"],
    "exclusions": ["Prices", "Suppliers", "Reinforcement", "Fasteners", "Wire",
                   "No extra families included before category/source audit"],
    "mass_import_manifest_ready": False,
    "deterministic_family_first_middle_last_samples": "NOT_RUN: prerequisite STOP before family import"
}
save("source_manifest.audit.json", source_manifest)
save("source_comparison_35B1.json", sources)

after = {p: digest(p) for p in sorted(tracked)}
assert before == after, "STOP: files changed during source audit"
save("hashes_after.json", after)
evidence = {
    "status": "STOP", "started_at_utc": started, "completed_at_utc": stamp(),
    "client_timezone": "Europe/Samara", "stage1_3_acceptance": "ACCEPTED per current user contract",
    "historic_manifest_checks": historic_checks,
    "historic_gui_mismatch_resolution": {"manifest": str(gui_manifest_path),
        "current_hash": gui["current_gui_sha256"], "existing_fix_evidence_hashes_match": True,
        "targeted_checks_recorded": gui["targeted_checks_passed"], "new_unexpected_change": False},
    "accepted_tracked_entries_checked": len(baseline),
    "accepted_current_baseline_matches": True, "smoke": smoke,
    "development_registry_sha256": digest(dev), "development_checkpoint_sha256": digest(dev_checkpoint),
    "product_and_checkpoint_hashes_before": before, "product_and_checkpoint_hashes_after": after,
    "before_equals_after": True, "source_snapshot_hashes": {},
    "source_extracted_evidence_sha256": extracted_evidence_hashes,
    "snapshot_capture": "INCOMPLETE_HTTP_429; attempt1_failure.json preserved; no further network requests",
    "stop_reason": "35Б1 has conflicting geometry across four actual source standards. ГОСТ 26020-83 produces 329x6.2 and 310x8.5 rather than the confirmed 328x6 and 348x9. Name-only lookup must be AMBIGUOUS; no automatic standard selection.",
    "stage4_product_files_created": [], "catalog_rows_imported": 0,
    "required_stage4_tests_1_to_30": "NOT_RUN: no product implementation after prerequisite STOP",
    "production_registry_access": "NOT_PERFORMED: no open/read/hash/write",
    "production_profile_workbook_write": "NOT_PERFORMED", "stage5": "NOT_RUN",
    "safe_next_step": "User must explicitly define the applicable standard or standard-aware 35Б1 acceptance; preserve each standard's actual geometry. Resume only under that clarified contract.",
    "ready_for_independent_implementation_review": False
}
save("audit_results.json", evidence)
print(json.dumps({"status": "STOP", "accepted_entries": len(baseline), "smoke": smoke,
                  "before_after_equal": before == after,
                  "source_geometry": [{"standard": s["standard"], "geometry": s["geometry_mm"],
                                      "result": s["audit_only_user_rule_arithmetic"]} for s in sources]}, ensure_ascii=False, indent=2))
