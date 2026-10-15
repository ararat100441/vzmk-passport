"""Verify and seal HTML proof's first-request HTTP 429 STOP. No network or parser."""
from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
assert OUT.parent == ROOT / "validation" and OUT.name.startswith("stage4_html_source_proof_")


def load(name):
    return json.loads((OUT / name).read_text(encoding="utf-8-sig"))


def digest(path):
    path = Path(path).resolve()
    assert path.is_relative_to(ROOT)
    return sha256(path.read_bytes()).hexdigest()


before = load("protected_hashes_before.json")
after = load("protected_hashes_after.json")
assert before == after
assert all(digest(path) == expected for path, expected in before.items())
fetch = load("fetch_results.json")
assert fetch["stopped"] is True
assert len(fetch["records"]) == 1
assert fetch["records"][0]["kind"] == "FAMILY_INDEX"
assert fetch["records"][0]["http_status"] == 429
assert fetch["records"][0]["result"] == "STOP_FETCH"
assert not list((OUT / "cache").iterdir())
files = sorted(p for p in OUT.iterdir() if p.is_file() and p.name != "HTML_SOURCE_PROOF_STOP_MANIFEST.json")
manifest = {
    "scope": "STAGE4_HTML_SOURCE_PROOF_2", "status": "STOP_HTTP_429",
    "created_at_utc": datetime.now(timezone.utc).isoformat(),
    "requested_index": fetch["records"][0]["url"], "http_status": 429,
    "total_live_requests": 1, "profile_requests": 0, "successful_html_cache_pages": 0,
    "parser_probe_created": False, "expected_values_used_as_output": False,
    "content_sufficiency": "NOT_TESTED", "proof_evidence_sufficiency": "HTML DATA NOT SUFFICIENT",
    "protected_before_after_current_equal": True, "protected_hashes_current": before,
    "source_html_snapshot_sha256": {},
    "evidence_sha256": {str(path): digest(path) for path in files},
    "network_after_429": "NONE", "product_changes": [], "production_registry_access": "NOT_PERFORMED",
    "PDF_OCR_glyph_vision": "NOT_RUN", "ready_for_mass_html_cache_downloader_stage": False,
    "stage5": "NOT_RUN"
}
assert all(digest(path) == expected for path, expected in before.items())
destination = OUT / "HTML_SOURCE_PROOF_STOP_MANIFEST.json"
destination.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"status": "STOP_HTTP_429", "evidence_files": len(files),
    "protected_hashes_equal": True, "ready": False}, ensure_ascii=False))
