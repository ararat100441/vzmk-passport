"""Seal resumed prerequisite STOP audit; never label unfinished implementation PASS."""
from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
assert OUT.parent == ROOT / "validation" and OUT.name.startswith("stage4_")


def digest(path):
    path = Path(path).resolve()
    assert path.is_relative_to(ROOT)
    return sha256(path.read_bytes()).hexdigest()


result = json.loads((OUT / "audit_results.json").read_text(encoding="utf-8"))
before = json.loads((OUT / "accepted_hashes_before.json").read_text(encoding="utf-8"))
after = json.loads((OUT / "accepted_hashes_after.json").read_text(encoding="utf-8"))
assert before == after
assert all(digest(p) == h for p,h in before.items())
assert all(digest(p) == h for p,h in result["source_snapshots_sha256"].items())
smoke = json.loads((OUT / "smoke_results.json").read_text(encoding="utf-8"))
assert smoke["status"] == "PASS" and smoke["product_and_checkpoint_hashes"] == before
assert result["status"] == "SOURCE_IMPORT_STOP"
assert result["product_code_changes"] == [] and result["catalog_created"] is False
assert result["ready_for_independent_review"] is False
files = sorted(p for p in OUT.iterdir() if p.is_file() and p.name != "STAGE4_STOP_MANIFEST.json")
manifest = {"stage": 4, "scope": "RESUMED_SOURCE_AUDIT_ONLY", "status": "SOURCE_IMPORT_STOP",
    "ready_for_independent_review": False, "created_at_utc": datetime.now(timezone.utc).isoformat(),
    "integrity": "PASS_FOR_SOURCE_STOP_EVIDENCE_ONLY", "protected_hashes_current": before,
    "source_snapshots_sha256": result["source_snapshots_sha256"],
    "audit_evidence_sha256": {str(p): digest(p) for p in files},
    "tested_stage4_product": "NONE", "new_and_original_required_tests": "NOT_RUN",
    "network_stop": result["network_stop"], "parser_stop": result["deterministic_parser_stop"]}
assert all(digest(p) == h for p,h in before.items())
(OUT / "STAGE4_STOP_MANIFEST.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"status": manifest["status"], "audit_evidence_files": len(files),
    "source_snapshots": 7, "current_hashes_match": True, "ready_for_independent_review": False}, ensure_ascii=False))
