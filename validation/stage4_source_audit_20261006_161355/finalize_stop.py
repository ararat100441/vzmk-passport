"""Seal STOP evidence only; never publish a successful implementation manifest."""
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
before = result["product_and_checkpoint_hashes_before"]
after = result["product_and_checkpoint_hashes_after"]
assert before == after
assert all(digest(path) == expected for path, expected in before.items())
assert result["status"] == "STOP"
assert result["stage4_product_files_created"] == []
assert result["catalog_rows_imported"] == 0
assert result["ready_for_independent_implementation_review"] is False
assert all(digest(path) == expected for path, expected in result["source_extracted_evidence_sha256"].items())
paths = sorted(p for p in OUT.iterdir() if p.is_file() and p.name != "STAGE4_STOP_MANIFEST.json")
evidence_hashes = {str(p): digest(p) for p in paths}
manifest = {"stage": 4, "status": "STOP", "scope": "PREREQUISITE_SOURCE_AUDIT_ONLY",
    "ready_for_independent_review": False, "created_at_utc": datetime.now(timezone.utc).isoformat(),
    "evidence_integrity": "PASS_FOR_STOP_AUDIT_ONLY", "audit_evidence_sha256": evidence_hashes,
    "accepted_product_checkpoint_before_after_current_equal": True,
    "tested_stage4_product": "NONE", "required_stage4_tests": "NOT_RUN",
    "source_http_snapshot": "NOT_RETAINED; HTTP 429 capture failure preserved",
    "source_observations": "Already received successful ordinary HTTPS audit outputs; extracted evidence hashes verified",
    "stop_reason": result["stop_reason"], "safe_next_step": result["safe_next_step"]}
# Check current protected files once more immediately before publishing evidence seal.
assert all(digest(path) == expected for path, expected in before.items())
(OUT / "STAGE4_STOP_MANIFEST.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"status": "STOP", "integrity": manifest["evidence_integrity"],
                  "evidence_files": len(paths), "ready_for_independent_review": False}, ensure_ascii=False))
