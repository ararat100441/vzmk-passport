"""Imports only; do not load settings, instantiate services or open registry in Excel."""
from hashlib import sha256
import importlib
import json
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.dont_write_bytecode = True
OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
sys.path.insert(0, str(ROOT))
before = json.loads((OUT / "accepted_hashes_before.json").read_text(encoding="utf-8"))
assert all(sha256(Path(p).read_bytes()).hexdigest() == h for p,h in before.items())
modules = ("passport_app.config", "passport_app.service", "passport_app.vik", "passport_app.catalog",
           "passport_app.catalog_gui", "main", "stage2", "stage3")
for name in modules:
    importlib.import_module(name)
assert all(sha256(Path(p).read_bytes()).hexdigest() == h for p,h in before.items())
result = {"scope": "IMPORT_ONLY_SMOKE", "status": "PASS", "modules": list(modules),
    "product_and_DEV_hashes_before_after_current_equal": True,
    "product_and_checkpoint_hashes": before, "workbook_open_or_Apply_or_generation": "NOT_RUN"}
(OUT / "smoke_results.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print("Stage 1–3 imports: 8/8 PASS; protected hashes unchanged")
