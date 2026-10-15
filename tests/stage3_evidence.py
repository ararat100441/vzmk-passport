"""Small shared Stage 3 evidence gate: tested sources must equal current sources."""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PRODUCT_FILES = ('passport_app/catalog.py', 'passport_app/catalog_gui.py', 'stage3.py',
                 'Запустить Этап 3.ps1', 'Запустить Этап 3.cmd')


def timestamp():
    return datetime.now(timezone.utc).isoformat()


def file_hash(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def product_hashes():
    return {str(ROOT / name): file_hash(ROOT / name) for name in PRODUCT_FILES}


def require_same_code(before, after):
    if before != after:
        changed = [p for p in sorted(before.keys() | after.keys()) if before.get(p) != after.get(p)]
        raise ValueError('STOP: Stage 3 product code hash mismatch: ' + ', '.join(changed))


def evidence_root(checkpoint):
    out = Path(os.environ.get('PASPORT_STAGE3_TEST_ROOT', checkpoint['validation_root'])).resolve()
    if not out.is_relative_to(ROOT / 'validation') or not out.name.startswith('stage3_') or not out.is_dir():
        raise ValueError('Stage 3 evidence must use an existing disposable validation/stage3_* directory.')
    return out


def bind_result(result, before, started_at):
    after = product_hashes()
    require_same_code(before, after)
    return result | {'started_at_utc': started_at, 'completed_at_utc': timestamp(),
                     'product_sha256_before': before, 'product_sha256_after': after}


def verify_result(result, expected=None):
    current = product_hashes() if expected is None else expected
    before, after = result.get('product_sha256_before'), result.get('product_sha256_after')
    if not isinstance(before, dict) or not isinstance(after, dict):
        raise ValueError('STOP: test evidence is not bound to Stage 3 product SHA-256. Rerun targeted tests.')
    require_same_code(before, after)
    require_same_code(before, current)
    if result.get('status') != 'PASS' or not result.get('started_at_utc') or not result.get('completed_at_utc'):
        raise ValueError('STOP: missing completed PASS test evidence/timestamps.')
    if result.get('pass_count', 0) <= 0 or result.get('fail_count') != 0:
        raise ValueError('STOP: test evidence contains failures or missing PASS/FAIL counts.')
    return before
