"""B1 negative gates. Alter evidence on disposable copies, never product code."""
import json
from pathlib import Path
import subprocess
import sys
from unittest.mock import patch

from stage3_evidence import ROOT, product_hashes, timestamp, bind_result, file_hash


def run(out):
    before, started = product_hashes(), timestamp()
    manifest = ROOT / 'STAGE3_REVIEW_MANIFEST.json'
    original_manifest = file_hash(manifest)
    valid = json.loads((out / 'targeted_results.json').read_text(encoding='utf-8'))
    checks = []
    checkpoint = json.loads((ROOT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
    cases = [('legacy_unbound_PASS', Path(checkpoint['validation_root']))]
    for name, mode in [('stale_tested_hash', 'both'), ('before_after_mismatch', 'after'), ('missing_product_path', 'missing')]:
        folder = out / ('stage3_reject_' + name)
        folder.mkdir()
        altered = json.loads(json.dumps(valid))
        path = str(ROOT / 'passport_app/catalog.py')
        if mode == 'both':
            altered['product_sha256_before'][path] = altered['product_sha256_after'][path] = '0' * 64
        elif mode == 'after':
            altered['product_sha256_after'][path] = '0' * 64
        else:
            for field in ('product_sha256_before', 'product_sha256_after'):
                del altered[field][str(ROOT / 'stage3.py')]
        (folder / 'targeted_results.json').write_text(json.dumps(altered, ensure_ascii=False, indent=2), encoding='utf-8')
        cases.append((name, folder))
    for name, folder in cases:
        result = subprocess.run([sys.executable, str(ROOT / 'tests/stage3_finalize.py'), '--test-root', str(folder)],
                                capture_output=True, text=True, encoding='utf-8')
        assert result.returncode != 0, name + ': finalizer must refuse'
        assert 'Stage 3 finalization refused:' in result.stderr, result.stderr
        assert file_hash(manifest) == original_manifest, 'Successful manifest was changed by a rejection'
        checks.append({'scenario': name, 'status': 'PASS', 'exit_code': result.returncode,
                       'message': result.stderr.strip(), 'successful_manifest_unchanged': True})
    altered = dict(before)
    altered[str(ROOT / 'passport_app/catalog.py')] = '0' * 64
    with patch('stage3_evidence.product_hashes', return_value=altered):
        try:
            bind_result({'status': 'PASS'}, before, started)
        except ValueError as exc:
            checks.append({'scenario': 'suite_hash_changed_no_PASS', 'status': 'PASS', 'message': str(exc)})
        else:
            raise AssertionError('A suite must not bind PASS if the post-run hash differs')
    return bind_result({'status': 'PASS', 'checks': checks, 'pass_count': len(checks), 'fail_count': 0,
                        'manifest_sha256_before_after': original_manifest}, before, started)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--test-root', type=Path, required=True)
    out = parser.parse_args().test_root.resolve()
    result = run(out)
    (out / 'integrity_negative_results.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print('B1 negative evidence/finalizer gates: 5/5 PASS')
