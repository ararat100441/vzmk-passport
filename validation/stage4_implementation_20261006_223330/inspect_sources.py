"""Read accepted evidence only; capture the protected baseline before implementation."""
import collections
import hashlib
import json
from pathlib import Path
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent
A = ROOT / 'validation/stage4_km_kmd_audit_20261006_204347'
B = ROOT / 'validation/stage4_ptoflow_audit_20261006_215239'
def read(p):
    return json.loads(p.read_text(encoding='utf-8-sig'))
def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()
protected = read(B / 'protected_hashes_after.json')['hashes']
for p in B.rglob('*'):
    if p.is_file():
        protected[str(p)] = sha(p)
bad = [p for p, h in protected.items() if sha(Path(p)) != h]
assert not bad, bad
(OUT / 'protected_before.json').write_text(json.dumps(protected, ensure_ascii=False, indent=2), encoding='utf-8')
a = read(A / 'parsed_source_records.json')
b = read(B / 'group_source_evidence.json')
print('Protected:', len(protected), 'Source A:', len(a), 'Source B:', len(b))
for source, rows in [('A', a), ('B', b)]:
    print(source, dict(collections.Counter(r['standard'] for r in rows)))
    for fam in sorted({r['family'] for r in rows}):
        row = next(r for r in rows if r['family'] == fam)
        print(source, fam, json.dumps({k: row[k] for k in ('raw', 'parsed')}, ensure_ascii=False))
print('Evidence manifests:', {str(x): list(read(x / 'AUDIT_MANIFEST.json')['evidence_files']) for x in (A, B)})
print('A corner source_profile examples:', sorted({str(r['effective']['source_profile']) for r in a if 'Уголок' in r['family']}))
print('B family count:', dict(collections.Counter(r['family'] for r in b)))
print('DEV checkpoints:', [(str(p.relative_to(ROOT)), list(read(p))) for p in (ROOT/'DEV').glob('*CHECKPOINT*.json')])
