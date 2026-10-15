"""Exact version proof for the user-authorized Stage 7 presentation correction.

Historical producer receipts retain their original code hashes. Only these three
known Stage 7 transitions are accepted, with the original bytes archived BEFORE
editing. Inputs, Stage 1–6 code, outputs and receipt chains remain strict current
hash checks. A new quality receipt binds the current generator separately.
"""
from pathlib import Path
from .provenance import Stage7Error, sha, verify

ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / 'validation/stage7_user_e2e_fix_20261009_212419/code_before'
VERSIONS = {
    'word.py': ('689643d7dd3285c7f8666587991527f2ec3847b57eff2f09b4d35e6972edf098',
                'ddf3b8d70838a07c16bd9d3e4a59cd42a658f18eb7037121e569b3a092b13931'),
    'core.py': ('5320aba78b62d8732d70f3a4e5b7fa5128a17e4b370577c7796563686ac873a3',
                '34ad56913f47a6562caaecc85ee0c0a33824f8daf3c60d9841d78584b55be559'),
    'provenance.py': ('8aafcd90be5e2152a0b0824900077e5c37b9e32d33ea763b5ffde9da38a1e567',
                      '9a2b3ed8414ff4d32e8f999c4a848affebe2ce6492989873d83d18b6f98ba261'),
}


def historical_code_evidence(bindings):
    evidence = {}
    for name, (old, current) in VERSIONS.items():
        path = ROOT / 'stage7_quality' / name
        expected = bindings.get(str(path))
        if expected == old:
            if sha(path) != current:
                raise Stage7Error('UNAPPROVED_STAGE7_CODE_REVISION: ' + str(path))
            evidence[str(ARCHIVE / name)] = old
    verify(evidence)
    return evidence


def verify_producer_code(bindings):
    evidence = historical_code_evidence(bindings)
    for path, expected in bindings.items():
        key = Path(path)
        transition = VERSIONS.get(key.name) if key.parent == ROOT / 'stage7_quality' else None
        if transition and expected == transition[0]:
            continue  # Both exact current and original snapshot were checked above.
        verify({path: expected})
    return evidence
