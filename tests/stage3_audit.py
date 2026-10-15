"""Read-only Stage 3 start checkpoint. Never accepts an old registry digest."""
import hashlib
import json
import os
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from passport_app.config import Settings, PassportError
from passport_app.excel import locked_file, sha256
from passport_app.registry import read_registry


def protected():
    settings = Settings.load(ROOT / 'settings.json')
    files = [ROOT / 'REVIEW_MANIFEST.json', ROOT / 'STAGE2_REVIEW_MANIFEST.json']
    for folder, directories, names in os.walk(settings.documents_root):
        directories[:] = [d for d in directories if (Path(folder) / d).resolve() != ROOT]
        files += [Path(folder) / n for n in names
                  if Path(n).suffix.lower() in ('.xlsx', '.xlsm', '.docx', '.xls')]
    return {str(p): sha256(p) for p in files}


def main():
    settings = Settings.load(ROOT / 'settings.json')
    checks = {}
    mismatches = []
    for name in ('REVIEW_MANIFEST.json', 'STAGE2_REVIEW_MANIFEST.json'):
        manifest = json.loads((ROOT / name).read_text(encoding='utf-8-sig'))
        for value, expected in manifest['code_sha256'].items():
            path = ROOT / value
            actual = sha256(path)
            checks[str(path)] = actual
            if actual != expected:
                mismatches.append(str(path))
    if mismatches:
        raise PassportError('STOP: неожиданные изменения accepted files: ' + repr(mismatches))
    # Require the registry to be saved and closed, without touching Excel.
    with locked_file(settings.registry_path) as data:
        digest = hashlib.sha256(data).hexdigest()
    registry = read_registry(settings.registry_path)
    if registry.digest != digest:
        raise PassportError('STOP: production изменился при фиксации checkpoint.')
    out = ROOT / 'validation' / ('stage3_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
    out.mkdir(parents=True, exist_ok=False)
    evidence = {'stage': 3, 'created_at': datetime.now().astimezone().isoformat(),
                'production_path': str(settings.registry_path), 'production_sha256': digest,
                'accepted_files_sha256': checks, 'protected_sha256': protected(),
                'projects': list(registry.projects), 'project_count': len(registry.projects),
                'validation_root': str(out), 'baseline_policy': 'CURRENT_SAVED_CLOSED_REGISTRY'}
    if sha256(settings.registry_path) != digest:
        raise PassportError('STOP: production изменился во время аудита.')
    (ROOT / 'STAGE3_CHECKPOINT.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding='utf-8')
    (out / 'start_checkpoint.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'checkpoint': str(ROOT / 'STAGE3_CHECKPOINT.json'), 'validation_root': str(out),
                      'production_sha256': digest, 'accepted_files': len(checks),
                      'protected_files': len(evidence['protected_sha256']), 'projects': len(registry.projects)}, ensure_ascii=False))


if __name__ == '__main__':
    main()
