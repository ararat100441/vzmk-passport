"""Freeze review artifacts; REVIEW_MANIFEST.json is the last written file."""
from pathlib import Path
import hashlib
import json
import difflib
from datetime import datetime, timezone

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
MANIFEST = RUN / 'REVIEW_MANIFEST.json'
assert not MANIFEST.exists(), 'manifest already frozen'


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(name):
    return json.loads((RUN / name).read_text(encoding='utf8'))


def save(name, value):
    (RUN / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf8')


baseline, migration, completed = read('baseline.json'), read('real_migration_result.json'), read('real_migration_completion.json')
before, after = read('product_hashes_BEFORE_TESTS.json'), read('product_hashes_AFTER_TESTS.json')
current = {p: sha(p) for p in before}
assert before == after == current
assert completed['status'] == 'PASS'
assert all(t['status'] == 'PASS' for t in read('isolated_tests.json'))
assert len(read('isolated_tests.json')) == 32
assert sha(migration['path']) == migration['new_sha256']
assert sha(migration['backup']) == migration['backup_sha256'] == baseline['source_sha256']
protected = read('protected_before.json')
assert protected == read('protected_after.json') == {p: sha(p) for p in protected}
save('product_hashes_CURRENT.json', current)
comparison = {p: {'ORIGINAL': baseline['product_original'].get(p, 'ABSENT/new'),
                  'BEFORE_TESTS': before[p], 'AFTER_TESTS': after[p], 'CURRENT': current[p]}
              for p in current}
save('product_hash_comparison.json', comparison)
changed = [p for p, initial in baseline['product_original'].items() if current[p] != initial]
assert set(changed) == {str(ROOT / p) for p in
                      ('passport_app/catalog.py', 'passport_app/catalog_gui.py', 'uzk_stage5/inputs.py')}
diffs = []
for relative in ('passport_app/catalog.py', 'passport_app/catalog_gui.py', 'uzk_stage5/inputs.py'):
    original = RUN / 'before' / relative
    assert sha(original) == baseline['product_original'][str(ROOT / relative)]
    diffs.extend(difflib.unified_diff(original.read_text(encoding='utf8').splitlines(True),
                 (ROOT / relative).read_text(encoding='utf8').splitlines(True),
                 fromfile='before/' + relative, tofile='after/' + relative))
for relative in ('passport_app/catalog_migration.py', 'CATALOG_SCHEMA_STAGE6_CONTRACT.txt'):
    diffs.extend(difflib.unified_diff([], (ROOT / relative).read_text(encoding='utf8').splitlines(True),
                                    fromfile='/dev/null', tofile='after/' + relative))
(RUN / 'product_changes.patch').write_text(''.join(diffs), encoding='utf8')
review_snapshot = RUN / 'review_product'
for relative in ('passport_app/catalog.py', 'passport_app/catalog_gui.py', 'passport_app/catalog_migration.py',
                 'uzk_stage5/inputs.py', 'CATALOG_SCHEMA_STAGE6_CONTRACT.txt'):
    destination = review_snapshot / relative
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes((ROOT / relative).read_bytes())
save('external_artifacts.json', {'real_catalog': {'path': migration['path'], 'sha256': migration['new_sha256']},
     'backup': {'path': migration['backup'], 'sha256': migration['backup_sha256']}})
expected = read('stage5_real_catalog_preview.json')
calculations = '\n'.join(f"  {m['mark']}: Q={m['Q']}; " + '; '.join(f"{s['L']}×{s['S']}" for s in m['segments'])
                         for m in expected['marks'])
report = f'''1. ACTUAL STATE
Project: Гараж Казахстан. Current real catalog: {migration['path']}
Before: 195 unique marks / 196 worksheet rows. After: 195 / 196.
The current-task source SHA was accepted; no historic audit SHA requirement.
Source SHA: {baseline['source_sha256']}

2. OLD CATALOG ASSESSMENT
Legacy A Марка | B Профиль | C Краска | D № АКЗ. All rows read.
Nonempty paint: 0. AKZ blank: 195; real assigned zero: 0.
Numeric zero and populated 1/2 are covered by isolated typed fixtures.
Native lock acquisition passed before assessment and again during migration.

3. SCHEMA MIGRATION
Exactly A Марка | B Профиль | C № АКЗ. One atomic real replacement.
Guarded current SHA; original bytes backed up before temporary XLSX creation.
Native normal-load Excel validation completed BEFORE ReplaceFileW.
195 marks/profiles and all AKZ values/types/styles/formats preserved.
Only worksheet XML and its workbook FilterDatabase column reference changed.
All other ZIP parts and the literal A/B cell XML are identical.

4. CHANGED FILES
passport_app/catalog.py: new 3-column creation, 3/legacy4 loading and width-aware append.
passport_app/catalog_gui.py: two messages now describe preservation of № АКЗ.
uzk_stage5/inputs.py: accepts both catalog headers, still reads only mark/profile.
ADDED passport_app/catalog_migration.py: explicit migration/native gate/guarded rollback.
ADDED CATALOG_SCHEMA_STAGE6_CONTRACT.txt: future contract only, no Stage6 implementation.
Real project catalog migrated; no other project migration.
Review patch and exact before source copies are included; all recovered before bytes
match the initial product SHA (including reconstructed inputs.py).

5. STAGE 3 IMPORT COMPATIBILITY
NEW / FILL / SAME / CONFLICT / CHECK_REQUIRED / ERROR: PASS.
Actual CatalogService Apply and actual Tk widget Check/Apply: PASS on isolated files.
TXT/TSV/Markdown and source line provenance preserved.
Manual 0/1/2 survive, new assignments blank; stale manual edits reject Apply.
Legacy Apply retains both paint and AKZ without migrating that fixture implicitly.

6. STAGE 5 COMPATIBILITY
New3 and legacy4 mapping: equal, 195 records.
Full actual preview against isolated migrated catalog: PASS.
Full actual preview against real migrated catalog: PASS, 4 marks / 33 items.
{calculations}
Word renderer unchanged; generate() never called; no new УЗК.docx.

7. MIGRATION VALIDATION
32/32 final isolated tests PASS before the real migration.
Nonempty paint (text, 0, formula, whitespace and multiple values): STOP.
Invalid AKZ (text, negative, fraction, bool, formula, error): STOP without conversion.
Stale SHA/native lock/native validation failure/atomic failure: original preserved.
Exact-byte guarded rollback tested only on an isolated fixture.

8. NATIVE EXCEL RESULT
Microsoft Excel 16.0, CorruptLoad=0 (normal load), ReadOnly=True,
UpdateLinks=0, macros/events disabled, no Save. All 195 rows checked.
Temporary real migration book BEFORE replace: PASS. Real resulting book: PASS.
Successful normal-load opening establishes no recovery was needed/attempted.
Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open

9. NEW CATALOG SHA
{migration['new_sha256']}

10. BACKUP / ROLLBACK
Backup: {migration['backup']}
Backup SHA: {migration['backup_sha256']} (exact original bytes).
Real rollback NOT_RUN; no restoration needed. Explicit rollback_catalog requires
the unchanged current SHA and immutable backup SHA, native checks then atomically
restores original bytes. STOP on intervening user work; never blindly overwrite it.

11. EVIDENCE INTEGRITY
baseline.json records original hashes; product_hashes_BEFORE_TESTS / AFTER_TESTS /
CURRENT are identical tested bytes. product_hash_comparison.json exposes original
versus final changes. 115 protected file hashes unchanged. Review product copies,
patch, row assessments, test runner, native results, previews and audits included.
REVIEW_MANIFEST.json is written last and hashes all evidence/support files.
Manifest itself is verified read-only after writing; its SHA is reported to the user.

12. REGRESSION IMPACT
Only three existing product source files changed, plus two new files above.
DEV ВЗМК_DEV.xlsx / УЗК_DEV.xlsm, Stage4 Профили.xlsx and sidecars, real VOM,
old acts/templates, Stage5 core and Word renderer unchanged.
Production registry access attempts: 0 in instrumented runs. Explicit native file
paths recorded; native test locks/read operations use DEV/fixture/catalog paths.
Historical E2E runtime and old review manifest were not updated by this scope.

13. NOT_RUN
Full manual USER E2E/clipboard/physical keyboard interaction, full unrelated-stage
regression, new Word generation, real rollback, other-project migration, production
registry operations and Stage6 generation/implementation: NOT_RUN.

14. READY FOR TARGETED INDEPENDENT REVIEW — YES
Readiness only. Independent review has not been performed in this task.
Stop after this target; Stage6 is not started.
'''
(RUN / 'REVIEW_REPORT.txt').write_text(report, encoding='utf8')
# LAST WRITE: no artifact is changed after this point.
manifest = {'kind': 'STAGE3_TARGETED_SCHEMA_REVIEW', 'created_utc': datetime.now(timezone.utc).isoformat(),
            'ready_for_targeted_independent_review': True, 'isolated_tests': '32/32 PASS',
            'real_migration_count': 1, 'product_current': current, 'protected_files': protected,
            'external_artifacts': read('external_artifacts.json'),
            'evidence': {str(p.relative_to(RUN)): sha(p) for p in sorted(RUN.rglob('*')) if p.is_file()}}
save(MANIFEST.name, manifest)
assert all(sha(RUN / p) == expected for p, expected in manifest['evidence'].items())
assert all(sha(p) == expected for p, expected in manifest['product_current'].items())
assert all(sha(p) == expected for p, expected in protected.items())
assert sha(migration['path']) == migration['new_sha256']
assert sha(migration['backup']) == migration['backup_sha256']
print(json.dumps({'manifest': str(MANIFEST), 'sha256': sha(MANIFEST),
                  'evidence_files': len(manifest['evidence']), 'protected_files': len(protected),
                  'status': 'PASS; last write was manifest'}, ensure_ascii=True))
