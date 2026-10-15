"""Repeat only the corrected reviewer alias assertion; bind all retained results."""
from pathlib import Path
from copy import deepcopy
from datetime import datetime, timezone
import ast, hashlib, json, os, sys, traceback

OUT = Path(__file__).resolve().parent
ROOT = OUT.parents[1]
sys.dont_write_bytecode = True
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding='utf-8')

def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def save(name, value):
    (OUT / name).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')

def require(condition, message='assertion failed'):
    if not condition:
        raise AssertionError(message)

network = []
def audit(event, args):
    if event in ['socket.connect', 'socket.getaddrinfo', 'socket.sendto']:
        network.append(event)
        raise AssertionError('Network prohibited')
    if event == 'open' and isinstance(args[0], (str, bytes, os.PathLike)):
        path = Path(os.fsdecode(args[0])).resolve()
        require(path.name != 'ВЗМК.xlsx', 'Production registry prohibited')
        mode, flags = args[1:3]
        writes = (isinstance(mode, str) and any(c in mode for c in 'wax+')) or flags & (os.O_WRONLY | os.O_RDWR | os.O_CREAT | os.O_TRUNC)
        if path.is_relative_to(ROOT):
            require(not writes or path.is_relative_to(OUT), 'Review writes must remain isolated')
sys.addaudithook(audit)

from profiles_stage4 import storage as s
from profiles_stage4.model import Catalog, normalize_profile, BENT_RULE, CLOSED_RULE
CAT = s.load_catalog()
STATE = CAT.user_state
before = read(OUT / 'preservation_before.json')
after = read(OUT / 'preservation_after.json')
current = {p: sha(p) for p in before}
require(before == after == current, 'STOP: current artifacts or protected evidence changed')

initial = read(OUT / 'initial_harness_results.json')
require(initial['passed'] == 11 and initial['failed'] == 1)
test_name = 'alias variants raw preservation type and collision safety'
require([t['name'] for t in initial['tests'] if t['status'] == 'FAIL'] == [test_name])
require("KeyError: 'profile'" in next(t['traceback'] for t in initial['tests'] if t['status'] == 'FAIL'))

tree = ast.parse((OUT / 'review.py').read_text(encoding='utf-8'))
function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == 'aliases')
exec(compile(ast.Module(body=[function], type_ignores=[]), str(OUT / 'review.py'), 'exec'), globals())
try:
    alias_test = {'name': test_name, 'status': 'PASS', 'detail': aliases()}
except Exception:
    alias_test = {'name': test_name, 'status': 'FAIL', 'traceback': traceback.format_exc()}
save('corrected_alias_result.json', alias_test)

tests = [alias_test if t['name'] == test_name else t for t in initial['tests']]
passed = sum(t['status'] == 'PASS' for t in tests)
failed = len(tests) - passed
require(current == {p: sha(p) for p in current} and not network)
save('preservation_current_final.json', current)
save('results.json', {'tests': tests, 'passed': passed, 'failed': failed})

result = read(OUT / 'INDEPENDENT_REVIEW.json')
result.update(verdict='PASS' if not failed else 'FAIL', passed=passed, failed=failed, tests=tests,
              ready_for_targeted_user_E2E=not failed,
              harness_correction={
                  'classification': 'REVIEW_HARNESS_ONLY',
                  'initial_failure_retained': 'initial_harness_results.json',
                  'initial_script_retained': 'review.initial.py',
                  'problem': 'Reviewer assumed optional closed profile 140x60x4 exists; actual safe result is NOT_FOUND.',
                  'corrected_assertion': 'Closed syntax may be unavailable, but must never resolve to the open bent channel.',
                  'rerun_scope': 'Complete alias function only; other 11 PASS results retain exact before/after/current binding.',
                  'product_changes': 0,
              },
              findings_by_severity={'BLOCKER': 0, 'MAJOR': 0, 'MINOR': 0} if not failed else None,
              required_fix='NONE' if not failed else 'REVIEW_FAILURE_NEEDS_TRIAGE',
              protection={'files': len(current), 'protected': 1178, 'before_after_current_equal': True,
                          'real_DEV_writes': 0, 'production_registry_access': 0, 'HTTP_requests': 0})
save('INDEPENDENT_REVIEW.json', result)

# Manifest is deliberately the last evidence write.
evidence = {str(p.relative_to(OUT)): sha(p) for p in OUT.rglob('*')
            if p.is_file() and p.name != 'REVIEW_MANIFEST.json'}
save('REVIEW_MANIFEST.json', {
    'created_at': datetime.now(timezone.utc).isoformat(),
    'verdict': result['verdict'], 'passed': passed, 'failed': failed,
    'current_product_hashes': result['product_hashes'], 'current_DEV_artifacts': result['DEV_artifacts'],
    'preserved_files': len(current), 'BEFORE_AFTER_CURRENT_equal': True,
    'evidence_files': evidence,
})
print(json.dumps({k: result[k] for k in ['verdict', 'passed', 'failed', 'ready_for_targeted_user_E2E']}, ensure_ascii=False))
if failed:
    print(alias_test['traceback'])
    raise SystemExit(1)
