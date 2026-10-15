"""Exclusive operation records, immutable input bindings and fail-closed receipts."""
from contextlib import contextmanager
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
from uuid import uuid4


class Stage7Error(ValueError):
    pass


def now():
    return datetime.now(timezone.utc).isoformat()


def allowed_source(path):
    path = Path(path).resolve()
    if path.name.casefold() in ('взмк.xlsx', 'узк.xlsm'):
        raise Stage7Error('PRODUCTION_ACCESS_FORBIDDEN: ' + str(path))
    return path


def sha(path):
    return hashlib.sha256(allowed_source(path).read_bytes()).hexdigest()


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True).encode('utf8')).hexdigest()


def read(path):
    return json.loads(allowed_source(path).read_text(encoding='utf-8-sig'))


def write(path, value):
    path = Path(path)
    with path.open('x', encoding='utf8') as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2, default=str)
    return {'path': str(path.resolve()), 'sha256': sha(path)}


def verify(bindings):
    for path, expected in bindings.items():
        if not allowed_source(path).is_file() or sha(path) != expected:
            raise Stage7Error('INPUT_OR_CODE_CHANGED: ' + str(path))


@contextmanager
def frozen(bindings):
    """Windows handles permit reads and deny writes/deletes, including Office saves.

    Accepted Stage 2 locks its own private copies separately. These handles belong
    only to the original inputs/code; no incompatible second lock on those copies.
    """
    import win32con
    import win32file
    handles = []
    try:
        verify(bindings)
        for path in sorted(bindings):
            handles.append(win32file.CreateFile(str(allowed_source(path)), win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING, 0, None))
        verify(bindings)
        yield
        verify(bindings)
    finally:
        for handle in reversed(handles):
            handle.Close()


def output_boundary(path, root):
    path, root = Path(path).resolve(), Path(root).resolve()
    if not path.is_relative_to(root) or path == root:
        raise Stage7Error('OUTPUT_OUTSIDE_OPERATION')
    if path.exists():
        raise Stage7Error('OUTPUT_EXISTS: ' + str(path))
    return path


class Recorder:
    """There is deliberately no public 'make successful receipt' operation.

    run invokes the named registered callable, then validates the fresh output.
    File hashes are an auditable local chain, not a digital signature or a claim
    that arbitrary manually edited JSON is authentic.
    """
    def __init__(self, root, assembly_id, inputs, code):
        self.root = Path(root).resolve()
        self.assembly_id = assembly_id
        self.inputs, self.code = dict(inputs), dict(code)
        self.calls = {}

    def register(self, name, function):
        if name != function.__module__ + '.' + function.__name__:
            raise Stage7Error('GENERATOR_IDENTITY_MISMATCH')
        self.calls[name] = function

    def run(self, name, status, expected_output, arguments, validator):
        if name not in self.calls or status not in ('TEST', 'SAMPLE_ONLY'):
            raise Stage7Error('UNREGISTERED_GENERATOR_OR_OFFICIAL')
        output = output_boundary(expected_output, self.root)
        operation_id = uuid4().hex
        operation = self.root / 'operations' / operation_id
        operation.mkdir(parents=True, exist_ok=False)
        common = {'assembly_run_id': self.assembly_id, 'operation_id': operation_id,
                  'project': 'Гараж Казахстан', 'passport': '9/2026',
                  'generator': name, 'status': status}
        try:
            with frozen(self.inputs | self.code):
                start = write(operation / 'START.json', common | {'event': 'START', 'time': now(),
                    'inputs': self.inputs, 'code': self.code, 'arguments': arguments,
                    'expected_output': str(output)})
                returned = self.calls[name](**arguments)
                actual = returned.get('output') if isinstance(returned, dict) else returned
                if actual is None or Path(actual).resolve() != output or not output.is_file():
                    raise Stage7Error('ACTUAL_GENERATOR_OUTPUT_MISMATCH')
                result = write(operation / 'RESULT.json', common | {'event': 'RESULT', 'time': now(),
                    'start': start, 'output': str(output), 'sha256': sha(output), 'returned': returned})
                checks = validator(output, operation)
                if checks.get('status') != 'PASS':
                    raise Stage7Error('OUTPUT_VALIDATION_FAILED')
                verify(self.inputs | self.code)
                if sha(output) != read(result['path'])['sha256']:
                    raise Stage7Error('OUTPUT_CHANGED_DURING_VALIDATION')
                validation = write(operation / 'VALIDATION.json', common | {'event': 'VALIDATION',
                    'time': now(), 'start': start, 'result': result, 'checks': checks})
                receipt = write(operation / 'STAGE7_OPERATION_RECEIPT.json', common | {
                    'schema': 'STAGE7_OPERATION_RECEIPT_V1', 'time': now(),
                    'state': 'COMPLETE_VALIDATED_TEST', 'results_confirmed': False,
                    'official_issue': False, 'inputs': self.inputs, 'code': self.code,
                    'output': str(output), 'output_sha256': sha(output),
                    'start': start, 'result': result, 'validation': validation})
                return receipt
        except Exception as exc:
            write(operation / 'FAILURE.json', common | {'time': now(),
                'state': 'FAILED_VALIDATION_DO_NOT_USE', 'error': str(exc),
                'output_if_any': str(output) if output.exists() else None})
            raise


def verify_receipt(reference, assembly_id, registry_sha, output, output_sha, status):
    verify({reference['path']: reference['sha256']})
    receipt = read(reference['path'])
    if (receipt.get('schema') != 'STAGE7_OPERATION_RECEIPT_V1'
            or receipt.get('state') != 'COMPLETE_VALIDATED_TEST'
            or receipt.get('assembly_run_id') != assembly_id
            or receipt.get('status') != status
            or receipt.get('results_confirmed') is not False
            or receipt.get('official_issue') is not False
            or receipt.get('output') != output or receipt.get('output_sha256') != output_sha):
        raise Stage7Error('RECEIPT_IDENTITY_MISMATCH')
    from .sources import REGISTRY
    if receipt['inputs'].get(str(REGISTRY)) != registry_sha:
        raise Stage7Error('RECEIPT_REGISTRY_MISMATCH')
    verify(receipt['inputs'] | {output: output_sha})
    from .presentation_revision import verify_producer_code
    verify_producer_code(receipt['code'])
    events = []
    for key, event in [('start', 'START'), ('result', 'RESULT'), ('validation', 'VALIDATION')]:
        ref = receipt[key]
        verify({ref['path']: ref['sha256']})
        record = read(ref['path'])
        if any(record.get(k) != receipt[k] for k in ('assembly_run_id', 'operation_id', 'generator', 'status')) or record['event'] != event:
            raise Stage7Error('RECEIPT_CHAIN_MISMATCH')
        events.append(record)
    start, result, validation = events
    if (start['expected_output'] != output or result['output'] != output
            or result['sha256'] != output_sha or result['start'] != receipt['start']
            or validation['start'] != receipt['start'] or validation['result'] != receipt['result']
            or validation['checks']['status'] != 'PASS'
            or start['inputs'] != receipt['inputs'] or start['code'] != receipt['code']
            or not (start['time'] <= result['time'] <= validation['time'] <= receipt['time'])):
        raise Stage7Error('RECEIPT_CHAIN_MISMATCH')
    return receipt
