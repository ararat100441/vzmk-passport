"""Explicit operator rollback of a proved torn TEST registry, never a debit.

Two non-replacing handle renames publish a complete verified book. The journal
describes all three namespace states before moving anything. An absent target
is recoverable only with that prior intent and the exact two retained identities.
"""
from contextlib import contextmanager, ExitStack
import io
import os
from pathlib import Path
from uuid import uuid4
from zipfile import ZipFile, BadZipFile
from . import storage as S, native, core
from .shortcopy import workspace

RESTORE_STATES = ('RESTORE_PREPARED','RESTORE_APPLYING')


def _validate(root, op, h):
    registry = root/'registry/УЗК_TEST.xlsm'
    if (op['state'] not in ('APPLYING',*RESTORE_STATES) or not op.get('registry_mutation')
            or op['key'] in h['closed'] or h['expected_sha256'] != op['before_sha256']
            or not op['plan'].get('registry_identity') or op['plan'].get('registry') != str(registry)
            or op['plan']['identity']['passport'] == '9/2026'):
        raise S.Stop('RESTORE_NOT_A_PROVED_PENDING_REGISTRY_WRITE')
    b = op.get('backup',{})
    folder = root/'backups'/op['operation_id']
    if (b.get('path') != str(folder/'УЗК_TEST.xlsm')
            or b.get('metadata') != str(folder/'SHA256.json')
            or b.get('sha256') != op['before_sha256']):
        raise S.Stop('RESTORE_BACKUP_BEFORE_BINDING_MISMATCH')
    S.verify_backup(b,root)
    p = op.get('preservation',{})
    expected = root/'transactions'/op['operation_id']/'XLSM_PRESERVATION.json'
    if p.get('path') != str(expected) or S.sha(expected) != p.get('sha256'):
        raise S.Stop('PRESERVATION_EVIDENCE_DAMAGED')
    proof = S.read(expected)
    original_stage = root/'transactions'/op['operation_id']/'stage.xlsm'
    if (proof.get('status') != 'PASS' or proof.get('sha256') != op.get('after_sha256')
            or op.get('stage') != str(original_stage) or S.sha(original_stage) != op['after_sha256']):
        raise S.Stop('RESTORE_ORIGINAL_STAGE_EVIDENCE_DAMAGED')
    return registry


@contextmanager
def _scope(root):
    with S.lock(root), S.namespace_guard(root/'registry'):
        h = S.history(root)
        op = h['active']
        if op is None:
            raise S.Stop('NO_PENDING_OPERATION_RESTORE_FORBIDDEN')
        paths = [root/'SANDBOX.json',*(root/'journal').glob('*.json')]
        if op.get('backup'):paths += [op['backup']['path'],op['backup']['metadata']]
        if op.get('preservation'):paths += [op['preservation']['path']]
        if op.get('stage'):paths += [op['stage']]
        for path in paths:S.within(path,root)
        with S.frozen(paths):
            h = S.history(root);op = h['active']
            registry = _validate(root,op,h)
            yield h,op,registry


def _torn(guard):
    try:
        with ZipFile(io.BytesIO(guard.read_bytes())) as z:
            if z.testzip() is None:
                raise S.Stop('RESTORE_VALID_UNRECOGNIZED_BOOK_REQUIRES_REVIEW')
    except (BadZipFile,EOFError,RuntimeError) as exc:
        return str(exc)
    return 'XLSM_CRC_FAILURE'


def _locations(root,op,stack,*,rename=False):
    registry = root/'registry/УЗК_TEST.xlsm'
    guards = {}
    if op['state'] == 'APPLYING':
        g = stack.enter_context(S.registry_guard(registry,_rename=rename))
        if g.identity != op['plan']['registry_identity']:
            raise S.Stop('UNKNOWN_EXTERNAL_REGISTRY_IDENTITY_RESTORE_FORBIDDEN')
        current = g.sha()
        if current == op['before_sha256']:
            raise S.Stop('REGISTRY_UNCHANGED_USE_JOURNAL_RECOVER')
        if current == op.get('after_sha256'):
            raise S.Stop('FULL_RESULT_USE_JOURNAL_RECOVER_NO_ROLLBACK')
        _torn(g)
        return 'TORN',{'target':g},{'target':{'identity':g.identity,'sha256':current}}
    r = op['restoration']
    for name,path in [('target',registry),('stage',Path(r['stage'])),('quarantine',Path(r['quarantine']))]:
        S.within(path,root)
        if path.exists():guards[name] = stack.enter_context(S.registry_guard(path,_rename=rename))
    actual = {n:{'identity':g.identity,'sha256':g.sha()} for n,g in guards.items()}
    old = {'identity':r['old_identity'],'sha256':r['damaged_sha256']}
    new = {'identity':r['new_identity'],'sha256':op['before_sha256']}
    known = {'PREPARED':{'target':old,'stage':new},
             'VACANT':{'quarantine':old,'stage':new},
             'PUBLISHED':{'quarantine':old,'target':new}}
    for phase,expected in known.items():
        if actual == expected:return phase,guards,actual
    raise S.Stop('RESTORE_NAMESPACE_OR_CONTENT_AMBIGUOUS_NO_CHANGES')


def _payload(root,h,op,phase,actual):
    return {'status':'RESTORE_AVAILABLE_TEST','sandbox':str(root),'operation_id':op['operation_id'],
            'key':op['key'],'operation_state':op['state'],'phase':phase,
            'registry':str(root/'registry/УЗК_TEST.xlsm'),'backup':op['backup'],
            'verified_before_sha256':op['before_sha256'],'locations':actual,
            'journal_tip':S.sha(root/'journal'/f'{len(h["events"])-1:08d}.json'),
            'code':core.code_bindings(),'official_issue':False,
            'reason':'Повреждённая частичная запись XLSM; COMMITTED отсутствует.' if phase=='TORN'
                     else 'Завершить явно подтверждённое восстановление после сбоя.',
            'consequences':'Вернуть точные байты BEFORE. Исходная операция RECOVERED_ABORTED; '
                           'паспорт не CLOSED, списаний 0. Для закрытия нужен новый preview и подтверждение.'}


def preview(sandbox_root):
    root,_ = S.sandbox(sandbox_root)
    with _scope(root) as (h,op,registry), ExitStack() as stack, workspace() as verification:
        phase,guards,actual = _locations(root,op,stack)
        # Both format/structure and actual Excel normal read-only open are required.
        native.inspect(op['backup']['path'],verification=verification)
        payload = _payload(root,h,op,phase,actual)
        return {'schema':'STAGE8_RESTORE_PREVIEW_V1','payload':payload,'seal':S.digest(payload)}


def confirmation(plan):
    return 'ВОССТАНОВИТЬ ТЕСТ '+plan['payload']['operation_id']+' '+plan['seal'][:12]


def describe(plan):
    p = plan['payload']
    return '\n'.join(['Обнаружена незавершённая операция: '+p['operation_id'],
        'Статус: '+p['operation_state']+' / '+p['phase'],'Причина: '+p['reason'],
        'Проверенный backup: '+p['backup']['path'],'SHA-256 backup: '+p['verified_before_sha256'],
        'Затрагиваемый реестр: '+p['registry'],'Восстановление доступно в изолированном TEST.',p['consequences']])


def restore(plan,explicit_confirmation,*,_fault=None):
    if (plan.get('schema') != 'STAGE8_RESTORE_PREVIEW_V1'
            or plan.get('seal') != S.digest(plan['payload'])):
        raise S.Stop('RESTORE_PREVIEW_TAMPERED')
    if explicit_confirmation != confirmation(plan):
        raise S.Stop('RESTORE_EXPLICIT_CONFIRMATION_REQUIRED')
    p = plan['payload'];root,_ = S.sandbox(p['sandbox'])
    with _scope(root) as (h,op,registry), ExitStack() as stack, workspace() as verification:
        if p.get('code') != core.code_bindings():raise S.Stop('RESTORE_CODE_CHANGED_REPEAT_PREVIEW')
        stack.enter_context(S.frozen(p['code']))
        core.check_bindings(p['code'])
        phase,guards,actual = _locations(root,op,stack,rename=True)
        if _payload(root,h,op,phase,actual) != p:
            raise S.Stop('RESTORE_PREVIEW_STALE_REPEAT_CHECK')
        # Test the exact mandatory native route before any target rename,
        # including resumed PREPARED/VACANT/PUBLISHED operations.
        native.inspect(op['backup']['path'],verification=verification)
        def fault(point,target):
            if _fault:_fault(point,target)
        if phase == 'TORN':
            rid = uuid4().hex
            folder = root/'transactions'/op['operation_id']/('restore_'+rid)
            folder.mkdir(exist_ok=False)
            stack.enter_context(S.namespace_guard(folder))
            output = folder/'verified.xlsm'
            # Stage privately and flush before any target namespace mutation.
            with Path(op['backup']['path']).open('rb') as source, output.open('xb') as destination:
                block = source.read(4096)
                while block:
                    destination.write(block)
                    destination.flush()
                    fault('restore_copy_chunk',output)
                    block = source.read(1024*1024)
                os.fsync(destination.fileno())
            fault('restore_stage_copy',output)
            if S.sha(output) != op['before_sha256']:raise S.Stop('RESTORE_STAGE_SHA_MISMATCH')
            native.inspect(output,diagnostics=folder/'office',verification=verification)
            new = stack.enter_context(S.registry_guard(output,_rename=True))
            new.check(new.identity,op['before_sha256'])
            guards['stage'] = new
            r = {'recovery_id':rid,'backup':op['backup'],'registry':str(registry),
                 'old_identity':guards['target'].identity,'new_identity':new.identity,
                 'damaged_sha256':actual['target']['sha256'],'verified_restored_sha256':op['before_sha256'],
                 'stage':str(output),'quarantine':str(folder/'damaged.xlsm')}
            S.append(root,'RESTORE_PREPARED',op,restoration=r,confirmed_by_user=True,
                     restore_preview_seal=plan['seal'],decrements=0)
            op = S.history(root)['active'];phase = 'PREPARED'
            fault('after_restore_intent',registry)
        r = op['restoration'];folder = Path(r['stage']).parent
        if op['state'] == 'RESTORE_PREPARED':
            S.append(root,'RESTORE_APPLYING',op,decrements=0)
        if phase == 'PREPARED':
            guards['target'].check(r['old_identity'],r['damaged_sha256'])
            guards['stage'].check(r['new_identity'],op['before_sha256'])
            fault('before_restore_quarantine',registry)
            guards['target'].rename_new(r['quarantine'])
            guards['quarantine'] = guards.pop('target')
            phase = 'VACANT'
            fault('after_restore_quarantine',registry)
        if phase == 'VACANT':
            guards['stage'].check(r['new_identity'],op['before_sha256'])
            # No REPLACE flag: a competing creator in the name gap is preserved.
            guards['stage'].rename_new(registry)
            guards['target'] = guards.pop('stage')
            fault('after_restore_publish',registry)
        target = guards['target']
        target.check(r['new_identity'],op['before_sha256'])
        guards['quarantine'].check(r['old_identity'],r['damaged_sha256'])
        # Read via the retained DELETE handle; never release identity protection
        # merely to let Excel open the source. Evidence names the actual target.
        result = native.inspect(registry,diagnostics=folder/'office',guard=target,verification=verification)
        if result['sha256'] != op['backup']['sha256']:raise S.Stop('RESTORE_NATIVE_SHA_MISMATCH')
        proof = folder/('NATIVE_RESTORED_'+uuid4().hex+'.json')
        S.write(proof,{'native':result,'registry':str(registry),'backup':op['backup'],
                'old_identity':r['old_identity'],'new_identity':r['new_identity'],
                'verified_restored_sha256':op['before_sha256'],'recovery_id':r['recovery_id']})
        target.check(r['new_identity'],op['before_sha256'])
        S.verify_backup(op['backup'],root)
        fault('before_restore_terminal',registry)
        S.append(root,'RECOVERED_ABORTED',op,recovery_id=r['recovery_id'],closed_status=None,
                 decrements=0,restored_identity=r['new_identity'],verified_restored_sha256=op['before_sha256'],
                 native_evidence={'path':str(proof),'sha256':S.sha(proof)},official_issue=False)
        return {'status':'RECOVERED_ABORTED','operation_id':op['operation_id'],
                'recovery_id':r['recovery_id'],'closed':False,'decrements':0,'official_issue':False,
                'old_identity':r['old_identity'],'new_identity':r['new_identity'],
                'verified_restored_sha256':op['before_sha256'],'native_evidence':str(proof)}
