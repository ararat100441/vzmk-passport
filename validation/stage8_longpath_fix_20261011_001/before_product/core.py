"""Preview -> explicit confirmation -> durable intent -> backup -> stage -> commit.

The journal is the closing register. Stages 1-7 and their registry schema do not
change. This release authorizes CLOSED_TEST only, in bootstrapped sandboxes.
"""
from pathlib import Path
from uuid import uuid4
from . import storage as S
from .model import prepare, key, number
from .native import stage, inspect


def code_bindings():
    root = Path(__file__).parent
    return {str(p.resolve()): S.sha(p) for p in [*sorted(root.glob('*.py')), root.parent/'stage8.py']}


def check_bindings(bindings):
    for p,h in bindings.items():
        S.no_links(p)
        if S.sha(p) != h:
            raise S.Stop('INPUT_OR_CODE_CHANGED: '+p)


def state(root):
    h = S.history(root)
    if h['active'] is None and S.sha(root/'registry/УЗК_TEST.xlsm') != h['expected_sha256']:
        raise S.Stop('REGISTRY_CHANGED_OR_JOURNAL_LOST')
    if h['active'] is None:
        identity = h['expected_identity']
        if identity is not None and S.file_identity(root/'registry/УЗК_TEST.xlsm') != identity:
            raise S.Stop('REGISTRY_FILE_ID_CHANGED')
    return h


def request_key(root, path):
    q = S.read(S.within(path,root/'inputs'))
    if q.get('schema') != 'STAGE8_CLOSE_REQUEST_V1' or q.get('mode') != 'ISOLATED_TEST' or q.get('official_issue') is not False:
        raise S.Stop('OFFICIAL_CLOSURE_FORBIDDEN')
    if q.get('passport') == '9/2026':
        raise S.Stop('PASSPORT_9_2026_DEV_TEST_FORBIDDEN')
    return S.digest({'project':key(q['project']), 'passport':number(q['passport'])})


def seal(payload):
    return {'schema':'STAGE8_PREVIEW_V1', 'payload':payload, 'seal':S.digest(payload)}


def preview(sandbox_root, request_path):
    root, _ = S.sandbox(sandbox_root)
    with S.lock(root), S.registry_guard(root/'registry/УЗК_TEST.xlsm') as guard:
        h = state(root)
        if h['active'] is not None:
            raise S.Stop('INCOMPLETE_OPERATION_RECOVER_FIRST: '+h['active']['operation_id'])
        identity = request_key(root,request_path)
        if identity in h['closed']:
            return seal({'status':'ALREADY_CLOSED_TEST', 'sandbox':str(root),
                         'key':identity, 'closed':h['closed'][identity]})
        payload = prepare(root,request_path) | {'status':'READY_TEST', 'sandbox':str(root),
                    'code':code_bindings(), 'registry_identity':guard.identity}
        if payload['before_sha256'] != h['expected_sha256']:
            raise S.Stop('REGISTRY_CHANGED')
        with S.frozen(list(payload['bindings']) + list(payload['code'])):
            check_bindings(payload['bindings'] | payload['code'])
            guard.check(payload['registry_identity'],payload['before_sha256'])
        return seal(payload)


def confirmation(plan):
    if plan['payload']['status'] == 'ALREADY_CLOSED_TEST':
        return ''
    return 'ЗАКРЫТЬ ТЕСТ '+plan['payload']['identity']['passport']+' '+plan['seal'][:12]


def close(plan, explicit_confirmation, *, _fault=None):
    if plan.get('schema') != 'STAGE8_PREVIEW_V1' or S.digest(plan['payload']) != plan.get('seal'):
        raise S.Stop('PREVIEW_TAMPERED')
    p = plan['payload']
    root, _ = S.sandbox(p['sandbox'])
    with S.lock(root), S.registry_guard(root/'registry/УЗК_TEST.xlsm') as guard:
        h = state(root)
        if h['active'] is not None:
            raise S.Stop('INCOMPLETE_OPERATION_RECOVER_FIRST')
        if p['key'] in h['closed']:
            return {'status':'ALREADY_CLOSED_TEST', 'official_issue':False, 'operation':h['closed'][p['key']]}
        if p['status'] != 'READY_TEST' or explicit_confirmation != confirmation(plan):
            raise S.Stop('EXPLICIT_CONFIRMATION_REQUIRED')
        guard.check(p['registry_identity'],p['before_sha256'])
        current = prepare(root,p['request_path']) | {'status':'READY_TEST','sandbox':str(root),
                    'code':code_bindings(),'registry_identity':guard.identity}
        if current != p:
            raise S.Stop('PREVIEW_STALE_REPEAT_CHECK')
        registry = Path(p['registry'])
        journal_files = list((root/'journal').glob('*.json'))
        with S.frozen(list(p['bindings']) + list(p['code']) + journal_files + [root/'SANDBOX.json']):
            check_bindings(p['bindings'] | p['code'])
            if S.sha(registry) != p['before_sha256']:
                raise S.Stop('REGISTRY_CHANGED_AFTER_PREVIEW')
            op = {'key':p['key'],'operation_id':uuid4().hex}
            folder = root/'transactions'/op['operation_id']
            folder.mkdir(exist_ok=False)
            intent = S.append(root,'INTENT',op,plan=p,preview_seal=plan['seal'],
                              before_sha256=p['before_sha256'],registry_mutation=bool(p['deductions']),
                              confirmed_by_user=True,closed_status='CLOSED_TEST',official_issue=False)
            if _fault:
                _fault('after_intent',registry)
            if not p['deductions']:
                # Same durable protocol; no XLSM backup/save/replacement at all.
                S.append(root,'BACKUP_VERIFIED',op,backup=None,reason='NO_UZK_REQUIRED_NO_REGISTRY_WRITE')
                S.append(root,'STAGED',op,after_sha256=p['before_sha256'],preservation=None)
                S.append(root,'APPLYING',op,registry_mutation=False)
                return _commit(root,op,False)
            reference = S.backup(root,op,registry,p['before_sha256'])
            if _fault:
                _fault('after_backup',Path(reference['path']))
            S.verify_backup(reference,root)
            S.append(root,'BACKUP_VERIFIED',op,backup=reference)
            output = folder/'stage.xlsm'
            preservation = stage(registry,output,p['deductions'],_fault)
            S.write(folder/'XLSM_PRESERVATION.json',preservation)
            after = S.sha(output)
            S.flush_file(output)
            if after == p['before_sha256']:
                raise S.Stop('DEDUCTION_DID_NOT_CHANGE_WORKBOOK')
            S.append(root,'STAGED',op,after_sha256=after,stage=str(output),
                     preservation={'path':str(folder/'XLSM_PRESERVATION.json'),'sha256':S.sha(folder/'XLSM_PRESERVATION.json')})
            if _fault:
                _fault('after_save',output)
            with S.frozen([output,reference['path'],reference['metadata']]):
                check_bindings(p['bindings'] | p['code'])
                S.verify_backup(reference,root)
                if S.sha(registry) != p['before_sha256'] or S.sha(output) != after:
                    raise S.Stop('PRE_PUBLISH_SHA256_MISMATCH')
                S.append(root,'APPLYING',op,registry_mutation=True)
                if _fault:
                    _fault('before_replace',output)
                guard.publish(output,p['registry_identity'],p['before_sha256'],after,_fault)
                if _fault:
                    _fault('after_replace',registry)
                guard.check(p['registry_identity'],after)
                return _commit(root,op,False)


def _commit(root, op, recovered):
    S.append(root,'COMMITTED',op,recovered=recovered,closed_status='CLOSED_TEST',official_issue=False)
    return {'status':'CLOSED_TEST','official_issue':False,'recovered':recovered,
            'operation':S.history(root)['closed'][op['key']]}


def recover(sandbox_root):
    """Never save or replace XLSM on recovery. Only finalize or abort the journal."""
    root, _ = S.sandbox(sandbox_root)
    with S.lock(root):
        pending = S.history(root)['active']
        if pending and pending['state'] in ('RESTORE_PREPARED','RESTORE_APPLYING'):
            raise S.Stop('OPERATOR_RESTORE_PENDING_EXPLICIT_CONFIRMATION_REQUIRED')
        return _recover_locked(root)


def _recover_locked(root):
    with S.registry_guard(root/'registry/УЗК_TEST.xlsm') as guard:
        h = state(root)
        op = h['active']
        if op is None:
            return {'status':'NO_PENDING_OPERATION','registry_writes':0}
        registry = root/'registry/УЗК_TEST.xlsm'
        guards = [root/'SANDBOX.json',*(root/'journal').glob('*.json')]
        if op.get('backup'):
            guards += [op['backup']['path'],op['backup']['metadata']]
        if op.get('preservation'):
            guards += [op['preservation']['path']]
        with S.frozen(guards):
            h = S.history(root)
            op = h['active']
            expected_identity = op['plan'].get('registry_identity')
            if expected_identity is None:
                raise S.Stop('LEGACY_PENDING_OPERATION_REQUIRES_REVIEW')
            if guard.identity != expected_identity:
                raise S.Stop('AMBIGUOUS_REGISTRY_FILE_ID_STOP_NO_REPLAY')
            if op.get('backup'):
                S.verify_backup(op['backup'],root)
            if op.get('preservation'):
                path = S.within(op['preservation']['path'],root/'transactions')
                if S.sha(path) != op['preservation']['sha256']:
                    raise S.Stop('PRESERVATION_EVIDENCE_DAMAGED')
            current = S.sha(registry)
            # An unchanged registry always permits a safe abort, never a retry.
            if current == op['before_sha256']:
                S.append(root,'ABORTED',op,reason='REGISTRY_UNCHANGED_NO_REPLAY',registry_writes=0)
                return {'status':'ABORTED_UNCHANGED','registry_writes':0,'operation_id':op['operation_id']}
            if (op['state'] == 'APPLYING' and op.get('registry_mutation') is True
                    and current == op.get('after_sha256') and op.get('backup') and op.get('preservation')):
                native = inspect(registry)
                return _commit(root,op,True) | {'registry_writes':0,'native':native}
            raise S.Stop('AMBIGUOUS_OPERATION_STOP_NO_REPLAY')


def describe(plan):
    p = plan['payload']
    if p['status'] == 'ALREADY_CLOSED_TEST':
        return 'Паспорт уже закрыт в TEST. Повторное списание запрещено.\nОперация: '+p['closed']['operation_id']
    lines = ['ТОЛЬКО ИЗОЛИРОВАННЫЙ TEST — официальная выдача запрещена.',
        'Проект: '+p['identity']['project'],'Паспорт: '+p['identity']['passport'],
        'Выбранная отгрузка: '+p['identity']['shipment_id'],
        'Статус проверки: PASS / подтверждённые искусственные TEST-данные.',
        'Обязательные документы: все в наличии.','Состав документов:']
    lines += ['  '+d['kind']+' — CONFIRMED_TEST — '+Path(d['path']).name for d in p['documents']]
    lines += ['Количество подтверждённых изделий УЗК: '+str(p['controlled_units']), 'Остатки УЗК до → после:']
    lines += [f"  {d['mark']}: {d['before']} − {d['quantity']} = {d['after']} шт." for d in p['deductions']]
    if not p['deductions']:
        lines.append('  УЗК не требуется. Реестр УЗК не изменяется.')
    lines += ['Закрытие TEST станет окончательным в журнале; повторное открытие не спишет остатки.']
    return '\n'.join(lines)


def inspect_stage7(path):
    """Read-only adapter; a Stage 7 TEST manifest never enters the write engine."""
    q = S.read(path)
    if q.get('schema') != 'STAGE7_SELECTED_TEST_MANIFEST_V1':
        raise S.Stop('STAGE7_MANIFEST_REQUIRED')
    membership = q.get('membership',{})
    return {'status':'BLOCKED_DEV_TEST','official_issue':False,'registry_writes':0,
        'project':q.get('project'),'passport':q.get('passport'),'shipment':membership.get('vom_path'),
        'documents':[{'kind':d.get('type'),'status':d.get('status'),'path':d.get('path')} for d in q.get('records',[])],
        'controlled_units':next((d.get('verification',{}).get('quantity') for d in q.get('records',[]) if d.get('type')=='UZK'),None),
        'reason':'DEV/TEST; SAMPLE_ONLY УЗК; фактический контроль не подтверждён. Официальное закрытие и списание запрещены.',
        'balances':'Списание запрещено; изменение остатков: 0.'}
