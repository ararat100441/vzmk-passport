"""Explicitly approved three-file compensation; no GUI or document changes.

The preserved CURRENT set, not the old operation's ORIGINAL, is the rollback
target on failure. A durable intent blocks issuing until completion/recovery.
"""
from pathlib import Path
from contextlib import ExitStack
from copy import deepcopy
from datetime import datetime,timezone
import hashlib,json,os,shutil
from uuid import uuid4
from .storage import Stop,safe,sha,read,save,digest,mutex,post_commit_path,windows_file_identity
from .coordinated import _lock_all,_write_bytes
from passport_app.excel import atomic_replace,locked_file

SCHEMA='VZMK_COMPENSATION_1'

def _sealed(path):
    value=read(safe(path));seal=value.pop('seal',None)
    if seal!=digest(value):raise Stop('Повреждён манифест компенсации.')
    return value

def _write_new_json(path,value):
    payload=json.dumps(value|{'seal':digest(value)},ensure_ascii=False,indent=2).encode('utf8')
    _write_bytes(path,payload)

def _restore_journal(path,payload):
    temp=Path(path).with_name('.compensation-journal-'+uuid4().hex+'.tmp')
    _write_bytes(temp,payload);os.replace(temp,path)

def _event(stage,event,**fields):
    with (stage/'audit.jsonl').open('ab') as stream:
        stream.write((json.dumps({'at':datetime.now(timezone.utc).isoformat(),'event':event,**fields},ensure_ascii=False)+'\n').encode('utf8'))
        stream.flush();os.fsync(stream.fileno())

def _documents(state,folder):
    folder=safe(folder);expected=state['output_hashes']|{'ISSUE.json':state['issue_sha256']}
    if {p.name for p in folder.iterdir()}!=set(expected):raise Stop('Изменён состав комплекта компенсации.')
    for name,h in expected.items():
        if Path(name).name!=name or sha(safe(folder/name))!=h:raise Stop('Изменён документ компенсации: '+name)
    return expected

def validate_closed(journal,state):
    """Validate immutable evidence, never pin live files after future issues."""
    journal=safe(journal);stage=journal.parent/'compensation'
    state=dict(state);state.pop('seal',None)
    info=state.get('compensation',{})
    receipt=stage/'receipt.json';original=stage/'journal.original.json';manifest=stage/'manifest.json'
    if (state.get('phase')!='COMPENSATED' or info.get('schema')!=SCHEMA
            or info.get('receipt')!=str(receipt) or sha(safe(receipt))!=info.get('receipt_sha256')
            or sha(safe(original))!=info.get('original_journal_sha256')
            or sha(safe(manifest))!=info.get('manifest_sha256')):raise Stop('Нет проверенного закрытия компенсации.')
    old=_sealed(original);r=_sealed(receipt);m=_sealed(manifest)
    base=deepcopy(state);base.pop('compensation',None);base['phase']='NEEDS_READBACK'
    if (old!=base or old['phase']!='NEEDS_READBACK' or r.get('schema')!=SCHEMA
            or r.get('id')!=state['id'] or m.get('id')!=state['id']
            or r.get('original_journal_sha256')!=info['original_journal_sha256']
            or r.get('manifest_sha256')!=info['manifest_sha256']
            or r.get('approval',{}).get('approved') is not True
            or r['approval'].get('manifest_sha256')!=info['manifest_sha256']):raise Stop('Неверная привязка закрытия компенсации.')
    expected={i['role']:i['path'] for i in state['participants']};approved={i['role']:i for i in m['participants']}
    if (journal!=Path(state['registry']).parent/'.vzmk-operations'/state['id']/'journal.json'
            or sha(safe(state['plan']))!=m['plan_sha256']):raise Stop('Повреждена привязка плана компенсации.')
    if len(r.get('participants',[]))!=3 or {i['role'] for i in r['participants']}!=set(expected):raise Stop('Повреждены участники компенсации.')
    for item in r['participants']:
        role=item['role'];before=stage/('before-'+role+Path(item['path']).suffix);after=stage/('after-'+role+Path(item['path']).suffix)
        if (item['path']!=expected[role] or any(item[k]!=approved[role][k] for k in ['path','before_sha256','after_sha256'])
                or item['before_copy']!=str(before) or item['after_copy']!=str(after)
                or sha(safe(before))!=item['before_sha256'] or sha(safe(after))!=item['after_sha256']):raise Stop('Повреждены копии компенсации '+role)
    audit=(stage/'audit.jsonl').read_bytes();size=r['audit_prefix_size']
    if hashlib.sha256(audit[:size]).hexdigest()!=r['audit_prefix_sha256']:raise Stop('Повреждён audit trail компенсации.')
    tail=[json.loads(line) for line in audit[size:].splitlines()]
    if len(tail)!=1 or tail[0].get('event')!='CLOSING_COMPENSATED' or tail[0].get('receipt_sha256')!=info['receipt_sha256']:raise Stop('Не подтверждено завершение audit trail.')
    unpublished=journal.parent/'unpublished_documents'
    if r.get('unpublished_documents')!=str(unpublished):raise Stop('Неверный путь неопубликованного комплекта.')
    _documents(old,unpublished)
    return r

def _context(manifest):
    manifest=safe(manifest);m=_sealed(manifest)
    if m.get('schema')!=SCHEMA:raise Stop('Неизвестная схема компенсации.')
    from .core import journal_read
    journal=safe(m['journal']);state=journal_read(journal)
    plan_path=safe(state['plan']);plan=read(plan_path);seal=plan.pop('seal',None)
    if (plan.get('schema')!='VZMK_PRODUCTION_PLAN_1' or digest(plan)!=seal or sha(plan_path)!=m['plan_sha256']
            or sha(journal)!=m['journal_sha256'] or state['id']!=m['id'] or state['id']!=plan['id']
            or state['phase']!='NEEDS_READBACK' or journal.parent.name!=state['id']
            or journal.parent.parent!=safe(state['registry']).parent/'.vzmk-operations'
            or state['target']!=plan['target'] or state['registry']!=plan['registry']):raise Stop('Изменилась исходная операция компенсации.')
    expected={'registry':plan['registry'],'catalog':plan['accounting']['catalog'],'uzk':plan['accounting']['uzk']}
    if (len(m['participants'])!=3 or len(state['participants'])!=3
            or {i['role'] for i in m['participants']}!=set(expected)
            or {i['role'] for i in state['participants']}!=set(expected)
            or len({i['path'] for i in m['participants']})!=3):raise Stop('Повреждён состав компенсации.')
    old={i['role']:i for i in state['participants']}
    for item in m['participants']:
        role=item['role'];o=old[role]
        if item['path']!=expected[role] or item['path']!=o['path']:raise Stop('Неверный путь компенсации.')
        for key,h in [('backup',o['original_sha256']),('candidate',o['candidate_sha256'])]:
            if sha(safe(o[key]))!=h:raise Stop('Повреждена исходная копия '+role)
        for key,h in [('before_copy',item['before_sha256']),('candidate',item['after_sha256']),('path',item['before_sha256'])]:
            if sha(safe(item[key]))!=h:raise Stop('Изменён файл компенсации '+role+' '+key)
    for proof in m['proofs']:
        if sha(safe(proof['path']))!=proof['sha256'] or read(proof['path']).get('status')!=proof['status']:raise Stop('Изменилось доказательство проверки кандидатов.')
    _documents(state,state['target'])
    return manifest,m,journal,state

def apply(manifest,approval):
    manifest,m,journal,state=_context(manifest)
    if (approval.get('approved') is not True or approval.get('manifest_sha256')!=sha(manifest)
            or not approval.get('owner_statement') or not approval.get('at')):raise Stop('Нет явного подтверждения владельца для этих кандидатов.')
    stage=journal.parent/'compensation';unpublished=journal.parent/'unpublished_documents'
    if stage.exists() or unpublished.exists():raise Stop('У компенсации уже есть история; повторная запись запрещена.')
    from .core import journal_write
    original_bytes=journal.read_bytes();old=deepcopy(state);items=deepcopy(m['participants'])
    replaced=[];attempted=[];moved=False;committed=False
    with ExitStack() as stack:
        current=_lock_all(stack,items,True)
        # Recheck all preconditions after obtaining all participant locks.
        if sha(journal)!=m['journal_sha256']:raise Stop('Журнал изменился перед компенсацией.')
        for item in items:
            if hashlib.sha256(current[item['path']]).hexdigest()!=item['before_sha256']:raise Stop('Новые изменения рабочего файла '+item['role'])
            item['before_identity']=windows_file_identity(item['path'])
        target=safe(state['target']);_documents(state,target)
        stage.mkdir(exist_ok=False)
        _write_bytes(stage/'journal.original.json',original_bytes)
        _write_bytes(stage/'manifest.json',Path(manifest).read_bytes())
        for item in items:
            role=item['role'];suffix=Path(item['path']).suffix
            item['before_copy']=str(stage/('before-'+role+suffix));item['after_copy']=str(stage/('after-'+role+suffix))
            _write_bytes(item['before_copy'],current[item['path']])
            payload=Path(safe(item['candidate'])).read_bytes()
            if hashlib.sha256(payload).hexdigest()!=item['after_sha256']:raise Stop('Кандидат изменился перед записью.')
            _write_bytes(item['after_copy'],payload)
            local=Path(item['path']).with_name('.vzmk-compensation-'+uuid4().hex+suffix)
            _write_bytes(local,payload);item['local_candidate']=str(local)
            item['replaced_current']=str(stage/('replaced-current-'+role+suffix))
        intent={'schema':SCHEMA,'id':state['id'],'journal':str(journal),'original_journal_sha256':m['journal_sha256'],
                'manifest_sha256':sha(manifest),'participants':items,'approval':approval,'target':str(target)}
        _write_new_json(stage/'intent.json',intent)
        _event(stage,'INTENT_VERIFIED')
        state['phase']='COMPENSATING';state['compensation']={'schema':SCHEMA,'manifest_sha256':sha(manifest)}
        try:
            journal_write(journal,state)
            for item in items:
                _event(stage,'REPLACE_INTENT',role=item['role'])
                attempted.append(item)
                atomic_replace(item['local_candidate'],item['path'],backup=item['replaced_current'])
                replaced.append(item)
                # Guard each new file before proceeding to the next participant.
                payload=stack.enter_context(locked_file(item['path'],writable=True))
                if hashlib.sha256(payload).hexdigest()!=item['after_sha256']:raise Stop('Не совпал записанный SHA '+item['role'])
                _event(stage,'REPLACED',role=item['role'],sha256=item['after_sha256'])
            _documents(state,target);target.rename(unpublished);moved=True
            _event(stage,'DOCUMENTS_UNPUBLISHED')
            audit=(stage/'audit.jsonl').read_bytes()
            receipt={'schema':SCHEMA,'id':state['id'],'original_journal_sha256':m['journal_sha256'],
                     'manifest_sha256':sha(manifest),'participants':items,'approval':approval,
                     'unpublished_documents':str(unpublished),'closed_at':datetime.now(timezone.utc).isoformat(),
                     'audit_prefix_size':len(audit),'audit_prefix_sha256':hashlib.sha256(audit).hexdigest()}
            _write_new_json(stage/'receipt.json',receipt)
            _event(stage,'CLOSING_COMPENSATED',receipt_sha256=sha(stage/'receipt.json'))
            closed=deepcopy(old);closed['phase']='COMPENSATED';closed['compensation']={'schema':SCHEMA,
                'receipt':str(stage/'receipt.json'),'receipt_sha256':sha(stage/'receipt.json'),
                'original_journal_sha256':m['journal_sha256'],'manifest_sha256':sha(manifest)}
            journal_write(journal,closed);committed=True
        except BaseException as exc:
            # Independent before snapshots allow restoration while old file
            # handles remain open. Never roll back to the pre-user ORIGINAL.
            errors=[]
            for item in reversed(attempted):
                try:
                    if not Path(item['path']).exists():
                        backup=Path(item['replaced_current'])
                        if backup.exists():
                            now=windows_file_identity(backup);was=item['before_identity']
                            if (now['volume'],now['file_id'])!=(was['volume'],was['file_id']):raise Stop('Неизвестная копия при частичной замене.')
                            backup.rename(item['path'])
                        else:_write_bytes(item['path'],Path(item['before_copy']).read_bytes())
                        continue
                    now=windows_file_identity(item['path']);was=item['before_identity']
                    if (now['volume'],now['file_id'])==(was['volume'],was['file_id']):continue
                    if item not in replaced and sha(post_commit_path(item['path']))!=item['after_sha256']:
                        raise Stop('Неизвестное изменение при сбое '+item['role'])
                    restore=Path(item['path']).with_name('.vzmk-compensation-restore-'+uuid4().hex+Path(item['path']).suffix)
                    _write_bytes(restore,Path(item['before_copy']).read_bytes())
                    atomic_replace(restore,item['path'],backup=stage/('failed-after-'+item['role']+Path(item['path']).suffix))
                except BaseException as error:errors.append(str(error))
            if moved or unpublished.exists() and not target.exists():
                try:unpublished.rename(target)
                except BaseException as error:errors.append(str(error))
            if not errors:
                _restore_journal(journal,original_bytes)
                _event(stage,'FAILED_RESTORED_CURRENT',error=str(exc))
            else:_event(stage,'RECOVERY_REQUIRED',error=str(exc),recovery_errors=errors)
            raise Stop('Компенсация прервана; '+('CURRENT восстановлен.' if not errors else 'требуется проверяемое восстановление: '+str(errors))+' '+str(exc)) from exc
    if not committed:raise Stop('Компенсация не завершена.')
    validate_closed(journal,read(journal))
    for item in items:
        if sha(post_commit_path(item['path']))!=item['after_sha256']:raise Stop('Файл изменился после закрытия компенсации '+item['role'])
    return {'status':'PASS','phase':'COMPENSATED','journal':str(journal),'receipt':str(stage/'receipt.json'),
            'unpublished_documents':str(unpublished),'sha256':{i['role']:i['after_sha256'] for i in items}}

def recover_interrupted(journal):
    """Finish a failed intent by restoring the verified pre-compensation set."""
    journal=safe(journal);stage=journal.parent/'compensation';intent=_sealed(stage/'intent.json')
    if intent.get('schema')!=SCHEMA or intent.get('journal')!=str(journal) or intent.get('id')!=journal.parent.name:raise Stop('Повреждена привязка восстановления компенсации.')
    original=stage/'journal.original.json'
    if sha(safe(original))!=intent['original_journal_sha256']:raise Stop('Повреждён исходный журнал компенсации.')
    old=_sealed(original);items=intent['participants']
    if old['id']!=intent['id'] or old['phase']!='NEEDS_READBACK':raise Stop('Неверный исходный журнал компенсации.')
    expected={i['role']:i['path'] for i in old['participants']}
    if len(items)!=3 or {i['role'] for i in items}!=set(expected):raise Stop('Повреждены участники восстановления.')
    from .core import journal_read
    live=journal_read(journal);base=deepcopy(live);base.pop('compensation',None);base['phase']='NEEDS_READBACK'
    if (live['phase']!='COMPENSATING' or base!=old
            or live.get('compensation',{}).get('schema')!=SCHEMA
            or live['compensation'].get('manifest_sha256')!=intent['manifest_sha256']
            or sha(safe(stage/'manifest.json'))!=intent['manifest_sha256']):raise Stop('Нет проверенной прерванной компенсации.')
    with ExitStack() as stack:
        payloads=_lock_all(stack,items,True)
        for item in items:
            if (item['path']!=expected[item['role']] or item['before_copy']!=str(stage/('before-'+item['role']+Path(item['path']).suffix))
                    or sha(safe(item['before_copy']))!=item['before_sha256']
                    or hashlib.sha256(payloads[item['path']]).hexdigest() not in [item['before_sha256'],item['after_sha256']]):raise Stop('Неизвестное изменение; восстановление компенсации запрещено.')
        target=safe(intent['target']);unpublished=journal.parent/'unpublished_documents'
        if target.exists()==unpublished.exists():raise Stop('Неизвестное состояние комплекта компенсации.')
        _documents(old,target if target.exists() else unpublished)
        for item in items:
            if hashlib.sha256(payloads[item['path']]).hexdigest()==item['before_sha256']:continue
            restore=Path(item['path']).with_name('.vzmk-compensation-restore-'+uuid4().hex+Path(item['path']).suffix)
            _write_bytes(restore,Path(item['before_copy']).read_bytes());atomic_replace(restore,item['path'],backup=stage/('interrupted-after-'+uuid4().hex+Path(item['path']).suffix))
        if unpublished.exists():unpublished.rename(target)
        _restore_journal(journal,original.read_bytes());_event(stage,'INTERRUPTED_RESTORED_CURRENT')
    if any(sha(safe(i['path']))!=i['before_sha256'] for i in items):raise Stop('Не подтверждено восстановление CURRENT.')
    return {'status':'PASS','phase':'NEEDS_READBACK','compensation':'FAILED_RESTORED_CURRENT'}
