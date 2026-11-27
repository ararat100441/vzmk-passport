"""Amendments of a recorded, program-issued shipment."""
from collections import Counter, defaultdict
from dataclasses import replace
from datetime import datetime
from decimal import Decimal
from fractions import Fraction
from pathlib import Path
import re

from . import model
from .history_binding import appendix_accounting
from .shared_uzk import _key, planned_control, read_shared_uzk
from .catalog_ledger import shipment_quantities
from .storage import Stop, digest, read, safe, sha


def document_hashes(folder):
    folder=safe(folder)
    if not folder.is_dir():
        raise Stop('Не найдена папка выпущенного паспорта.')
    paths=list(folder.iterdir())
    if any(not safe(p).is_file() for p in paths):
        raise Stop('В папке выпущенного паспорта есть вложенные папки. Корректировка остановлена.')
    return {p.name:sha(p) for p in paths}


def original_shipment(cfg,project,number,profile=None,registry=None):
    if not cfg.get('use_shared_uzk'):
        raise Stop('Для корректировки нужен общий УЗК и справочник остатков.')
    registry=registry or model.read_registry(cfg['registry'])
    selected=model.select(registry,project)
    if number==selected.first_number:
        raise Stop('Корректировка доступна для комплектов с сохранённым учётом выпуска программы.')
    entries=[e for e in registry.projects[project] if e['number']==number]
    if len(entries)!=1:
        raise Stop('В реестре нужна одна строка выбранного паспорта.')
    entry=entries[0]
    day=entry['date'].date() if isinstance(entry['date'],datetime) else entry['date']
    profile=profile or model.profile(cfg,project)
    target=model.output_passport_directory(cfg,project,profile)/(model.filename(number)+' от '+day.strftime('%d.%m.%Y'))
    hashes=document_hashes(target)
    issue_path=target/'ISSUE.json'
    if 'ISSUE.json' not in hashes:
        raise Stop('В выпущенном комплекте нет ISSUE.json с первоначальным учётом УЗК.')
    issue=read(issue_path)
    if (issue.get('project'),issue.get('number'),issue.get('training'))!=(project,number,bool(cfg.get('training'))):
        raise Stop('ISSUE.json относится к другому паспорту или режиму.')
    appendix=target/(model.filename(number)+' - appendix.xlsx')
    if hashes.get(appendix.name)!=issue.get('files',{}).get(appendix.name) or not appendix.is_file():
        raise Stop('Приложение №1 изменено после исходного выпуска.')
    actual=appendix_accounting(appendix,entry)
    vom=model.visible_vom(appendix)
    if vom.digest!=actual['sha256'] or sha(issue_path)!=hashes['ISSUE.json']:
        raise Stop('Данные исходного выпуска изменились во время чтения.')
    facts=issue.get('facts',{})
    rows=facts.get('marks',[])
    marks={str(r['mark']):r for r in rows}
    if len(marks)!=len(rows) or set(marks)!=set(actual['quantities']):
        raise Stop('Первоначальный учёт УЗК не соответствует приложению №1.')
    if any('uzk_manual_quantity' in row or 'uzk_registry_quantity' in row for row in rows):
        from .core import journal_read
        operation=issue.get('operation_id','')
        if not isinstance(operation,str) or not re.fullmatch('[a-f0-9]{32}',operation):
            raise Stop('УЗК: нет номера журнала для подтверждения происхождения контроля.')
        journal=safe(Path(cfg['registry']).parent/'.vzmk-operations'/operation/'journal.json')
        if not journal.is_file():raise Stop('УЗК: отсутствует журнал происхождения контроля '+operation+'.')
        state=journal_read(journal)
        if (state['phase'],state.get('target'),state.get('issue_sha256'))!=('COMPLETE',str(target),hashes['ISSUE.json']):
            raise Stop('УЗК: происхождение контроля не совпадает с завершённым журналом '+operation+'.')
    for name,row in marks.items():
        try:
            auto,manual=historical_control(row)
            total=auto+manual
        except (KeyError,ValueError,TypeError) as exc:
            raise Stop('Не подтверждено исходное количество УЗК '+name) from exc
        if not 0<=total<=actual['quantities'][name]:
            raise Stop('Не подтверждено исходное количество УЗК '+name)
    designed={r['mark'] for r in facts.get('uzk_design',[])}
    if designed!={m for m,r in marks.items() if int(r['uzk_quantity'])}:
        raise Stop('В исходном выпуске отсутствует расчёт УЗК по ранее учтённым маркам.')
    return target,hashes,vom,facts,actual['quantities']


def added_rows(original,combined,origins=None):
    base=Counter(original.rows)
    current=Counter(combined.rows)
    if base==current:
        raise Stop('Состав ВОМ не изменился.')
    if origins is None:
        # Older callers can identify unambiguous changes by whole-row multiplicity.
        added=list((current-base).elements())
        removed=base-current
        if {_key(r[0]) for r in added} & {_key(r[0]) for r in removed}:
            raise Stop('Повторно откройте «Отгрузочную ВОМ», чтобы подтвердить происхождение строк.')
    else:
        indices=origins['indices']
        if origins['appendix_sha256']!=original.digest or len(indices)!=len(combined.rows):
            raise Stop('Исходный комплект изменился. Заново откройте корректировку ВОМ.')
        seen=set();added=[]
        for row,index in zip(combined.rows,indices):
            if index is None:
                added.append(row)
                continue
            if type(index) is not int or not 0<=index<len(original.rows) or index in seen or row!=original.rows[index]:
                raise Stop('Исходная строка ВОМ изменена или повторно использована.')
            seen.add(index)
    canonical={_key(row[0]):row[0] for row in original.rows}
    for row in added:
        if canonical.get(_key(row[0]),row[0])!=row[0]:
            raise Stop('Используйте исходное обозначение марки '+str(canonical[_key(row[0])])+'.')
    return replace(combined,rows=tuple(added))


def historical_control(row):
    from .shared_uzk import _qty

    raw=row['uzk_quantity']
    if raw is None or not str(raw).strip():
        raise Stop('УЗК: отсутствует количество в истории марки '+row['mark'])
    total=_qty(raw,row['mark'],allow_zero=True)
    if 'uzk_manual_quantity' not in row and 'uzk_registry_quantity' not in row:
        # Older protected releases only scheduled control from the shared book.
        return total,0
    auto,manual=row.get('uzk_registry_quantity'),row.get('uzk_manual_quantity')
    if type(auto) is not int or type(manual) is not int or min(auto,manual)<0 or auto+manual!=total:
        raise Stop('УЗК: происхождение ранее учтённого контроля не подтверждено: '+row['mark'])
    return auto,manual


def control_plan(source,project,delta,facts,original_quantities,final_quantities,*,allow_manual=False):
    extra=planned_control(source,project,delta) if delta.rows else {}
    additions=shipment_quantities(delta)
    refunds={}
    plan={}
    for row in facts['marks']:
        mark=str(row['mark'])
        registry_count,manual_count=historical_control(row)
        inspected=registry_count+manual_count
        before=original_quantities[mark]
        retained=final_quantities.get(mark,0)-additions.get(mark,0)
        if retained==before or not inspected:
            refund=0
        elif retained==0:
            refund=registry_count
            manual_count=0
        elif inspected==before and not manual_count:
            refund=before-retained
        else:
            raise Stop(f'УЗК: марка {mark}, из {before} шт. проверено по реестру {registry_count}, '
                       f'вручную {manual_count}. Для выбранной части нельзя подтвердить '
                       'происхождение контроля и длину оставшихся изделий.')
        if refund:refunds[mark]=refund
        historical_auto=registry_count-refund
        historical=historical_auto+manual_count
        added=extra.get(mark,{}).get('uzk_quantity',0)
        limit=extra.get(mark,{}).get('manual_limit',0) if allow_manual else 0
        replacement_min=additions.get(mark,0) if refund else 0
        if added+limit<replacement_min:
            raise Stop(f'УЗК: удаляемые и добавляемые строки марки {mark} меняют объём '
                       'контроля. Требуется подтвердить УЗК заменяемых изделий '
                       'до корректировки паспорта.')
        if mark in final_quantities:
            plan[mark]={'uzk_quantity':historical+added,'registry_retained':historical_auto,
                        'manual_retained':manual_count,'registry_added':added,'manual_limit':limit,
                        'replacement_min':replacement_min,
                        'amendment_original':bool(historical) or not added}
    return plan|{m:{**rec,'registry_added':rec['uzk_quantity']} for m,rec in extra.items() if m not in plan},extra,refunds


def _remaining_geometry_by_segment(mark,rows,original_count,kept_count):
    """Partial removal is safe only with equivalent catalogue geometry per item."""
    geometry={}
    quantities=Counter()
    error='УЗК: для марки '+mark+' нельзя подтвердить геометрию оставшихся изделий.'
    for row in rows:
        count=row.get('control_quantity',row['quantity'])
        reference=row['geometry_reference']
        if type(count) is not int or count<1 or reference.get('source')=='MANUAL':
            raise Stop(error)
        if reference.get('sha256'):
            reference={k:v for k,v in reference.items() if k!='path'}
        segment=tuple(row['segments'])
        signature=(Fraction(Decimal(row['length_mm']))/count,row['profile'],row['thickness_mm'],
                   row['frequency_mhz'],row['angle_deg'],reference)
        if geometry.setdefault(segment,signature)!=signature:
            raise Stop(error)
        quantities[segment]+=count
    if not quantities or any(count!=original_count for count in quantities.values()):
        raise Stop(error)
    return dict.fromkeys(quantities,kept_count)


def retained_design(facts,plan):
    rows=facts.get('uzk_design',[])
    registry_parts=defaultdict(list)
    for row in rows:
        if row.get('control_source','REGISTRY')=='REGISTRY':
            registry_parts[row['mark']].append(row)
    remaining_by_mark={}
    for historical in facts['marks']:
        mark=historical['mark']
        original,_=historical_control(historical)
        kept=plan.get(mark,{}).get('registry_retained',0)
        if 0<kept<original:
            remaining_by_mark[mark]=_remaining_geometry_by_segment(mark,registry_parts[mark],original,kept)

    retained=[]
    for row in rows:
        mark=row['mark'];control=plan.get(mark,{})
        origin=row.get('control_source','REGISTRY')
        if origin=='MANUAL':
            if control.get('manual_retained'):retained.append(row)
            continue
        if not control.get('registry_retained'):continue
        old_count=row.get('control_quantity',row['quantity'])
        count=old_count
        remaining=remaining_by_mark.get(mark)
        if remaining is not None:
            segment=tuple(row['segments'])
            count=min(count,remaining[segment])
            if not count:continue
            remaining[segment]-=count
            if count!=old_count:
                row={**row,'length_mm':format(Decimal(row['length_mm'])*count/old_count,'f')}
        retained.append({**row,'control_source':'REGISTRY','control_quantity':count})
    return retained


def refund_sources(registry,target,source,project,issue_sha,refunds):
    """Trace each refund to the completed debit and its verified workbook."""
    from .core import journal_read

    if not refunds:return {},{}
    stages=safe(registry).parent/'.vzmk-operations'
    by_issue={}
    for journal in stages.glob('*/journal.json'):
        state=journal_read(journal)
        if (state.get('phase')=='COMPLETE' and state.get('registry')==str(registry)
                and state.get('target')==str(target) and state.get('id')==journal.parent.name
                and state.get('issue_sha256')):
            by_issue.setdefault(state['issue_sha256'],[]).append((journal,state))
    remaining=dict(refunds)
    returned=Counter()
    originals={}
    pins={}
    while remaining:
        matches=by_issue.pop(issue_sha,[])
        if len(matches)!=1:
            raise Stop('УЗК: нет единственного завершённого журнала списания для '+', '.join(remaining)+'.')
        journal,state=matches[0]
        pins[str(journal)]=sha(journal)
        if journal_read(journal)!=state:
            raise Stop('УЗК: журнал исходного выпуска изменился.')
        participants=[p for p in state.get('participants',[]) if p['role']=='uzk']
        if len(participants)!=1 or safe(participants[0]['path'])!=source:
            raise Stop('УЗК: журнал списания относится к другому общему реестру.')
        before=state.get('original_documents',{}).get('ISSUE.json')
        if before:
            previous=safe(journal.parent/'original_documents'/'ISSUE.json')
            if not previous.is_file() or sha(previous)!=before:
                raise Stop('УЗК: архив прежнего комплекта отсутствует или изменён.')
            prior=read(previous)
            if prior.get('project')!=project:
                raise Stop('УЗК: архив относится к другому объекту.')
            pins[str(previous)]=before
        item=participants[0]
        from .service_archive import operation_evidence
        plan_path=operation_evidence(journal,state,names=('plan.json',))['plan.json']
        if not plan_path.is_file():
            raise Stop('УЗК: отсутствует исходный план списания для '+', '.join(remaining)+'.')
        plan_sha=sha(plan_path)
        plan=read(plan_path);seal=plan.pop('seal',None)
        accounting=plan.get('accounting') or {}
        approved=(state.get('current_control') or {}).get('plan_seal')
        if (seal!=digest(plan) or (approved and approved!=seal)
                or (plan.get('id'),plan.get('project'),plan.get('target'),plan.get('registry'))
                !=(state['id'],project,str(target),str(registry))
                or accounting.get('uzk')!=str(source)
                or plan.get('fingerprints',{}).get(str(source))!=item['original_sha256']):
            raise Stop('УЗК: исходный план списания не совпадает с завершённой операцией.')
        if sha(plan_path)!=plan_sha:raise Stop('УЗК: исходный план изменился.')
        pins[str(plan_path)]=plan_sha
        debited={}
        for mark in remaining:
            debit=accounting.get('controls',{}).get(mark,0)
            cancelled=min(debit,returned[mark])
            returned[mark]-=cancelled
            returned[mark]+=accounting.get('refunds',{}).get(mark,0)
            if debit>cancelled:debited[mark]=debit-cancelled
        if debited:
            backup=safe(item['backup'])
            if (backup!=journal.parent/'backup-uzk.xlsm' or not backup.is_file()
                    or sha(backup)!=item['original_sha256']):
                raise Stop('УЗК: резервная копия перед списанием отсутствует или изменена.')
            rows=read_shared_uzk(backup,project=project,marks=debited)
            for mark,count in debited.items():
                rec=rows.get((_key(project),_key(mark)),{})
                if len(rec.get('rows',[]))!=1 or rec['remaining']<accounting['controls'][mark]:
                    raise Stop('УЗК: исходное списание не подтверждено для марки '+mark+'.')
                originals.setdefault(mark,{'project':rec['project'],'mark':rec['mark']})
                remaining[mark]-=min(count,remaining[mark])
                if not remaining[mark]:del remaining[mark]
            pins[str(backup)]=item['original_sha256']
        issue_sha=before
    return originals,pins
