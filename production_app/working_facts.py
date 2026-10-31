"""Current control entries for the single production workflow.

No verdict is inferred from a VOM, a quantity or an issued reference. Partial
entries survive Back; only the existing final PDF approval authorizes release.
"""
from copy import deepcopy
from pathlib import Path

from .storage import Stop, sha


def required_doc_fields(profile):
    from .real_forms import DOC_FIELD_LABELS
    from .word_controls import inspect_word_fields
    fields=set()
    for role in ('vik','uzk','quality'):
        if role not in profile.get('templates',{}):continue
        ref=selected_form(profile,role)
        if ref.get('exact_reference'):continue
        tokens,controls=inspect_word_fields(ref['path'])
        fields.update((tokens|controls)&set(DOC_FIELD_LABELS))
    return sorted(fields)


def selected_form(profile, role):
    """The same supported historical source that office.render actually uses."""
    original=profile.get('first_passport_sources',{}).get('sources',{}).get(role)
    if original:
        path=original['path']
        if sha(path)!=original['sha256']:
            raise Stop('Изменён исходный документ: '+str(path))
        if role in ('vik','uzk'):
            from .highlighted_reference import supports_vik_first,supports_uzk_first
            supported=(supports_vik_first if role=='vik' else supports_uzk_first)(path)
        elif role=='quality':
            from .quality_from_first import is_exact_first_quality
            supported=is_exact_first_quality(path)
        elif role.startswith('akz') and not role.startswith('akz_appendix'):
            from .akz_from_first import is_first_akz
            supported=is_first_akz(path)
        elif role=='appendix':supported=True
        elif role.startswith('akz_appendix'):
            from .akz_from_first import is_first_akz_appendix
            supported=is_first_akz_appendix(path)
        else:supported=False
        if supported:return {**original,'exact_reference':True}
    inventory=profile.get('first_passport_sources')
    if inventory is not None:
        number=inventory.get('first_number','не указан')
        reason=(inventory.get('errors',{}).get(role) or ('документ несовместим с утверждённым адаптером' if original else 'исходный документ не найден'))
        raise Stop('Исходный паспорт '+str(number)+', документ '+role+': '+reason+'. Другая форма автоматически не подставлена.')
    return profile['templates'][role]


def form_material(profile, variant):
    """A reusable form specification, never a historical measurement/verdict."""
    ref=selected_form(profile,'akz'+str(variant))
    from .akz_from_first import is_first_akz,_cells
    if is_first_akz(ref['path']):
        return str(_cells(ref['path']).get('E18',{}).get('value') or '').strip()
    return ''


def apply_current_controls(saved, *, require_complete=True):
    data=deepcopy(saved)
    controls=data.get('current_controls',{})
    missing=[];proofs=[];missing_profiles=[];missing_akz=[]
    variants=sorted({int(r['akz']) for r in data.get('marks',[])
                     if str(r.get('akz','')).strip() in ('0','1','2','3')})
    roles=['VIK']+(['UZK'] if any(float(r.get('uzk_quantity') or 0)>0
                                 for r in data.get('marks',[])) else [])
    roles+=['AKZ'+str(n) for n in variants]
    labels={'VIK':'ВИК','UZK':'УЗК',**{'AKZ'+str(n):'АКЗ №'+str(n) for n in variants}}
    indexed={}
    for role in roles:
        item=controls.get(role,{})
        if role in ('VIK','UZK') and item.get('result')!='ГОДЕН':
            missing.append(labels[role]+' — подтверждённое заключение «ГОДЕН» текущего контроля')
        raw=str(item.get('path') or '').strip()
        if not raw:
            missing.append(labels[role]+' — первичный протокол текущей отгрузки')
            continue
        path=Path(raw)
        if not path.is_file():
            missing.append(labels[role]+' — файл недоступен: '+raw)
            continue
        # Pin when explicitly saved by Back, not after the file changes.
        actual=sha(path)
        previous=item.get('sha256')
        if previous and previous!=actual:
            missing.append(labels[role]+' — протокол изменён; выберите и проверьте его заново: '+raw)
            continue
        proof={'path':str(path.resolve()),'sha256':actual,'role':role[:3]}
        if role.startswith('AKZ'):proof['variant']=int(role[3:])
        proofs.append(proof);indexed[role]=proof
    for row in data.get('marks',[]):
        if not str(row.get('profile') or '').strip():
            missing_profiles.append(str(row['mark']))
        # Quantities are only a plan; a positive finding is always explicit.
        row['vik_result']=controls.get('VIK',{}).get('result','')
        row['vik_basis']=indexed.get('VIK',{}).get('path','')
        if float(row.get('uzk_quantity') or 0)>0:
            row['uzk_result']=controls.get('UZK',{}).get('result','')
            row['basis']=indexed.get('UZK',{}).get('path','')
        else:
            row['uzk_result']='НЕ ТРЕБУЕТСЯ'
            row['basis']='План общего реестра УЗК: для данной марки 0 шт.'
        raw=str(row.get('akz','')).strip()
        if raw in ('0','1','2','3'):
            material=str(controls.get('AKZ'+raw,{}).get('material') or '').strip()
            if material:row['paint']=material
            elif not str(row.get('paint') or '').strip():
                missing.append('АКЗ №'+raw+' — наименование ЛКМ: '+str(row['mark']))
        elif not str(row.get('no_akz_basis') or '').strip():
            # A zero UZK plan is NOT evidence that coating is unnecessary.
            missing_akz.append(str(row['mark']))
        else:row['basis']+='; АКЗ не требуется: '+str(row['no_akz_basis'])
    for label,marks in (('Профиль',missing_profiles),('Назначение АКЗ',missing_akz)):
        if len(marks)==1:missing.append(label+' марки '+marks[0]+' — заполните данные проекта')
        elif marks:
            missing.append(label+' — не заполнено для '+str(len(marks))+' марок: '+
                           ', '.join(marks[:12])+(' … Все незаполненные строки видны в таблице марок.' if len(marks)>12 else ''))
    data['evidence']=proofs
    data['akz_proofs']={str(n):[indexed['AKZ'+str(n)]] if 'AKZ'+str(n) in indexed else [] for n in variants}
    if require_complete and missing:
        raise Stop('Для формирования пакета заполните в «Фактические результаты»:\n• '+'\n• '.join(dict.fromkeys(missing)))
    return data


def save_controls(values, previous=None):
    """Retain incomplete entries and the selected proof's identity, not consent."""
    result=deepcopy(values)
    for role,item in result.items():
        path=Path(item.get('path') or '')
        old=(previous or {}).get(role,{})
        if str(item.get('path') or '').strip() and path.is_file():
            if old.get('path')==item['path'] and old.get('sha256'):
                item['sha256']=old['sha256']
            else:item['sha256']=sha(path)
    return result
