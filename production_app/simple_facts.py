"""RC3 small facts screen: preserve draft inputs without issuing a passport.

Only mark/profile/AKZ/UZK quantity and currently measured climate ranges are
visible. Back is always allowed, even when fields are incomplete. A subsequent
Generate action performs independent strict validation of facts, protocols and
forms; never assert VIK/UZK/AKZ 'ГОДЕН' from VOM or registry quantities.
"""
from copy import deepcopy
from pathlib import Path

from .storage import Stop
from .measurement_ranges import measurements_from_ranges


def save_screen_draft(previous, project, number, training, visible_marks,
                      hidden_by_mark, climate_drafts):
    """Retain all existing certified source data, but never certify new data."""
    previous=deepcopy(previous or {})
    if not isinstance(visible_marks,list):
        raise Stop('Некорректные строки ВОМ в окне.')
    seen=set()
    rows=[]
    for item in visible_marks:
        if not isinstance(item,dict):
            raise Stop('Неизвестная строка ВОМ.')
        mark=str(item.get('mark','')).strip()
        if not mark or mark in seen:
            raise Stop('В окне дублируется или отсутствует марка.')
        seen.add(mark)
        previous_marks={str(r['mark']):r for r in previous.get('marks',[])}
        old=previous_marks.get(mark,{})
        hidden=hidden_by_mark.get(mark,{})
        # Keep only known prior facts plus currently visible values; a
        # disappearing column must never create a positive verdict.
        row={
            **old,
            **hidden,
            'mark':mark,
            'profile':str(item.get('profile') or '').strip(),
            'akz':item.get('akz',''),
            'uzk_quantity':item.get('uzk_quantity',0),
        }
        rows.append(row)
    draft={
        **previous,
        'project':str(project),
        'number':str(number),
        'kind':'TRAINING' if training else 'ACTUAL',
        'auto_registry_dates':True,
        'no_responsible_input':True,
        'responsible':'',
        'simple_ui':True,
        'marks':rows,
        'akz_drafts':deepcopy(climate_drafts),
    }
    # The operator must explicitly approve final documents AFTER PDF
    # preview, not upon using Back. Draft cannot authorize anything.
    draft['confirmed']=False
    draft['authorized']=False
    draft['forms_approved']=False
    draft.pop('documents_reviewed',None)
    draft['climates']=[]
    return draft


def prepare_for_check(saved):
    """Validate supplied current ranges for actually assigned AKZ only.

    No writes; the returned facts may still lack genuinely performed
    inspections. Independent core validation must reject those gaps.
    """
    if not isinstance(saved,dict) or not saved.get('simple_ui'):
        return deepcopy(saved)
    data=deepcopy(saved)
    variants=set()
    for row in data.get('marks',[]):
        raw=row.get('akz')
        if raw is None or isinstance(raw,str) and not raw.strip():
            row['akz']=None
        else:
            if isinstance(raw,bool) or str(raw).strip() not in ('0','1','2','3'):
                raise Stop('Марке '+str(row['mark'])+
                           ' назначен некорректный № АКЗ. Допустимы 0–3 '
                           'или пустое поле.')
            row['akz']=int(raw)
            variants.add(row['akz'])
    drafts=data.get('akz_drafts') or {}
    climates=[]
    for number in sorted(variants):
        inputs=drafts.get(str(number))
        if not isinstance(inputs,dict):
            raise Stop('Не введён климат АКЗ №'+str(number)+
                       '. Откройте «Фактические результаты».')
        coats=inputs.get('coats')
        if type(coats) is not int or coats not in (1,3):
            raise Stop('Не определён утверждённый 1/3-слойный формат '
                       'АКЗ №'+str(number)+'.')
        measurements={
            key:deepcopy(inputs.get(key))
            for key in ('air','surface','humidity')
        }
        measurements['layer_intervals']=deepcopy(
            inputs.get('layer_intervals',[]))
        climate=measurements_from_ranges(number,measurements,coats=coats)
        # Paint type comes from the CURRENT project's catalogue, never
        # hidden first-passport findings. Do not make up missing paint.
        colors=sorted({
            str(row.get('paint') or '').strip()
            for row in data['marks']
            if row['akz']==number and str(row.get('paint') or '').strip()
        })
        climate['material']=', '.join(colors)
        # Preserve an explicitly confirmed/current AKZ source from an
        # earlier entered session; never synthesize protocol filenames.
        specific=[
            proof for proof in (data.get('akz_proofs') or {}).get(str(number),[])
            if proof.get('role')=='AKZ' and proof.get('variant')==number
        ]
        climate['basis']='; '.join(
            Path(proof['path']).name for proof in specific
        )
        details=inputs.get('layer_details')
        if details:
            climate['layer_details']=deepcopy(details)
        climates.append(climate)
    data['climates']=climates
    return data
