"""Synchronize current VOM identities without fabricating assignments/findings."""
from copy import deepcopy
from .simple_facts import save_screen_draft

def synchronize(saved,project,number,training,plan,assignments):
    old={str(r['mark']):r for r in (saved or {}).get('marks',[])}
    visible=[]
    for mark,control in plan.items():
        previous=old.get(mark,{});assigned=assignments.get(mark,{})
        visible.append({**previous,'mark':mark,
            'profile':previous.get('profile') or assigned.get('profile',''),
            'akz':previous.get('akz') if previous.get('akz') not in (None,'') else assigned.get('akz',''),
            'paint':previous.get('paint') or assigned.get('paint',''),
            'standard':previous.get('standard',''),
            'uzk_quantity':control['uzk_quantity']})
    draft=save_screen_draft(saved,project,number,training,visible,
        {r['mark']:{'paint':r.get('paint','')} for r in visible},
        deepcopy((saved or {}).get('akz_drafts',{})))
    # A different shipment needs its own final review, even for matching marks.
    for key in ('current_control','current_controls','owner_confirmation_required'):
        draft.pop(key,None)
    return draft
