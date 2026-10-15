"""Exact human decision; not a source importer and not a Stage 5 implementation."""
from pathlib import Path
import sys
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
from profiles_stage4 import storage as s
from profiles_stage4.model import Catalog,BENT_RULE,CLOSED_RULE,CHANNEL_RULE,APPROVED
from profiles_stage4.offline import read_json,sha

def prepare():
    baseline=read_json(OUT/'baseline_generation'/s.CATALOG)
    cat=Catalog(baseline['records'],baseline['sources'],baseline['user_state'])
    project=ROOT.parent/'Гараж казахстан'
    kmd=str(project/'WAA0007144-AC2-КМД.pdf')
    attachment=Path('C:/Users/arara/.codex/attachments/75f98616-525d-42ff-8c62-88e21aaabc77/Вставленный текст.txt')
    decisions=[
        {'profile':'Гн.[140X60X4','family':'Гнутый швеллер','standard':'ГОСТ 8278-83','geometry':{'h':140,'b':60,'t':4},
         'references':[{'file':kmd,'pages':[413,414],'parts':[536,527]}]},
        {'profile':'Гн.[]120х3','family':'Квадратный профиль','standard':'ГОСТ 30245-2003','geometry':{'h':120,'b':120,'t':3},
         'references':[{'file':kmd,'pages':[458],'parts':[718]},{'file':str(project/'Детали.pdf'),'pages':[531]}]},
    ]
    for decision in decisions:
        state=s.prepare_manual_state(cat,decision['profile'],decision['family'],decision['standard'],decision['geometry'],
            'Гараж Казахстан; USER_CONFIRMED_UZK_RULE; '+str(decision['references']),
            verification={'kind':'KMD','confirmed_by':'USER','decision_date':'2026-10-07','references':decision['references']})
        cat=Catalog(cat.raw_records,cat.sources,state)
    state=cat.user_state
    common={'basis':'USER_CONFIRMED_UZK_RULE_20261007','status':APPROVED,'date':'2026-10-07','decision_attachment':str(attachment),'decision_sha256':sha(attachment)}
    state['uzk_approvals']=[
        {**common,'rule':BENT_RULE,'key':'Гн.[140X60X4','families':['Гнутый швеллер'],'standard':'ГОСТ 8278-83','references':decisions[0]['references']},
        {**common,'rule':CLOSED_RULE,'families':['Квадратный профиль','Прямоугольный профиль'],'standard':'ГОСТ 30245-2003','references':decisions[1]['references']},
        {**common,'rule':CHANNEL_RULE,'key':'16П','families':['Швеллер'],'standard':'ГОСТ 8240-97','references':[{'historical_marks':['Бп2-6','Бп2-13','Бп2-1'],'source':'Explicit user acceptance contract 2026-10-07'}]},
    ]
    s.dump(OUT/'user_confirmed_decision.json',{'date':'2026-10-07','source':str(attachment),'sha256':sha(attachment),'manual_records':decisions,'approvals':state['uzk_approvals'],
        'alias_policy':'Lookup-only leading channel [; exactly one existing rolled-channel identity; no collisions; exact raw key wins',
        'KMD_reaudit':'NOT_RUN; verified facts provided explicitly by user','scope':'ONE_ITEM; Stage 5 implementation excluded'})
    s.dump(OUT/'proposed_user_state.json',state)
    return state

if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    state=prepare(); print('Prepared two verified manual records and three explicit rule scopes; no DEV writes')
