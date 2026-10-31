"""Owner's explicit CURRENT-control declaration, never a simulated protocol.

The pre-release engineering plan contains no positive verdict. Only a single
explicit PDF approval may create this declaration, sealed to that exact plan.
"""
from copy import deepcopy
from pathlib import Path
from .storage import Stop,sha,read

SCHEMA='VZMK_CURRENT_CONTROL_OWNER_1'
STATEMENT=('Подтверждаю, что текущий ВИК выполнен и соответствует требованиям; '
           'плановый УЗК выполнен и соответствует требованиям; '
           'измерения применяемых АКЗ относятся к текущей отгрузке и достоверны. '
           'Исторические заключения и графические подписи образца не являются новым контролем.')

def engineering_facts(saved,profile):
    from .working_facts import form_material
    from .simple_facts import prepare_for_check
    data=deepcopy(saved);data['owner_confirmation_required']=True
    if not data.get('marks'):raise Stop('Откройте «Фактические результаты»: марки текущей ВОМ ещё не проверены.')
    data['evidence']=[]
    for row in data.get('marks',[]):
        raw=str(row.get('akz','')).strip()
        if not str(row.get('profile') or '').strip():raise Stop('Для марки '+str(row['mark'])+' заполните профиль.')
        if raw not in ('0','1','2','3'):raise Stop('Для марки '+str(row['mark'])+' заполните действующее назначение АКЗ. Пустое назначение не равно №0.')
        row['paint']=form_material(profile,int(raw)) or row.get('paint','')
        # These fields are not inferred or copied from an issued sample.
        for key in ('vik_result','uzk_result','vik_basis','basis'):row.pop(key,None)
    return prepare_for_check(data)

def declaration(plan,verified_sha):
    return {'schema':SCHEMA,'project':plan['project'],'number':plan['number'],
            'plan_seal':plan['seal'],'verified_sha256':verified_sha,
            'statement':STATEMENT,'attested':True}

def require_declaration(plan,confirmation):
    owner=confirmation.get('current_control')
    if not isinstance(owner,dict) or owner!=declaration(plan,confirmation.get('verified_sha256')):
        raise Stop('Нет явного заявления владельца о текущем контроле именно этого комплекта.')
    if not confirmation.get('approved') or not confirmation.get('reviewed'):
        raise Stop('Не подтверждены просмотр и выпуск.')
    if confirmation.get('verified_sha256')!=sha(Path(plan['run'])/'verified.json'):
        raise Stop('Подтверждение относится к другому проверенному PDF.')
    return owner
