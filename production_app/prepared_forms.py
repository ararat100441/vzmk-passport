"""Find INTERNAL cleared dynamic forms belonging to exactly one project.

These are previously prepared blank files placed under the application's
Templates/<project> directory. The user's first ISSUED documents provide the
visual reference, and Appendix #1 is taken directly from the first source.
Filled VIK/UZK/quality/AKZ are NEVER treated as cleaned forms.

This fallback prevents a manual 'Шаблоны объекта' dialog when an internal
project bundle already exists. Absence of approved blanks is an explicit
blocker, not permission to reuse historical inspection outcomes.
"""
from pathlib import Path

from .storage import Stop,sha,safe


def _key(value):
    return ' '.join(str(value).replace('ё','е').casefold().split())


def prepared_project(cfg,name):
    root=safe(cfg['documents'])
    candidates=[p for p in root.iterdir()
                if p.is_dir() and _key(p.name)==_key(name)
                and not p.is_symlink()
                and not getattr(p,'is_junction',lambda:False)()]
    if len(candidates)!=1:
        raise Stop('Нет единственного набора очищенных форм для проекта '+
                   name+' внутри программы. Найдено: '+str(len(candidates))+
                   '. Не допускается выдавать заполненные документы первого '
                   'паспорта за чистые бланки.')
    folder=safe(candidates[0])
    all_files=[p for p in folder.iterdir() if p.is_file()
               and not p.name.startswith('~$')]
    role_files={}
    for p in all_files:
        key=_key(p.stem)
        role=None
        if p.suffix.lower()=='.docx':
            if 'вик' in key:role='vik'
            elif 'узк' in key:role='uzk'
            elif 'качеств' in key or key=='quality':role='quality'
        elif p.suffix.lower()=='.xlsx':
            if key in ('first_appendix','appendix'):role='appendix'
            elif key.startswith('akz_appendix') and key[-1] in '0123':
                role='akz_appendix'+key[-1]
            elif key.startswith('akz') and key[-1] in '0123':
                role='akz'+key[-1]
        if role is None:continue
        if role in role_files:
            raise Stop('Два подготовленных шаблона для '+role+' у проекта '+name)
        role_files[role]=p
    required={'vik','uzk','quality'}
    missing=required-set(role_files)
    if missing:
        raise Stop('Нет очищенных форм для '+name+': '+
                   ', '.join(sorted(missing))+
                   '. Исторические заполненные акты нельзя использовать '
                   'без безопасной подготовки рабочих форм.')
    templates={}
    for role,p in role_files.items():
        original=safe(p)
        if not original.is_relative_to(folder):
            raise Stop('Чистая форма находится вне папки выбранного объекта.')
        templates[role]={'path':str(original),'sha256':sha(original)}
    return {'root':str(folder),'templates':templates,
            'forms_origin':'AUTO_DISCOVERED_PREPARED_INTERNAL'}
