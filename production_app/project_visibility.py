"""Visibility belongs to local settings, scoped to the selected registry."""
from pathlib import Path

from .storage import Stop, local_root, mutex, read, safe, save


def closed_projects(cfg):
    registry=cfg.get('registry')
    if not registry:
        return ()
    return tuple(cfg.get('closed_projects',{}).get(str(Path(registry).absolute()).casefold(),()))


def current_visibility(settings,cfg):
    if Path(settings).is_file():
        return {**cfg,'closed_projects':read(settings).get('closed_projects',{})}
    return cfg


def save_settings(settings,cfg):
    with mutex(settings):
        save(settings,current_visibility(settings,cfg))


def ensure_open(cfg,project):
    settings=local_root()/'rc3-marked-settings.json'
    if project in closed_projects(current_visibility(settings,cfg)):
        raise Stop('Объект «'+project+'» закрыт. Восстановите его в настройках путей.')


def set_closed(settings,cfg,project,*,closed):
    from . import core,model

    registry=safe(cfg['registry'])
    with mutex(registry),mutex(settings):
        persisted=read(settings) if Path(settings).is_file() else cfg
        key=str(registry).casefold()
        if str(safe(persisted['registry'])).casefold()!=key:
            raise Stop('Рабочий реестр изменился в другом окне. Откройте настройки заново.')
        scopes=dict(persisted.get('closed_projects',{}))
        names=set(scopes.get(key,()))
        if closed:
            if project not in model.read_registry(registry).projects:
                raise Stop('Объект отсутствует в выбранном реестре.')
            for journal in core.pending(registry):
                state=core.journal_read(journal)
                owner=state.get('project')
                if not owner or owner==project:
                    label=state.get('number') or journal.parent.name
                    raise Stop(f'Объект «{project}»: сначала завершите операцию {label} ({state["phase"]}).')
            names.add(project)
        else:
            names.discard(project)
        scopes[key]=sorted(names)
        save(settings,{**persisted,'closed_projects':scopes})
    return {**cfg,'closed_projects':scopes}
