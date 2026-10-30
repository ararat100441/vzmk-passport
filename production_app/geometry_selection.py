"""Object/profile defaults and explicitly confirmed per-mark exceptions.

Order is the physical order of approved workbook rows, never alphabetical.
Automatic defaults are engineering suggestions, never proof of compliance.
"""
from copy import deepcopy
from pathlib import Path
from profiles_stage4.model import normalize_profile
from .automatic_acts import _approved_rows
from .storage import Stop,sha

class GeometrySelection:
    def __init__(self,path,cfg,project):
        self.path=Path(path);self.pin=sha(self.path);self.known=_approved_rows(self.path)
        if sha(self.path)!=self.pin:raise Stop('Изменился источник Профили.xlsx во время чтения.')
        self.cfg=deepcopy(cfg);self.project=project
        self.shared=self.cfg.setdefault('geometry_profile_choices',{}).setdefault(project,{})
        self.exceptions=self.cfg.setdefault('geometry_mark_exceptions',{}).setdefault(project,{})
        self.legacy=self.cfg.get('geometry_choices',{}).get(project,{})

    def key(self,profile):
        key=normalize_profile(profile)
        if key not in self.known and key.startswith('['):
            rows=self.known.get(key[1:],[])
            if rows and all(p['family']=='Швеллер' for _,p in rows):key=key[1:]
        return key

    def options(self,profile):
        return list(dict.fromkeys(p['standard'] for _,p in self.known.get(self.key(profile),[]) if p['standard']))

    def resolve(self,mark,profile):
        if not str(profile or '').strip():return ''  # no invented profile
        key=self.key(profile);options=self.options(profile)
        if not options:raise Stop('Марка '+mark+': профиль '+profile+' отсутствует в утверждённом Профили.xlsx.')
        exception=self.exceptions.get(mark,{})
        chosen=exception.get('standard') if exception.get('profile')==key else None
        # r6 stored explicitly saved mark choices. Preserve them as individual
        # exceptions on first migration, never overwrite them with a new default.
        if not chosen and mark in self.legacy:
            chosen=self.legacy[mark]
            self.exceptions[mark]={'profile':key,'standard':chosen}
        chosen=chosen or self.shared.get(key)
        if chosen and chosen not in options:
            raise Stop('Изменился источник Профили.xlsx: сохранённый ГОСТ «'+chosen+
                       '» для профиля '+profile+' (марка '+mark+') больше недопустим. Проверьте источник и явно выберите ГОСТ заново.')
        return chosen or options[0]

    def choose(self,mark,profile,standard,*,individual):
        key=self.key(profile)
        if standard not in self.options(profile):raise Stop('Неутверждённый ГОСТ для марки '+mark)
        if individual:self.exceptions[mark]={'profile':key,'standard':standard}
        else:
            self.shared[key]=standard
            # Choosing the shared scope explicitly removes THIS mark's
            # exception. Other confirmed exceptions remain independent.
            self.exceptions.pop(mark,None)
        if mark in self.legacy:
            self.legacy=dict(self.legacy);self.legacy.pop(mark)
            self.cfg.setdefault('geometry_choices',{})[self.project]=self.legacy

    def settings(self):return deepcopy(self.cfg)

    def verify_source(self):
        if sha(self.path)!=self.pin:raise Stop('Изменился источник Профили.xlsx после выбора ГОСТ. Откройте фактические результаты заново.')
