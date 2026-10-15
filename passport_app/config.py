import json
from dataclasses import dataclass, field
from pathlib import Path


class PassportError(Exception):
    """Понятная ошибка проверки или формирования."""


@dataclass(frozen=True)
class Settings:
    documents_root: Path
    registry_path: Path
    projects: dict = field(default_factory=dict)

    @classmethod
    def load(cls, path):
        path = Path(path).resolve()
        data = json.loads(path.read_text(encoding='utf-8-sig'))
        def resolve(value):
            return (path.parent / value).resolve()
        return cls(resolve(data['documents_root']), resolve(data['registry_path']), data.get('projects', {}))

    def project_paths(self, project):
        item = self.projects.get(project)
        if item:
            root = (self.documents_root / item['passport_root']).resolve()
            template = (self.documents_root / item['template']).resolve() if item.get('template') else None
        else:
            if any(c in project for c in '<>:"/\\|?*'):
                raise PassportError(f'Для проекта «{project}» нужно указать фактические пути в settings.json.')
            root = self.documents_root / project / 'Паспорт'
            template = None
        if not root.is_relative_to(self.documents_root.resolve()):
            raise PassportError('Папка паспорта выходит за корневую папку документов.')
        if not root.is_dir():
            raise PassportError(f'Не найдена папка паспортов: {root}. Укажите фактический путь в settings.json.')
        if template and not template.is_relative_to(self.documents_root.resolve()):
            raise PassportError('Шаблон выходит за корневую папку документов.')
        return root, template
