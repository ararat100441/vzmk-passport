"""Internal paths vs the THREE everyday user-selected sources.

The user selects: VZMK registry, working document root, common UZK.xlsm.
Clean tokenized Word/Excel forms remain an INTERNAL software asset, not
another folder the PTO engineer must choose for each passport.
"""
from pathlib import Path
from .storage import Stop, safe
from .shared_uzk import find_shared_uzk


def internal_forms_root(cfg,*,installed_root=None):
    """Retain the already pinned approved forms, otherwise use bundle Templates.

    Never use a historical filled Word/Excel passport as an automatically
    approved clean form. The existing approved profile pins are checked in
    model.profile, and missing forms must yield a precise error.
    """
    old=str(cfg.get('documents') or '').strip()
    if old:
        folder=safe(old)
        if folder.is_dir():
            return str(folder)
    root=Path(installed_root).resolve() if installed_root else Path(__file__).resolve().parents[1]
    folder=safe(root/'Templates')
    if not folder.is_dir():
        raise Stop('Служебные чистые формы отсутствуют в составе программы: '+
                   str(folder)+'. Переустановите комплект без изменения '
                   'рабочих документов.')
    return str(folder)


def visible_settings(cfg,*,installed_root=None):
    """Read-only defaults for Settings screen, not a silent update of registry."""
    registry=str(cfg.get('registry') or '')
    documents=str(cfg.get('output') or '')
    uzk=str(cfg.get('uzk_registry') or '')
    if not uzk and registry:
        candidate=find_shared_uzk({'registry':registry},require=False)
        if candidate:
            uzk=str(candidate)
    return {'registry':registry,'output':documents,'uzk_registry':uzk}


def apply_settings(cfg,*,registry,output,uzk_registry,installed_root=None):
    """Construct safe working config; caller explicitly saves after review."""
    registry=safe(registry)
    output=safe(output)
    uzk_registry=safe(uzk_registry)
    if not registry.is_file() or registry.suffix.lower()!='.xlsx':
        raise Stop('Выберите существующий Excel реестр ВЗМК.')
    if not output.is_dir():
        raise Stop('Выберите существующую папку Рабoта/Документы.')
    if not uzk_registry.is_file() or uzk_registry.suffix.lower()!='.xlsm':
        raise Stop('Выберите существующий общий файл УЗК.xlsm.')
    if registry==uzk_registry:
        raise Stop('Общий УЗК и реестр ВЗМК должны быть разными файлами.')
    clean=internal_forms_root(cfg,installed_root=installed_root)
    return {
        **cfg,
        'registry':str(registry),
        'documents':clean,
        'output':str(output),
        'first_passport_documents':str(output),
        'uzk_registry':str(uzk_registry),
        'auto_first_passport':True,
        'training':False,
    }


def config_for_training(cfg,training_registry,training_output):
    """Read real archive and shared UZK ONLY, write no originals."""
    return {**cfg,'registry':str(training_registry),
            'output':str(training_output),'training':True}
