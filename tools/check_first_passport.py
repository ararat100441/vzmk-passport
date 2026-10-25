r"""Read-only inventory: locate originals of FIRST passport by actual registry.

Usage in the VZMK source checkout on Windows:

  py -3.12 -B tools/check_first_passport.py ^
    --registry "C:\...\Работа\ВЗМК (2).xlsx" ^
    --documents "C:\...\Работа\Документы" ^
    --project "Гараж Казахстан"

NEVER writes to the official registry, original acts, config, or filesystem.
No Word/Excel COM, GUI, PDF generation or program rebuild is performed.
"""
import argparse
import json
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from passport_app.registry import read_registry
from production_app.model import select
from production_app.first_passport import (
    discover_first_passport,discover_first_passport_for_config,
    validate_source_pins,
)


def main(registry,documents,project_name,*,settings=None):
    if settings:
        from production_app.storage import read
        cfg=read(settings)
        r=read_registry(Path(cfg['registry']))
        project=select(r,project_name)
        inventory=discover_first_passport_for_config(
            cfg,project,require_all=True)
    else:
        r=read_registry(Path(registry))
        project=select(r,project_name)
        inventory=discover_first_passport(documents,project,require_all=True)
    validate_source_pins(inventory)
    return {
        'status':'PASS_READ_ONLY',
        'project':inventory['project'],
        'first_passport':inventory['first_number'],
        'first_date':inventory['first_date'],
        'source_folder':inventory['passport_root'],
        'roles':{
            role:{'path':data['path'],'sha256':data['sha256']}
            for role,data in inventory['sources'].items()
        },
        'catalog':inventory['catalog'],
        'warning':'Исторические ВИК/УЗК/качество/АКЗ не являются '
                  'пустыми утверждёнными формами. Их старые факты '
                  'не переносятся при создании нового паспорта.',
    }


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--registry',help='Только при ручной передаче путей')
    parser.add_argument('--documents',help='Архив первых паспортов, корень')
    parser.add_argument('--settings',help='Точный settings.json установленной RC3')
    parser.add_argument('--project',required=True)
    args=parser.parse_args()
    if not args.settings and (not args.registry or not args.documents):
        parser.error('Нужен --settings либо пара --registry и --documents')
    try:
        print(json.dumps(main(args.registry,args.documents,args.project,
                                   settings=args.settings),
                         indent=2,ensure_ascii=False))
    except Exception as exc:
        print('FAIL_READ_ONLY: '+str(exc),file=sys.stderr)
        raise
