r"""READ ONLY: calculate first-passport-style VIK and UZK variable parts.

No office output, no change to VZMK.xlsx, VOM, UZK.xlsm or any issued act.

Native Windows Excel required for filtered/hidden VOM row detection.

Example (paths are explicit; this never runs unsafe DEV fallback):
  py -3.12 tools/inspect_automatic_acts.py
      --project "Гараж Казахстан"
      --vom "C:\...\ВОМ отгрузочный счет.xlsx"
      --uzk "C:\...\Работа\УЗК.xlsm"
      --assignments "C:\...\Гараж казахстан\Справочник марок.xlsx"
      --profiles "C:\...\Справочники\Профили.xlsx"

The profile workbook must be an approved current Stage4 source, not merely
a random Excel file with geometry. This tool prints only the calculated
structural portion; it NEVER assigns defect conclusions or signatures.
"""
import argparse
import json
import sys
from pathlib import Path

root=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(root))
from production_app.model import visible_vom
from production_app.storage import Stop,sha
from production_app.catalog_assignments import current_assignments
from production_app.shared_uzk import planned_control
from production_app.automatic_acts import (
    calculate_design_characteristics,ordered_vik_names,
)


def inspect(project,vom_path,uzk_path,assignments_path,profiles_path):
    targets=[Path(p).resolve(strict=True) for p in (
        vom_path,uzk_path,assignments_path,profiles_path)]
    frozen={str(p):sha(p) for p in targets}
    vom=visible_vom(targets[0])
    planned=planned_control(targets[1],project,vom)
    assignments=current_assignments(
        targets[2],frozen[str(targets[2])],planned,
    )
    facts={'marks':[
        {'mark':mark,
         'profile':assignments[mark].get('profile',''),
         'uzk_quantity':entry['uzk_quantity']}
        for mark,entry in planned.items()
    ]}
    details=calculate_design_characteristics(
        vom,facts,targets[3],expected_sha=frozen[str(targets[3])])
    for p in targets:
        if sha(p)!=frozen[str(p)]:
            raise Stop('Оригинал изменился во время проверки: '+str(p))
    return {
        'status':'STRUCTURAL_ONLY_READ_ONLY',
        'project':project,
        'visible_mark_count':len(planned),
        'vik_marks':list(ordered_vik_names(vom)),
        'uzk_planned_items':sum(x['uzk_quantity'] for x in planned.values()),
        'uzk_design':details,
        'sources_sha256':frozen,
        'disclaimer':'Плановые количества и геометрия не являются '
                     'доказательством фактического контроля или ГОДЕН.',
    }


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    for key in ('project','vom','uzk','assignments','profiles'):
        parser.add_argument('--'+key,required=True)
    args=parser.parse_args()
    report=inspect(args.project,args.vom,args.uzk,
                   args.assignments,args.profiles)
    print(json.dumps(report,ensure_ascii=False,indent=2))
    return 0


if __name__=='__main__':
    try:raise SystemExit(main())
    except (Stop,ValueError,OSError) as exc:
        print('STOP: '+str(exc),file=sys.stderr)
        raise SystemExit(1)
