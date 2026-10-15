from pathlib import Path
import sys,json
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from akz_stage6 import core as c,part2 as p2
catalog=c.FIXTURES/'stage6_part2_20261009_132204/TEST_Справочник марок.xlsx'
vom=c.PROJECT/'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
batch=p2.preview_groups('Гараж Казахстан','9/2026',vom,catalog)
result=p2.generate_groups(batch,[1,2,3],'stage6_part2_20261009_132204_FINAL_TEST_9_2026')
(OUT/'real_generation.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
assert result['status']=='COMPLETE_ALL_SELECTED',result
print(json.dumps(result,ensure_ascii=False),flush=True)
