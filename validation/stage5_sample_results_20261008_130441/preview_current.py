import json
from pathlib import Path
import sys

RUN = Path(__file__).resolve().parent
ROOT = RUN.parents[1]
sys.path.insert(0, str(ROOT))
from uzk_stage5.core import preview
from uzk_stage5.result_samples import load_sample_results

old = json.loads((ROOT/'validation/stage5_implementation_20261008_105100/final_run.json').read_text(encoding='utf-8-sig'))
plan = preview('Гараж Казахстан', '9/2026', old['preview']['vom']['path'])
assert plan.data['marks'] == old['preview']['marks']
assert plan.data['controlled_units'] == 33
sample = load_sample_results(plan.data)
(RUN/'current_preview.json').write_text(json.dumps({'preview': plan.data, 'seal': plan.seal,
    'sample': sample, 'accepted_calculation_unchanged': True}, ensure_ascii=False, indent=2), encoding='utf8')
print(json.dumps({'status':plan.data['status'], 'marks':len(plan.data['marks']),
                  'units':plan.data['controlled_units'], 'sample': [sample['graph7'],sample['graph8']],
                  'source_rows':sample['row_count']}, ensure_ascii=False))
