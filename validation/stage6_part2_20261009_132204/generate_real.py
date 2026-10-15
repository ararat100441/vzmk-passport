from pathlib import Path
import sys,json,shutil
ROOT=Path(__file__).resolve().parents[2];OUT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT))
from akz_stage6 import core as c,part2 as p2
from openpyxl import load_workbook
def dump(n,x):(OUT/n).write_text(json.dumps(x,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
pre=json.loads((OUT/'preflight.json').read_text(encoding='utf-8'))
fixture=c.FIXTURES/'stage6_part2_20261009_132204';fixture.mkdir(exist_ok=False)
catalog=fixture/'TEST_Справочник марок.xlsx';shutil.copyfile(c.PROJECT/'Справочник марок.xlsx',catalog)
assignments={r['mark']:n for n,r in enumerate(pre['fixture_candidate_rows'],1)}
with c.native_excel() as app:
 b=c.native_open(app,catalog,read_only=False)
 try:
  s=b.Worksheets('Марки');evidence=[]
  for r in range(2,s.UsedRange.Rows.Count+1):
   mark=s.Cells(r,1).Value2
   if mark in assignments:
    s.Cells(r,3).Value2=assignments[mark];evidence.append({'mark':mark,'catalog_row':r,'assignment':assignments[mark]})
  assert len(evidence)==3
  c.native_save(b,catalog)
 finally:b.Close(SaveChanges=False)
dump('DEV_assignments.json',{'catalog':str(catalog),'sha256':c.sha256(catalog),'assignments':evidence,'selected_vom_rows':pre['fixture_candidate_rows']})
vom=c.PROJECT/'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'
batch=p2.preview_groups('Гараж Казахстан','9/2026',vom,catalog)
dump('real_preview.json',{'payload':batch.payload,'seal':batch.seal})
for n in range(1,4):assert batch.payload['groups'][str(n)]['status']=='READY',batch.payload['groups'][str(n)]
result=p2.generate_groups(batch,[1,2,3],'stage6_part2_20261009_132204_TEST_9_2026')
dump('real_generation.json',result);print(json.dumps(result,ensure_ascii=False),flush=True)
assert result['status']=='COMPLETE_ALL_SELECTED',result
native=[]
for n in range(1,4):
 p=p2.group_preview(batch,n);target=Path(result['groups'][str(n)]['path']);names=p2.filenames('9/2026',n)
 before={name:c.sha256(target/name) for name in names[:2]}
 with c.native_excel() as app:
  evidence=p2.validate_saved_pair(app,target/names[0],target/names[1],p)
 after={name:c.sha256(target/name) for name in names[:2]}
 assert before==after
 native.append({'variant':n,'target':str(target),'before':before,'after':after,'evidence':evidence})
 dump('native_real_outputs.json',native)
 print('READ_ONLY_FULL_CALC_PASS',n,flush=True)
