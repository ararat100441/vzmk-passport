import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from akz_stage6.core import *
E=Path((ROOT/'validation/STAGE6_PART1_CURRENT.txt').read_text(encoding='utf-8'))
v=visible_vom(PROJECT/'Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls')
b=load_workbook(PROJECT/'Справочник марок.xlsx',data_only=True)
marks={row[0] for row in b.active.iter_rows(min_row=2,values_only=True)};b.close()
valid=[];invalid=[]
for r in v['rows']:
    if r['mark'] not in marks:continue
    try:
        matching({'rows':[r]},{r['mark']:{'assignment':0,'row':2}})
        valid.append(r)
    except Stage6Error as exc:invalid.append({'row':r,'error':str(exc)})
(E/'real_vom_read.json').write_text(json.dumps(v,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
(E/'area_assessment.json').write_text(json.dumps({'valid':valid,'invalid':invalid},ensure_ascii=False,indent=2,default=str),encoding='utf-8')
print('TEMPLATES',templates());print('VISIBLE',len(v['rows']),'HIDDEN',len(v['hidden_rows']))
print('VALID',len(valid),'INVALID',len(invalid));print(json.dumps(valid[:6],ensure_ascii=False,default=str))
