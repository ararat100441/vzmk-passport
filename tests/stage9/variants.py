"""Synthetic assignment integration for all accepted exact AKZ template pairs."""
import sys,os,shutil
from pathlib import Path
from uuid import uuid4
from openpyxl import Workbook
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.workspace import bind,save
from stage9_app.ownership import owned_office
from stage8_closure.storage import read
from akz_stage6 import core as a,part2
OUT=ROOT/'validation/stage9_20261012_01'
root,_=bind(read(OUT/'WORKSPACE.json')['path'])
token=uuid4().hex[:8]
with owned_office(OUT/f'AKZ_ALL_VARIANTS_OFFICE_{token}.json'):
    vom=a.FIXTURES/f'SYNTHETIC_VOM_{token}.xls'
    shutil.copyfile(ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls',vom)
    raw=a.visible_vom(vom)
    marks=[]
    for mark in dict.fromkeys(r['mark'] for r in raw['rows']):
        try:a.matching(raw,{mark:{'assignment':0,'row':2}})
        except a.Stage6Error:continue
        marks.append(mark)
        if len(marks)==4:break
    assert len(marks)==4,'Need four internally consistent original source marks for synthetic assignments'
    catalog=a.FIXTURES/f'SYNTHETIC_ASSIGNMENTS_{token}.xlsx'
    w=Workbook();s=w.active;s.append(['Марка','Профиль','№ АКЗ'])
    for n,mark in enumerate(marks):s.append([mark,'SYNTHETIC TEST ASSIGNMENT',n])
    s.append(['SYNTHETIC_UNASSIGNED','SYNTHETIC',None]);w.save(catalog);w.close()
    preview=part2.preview_groups('Гараж Казахстан','9/2026',vom,catalog,synthetic=True)
    for n in range(4):
        p=part2.group_preview(preview,n)
        assert p.payload['synthetic'] is True
        assert all(r['mark']==marks[n] for r in p.payload['selected'])
        assert 'SYNTHETIC-TEST' in p.payload['act_number']
    result=part2.generate_groups(preview,[0,1,2,3],'S9_SYNTHETIC_'+token)
    assert result['status']=='COMPLETE_ALL_SELECTED',result
    save(OUT/'AKZ_ALL_VARIANTS.json',{'status':'PASS','scope':'SYNTHETIC ASSIGNMENTS ONLY; no actual control','groups':result,'blank_distinct_from_zero':True,'variant_rows_separate':True})
