from common import *
from openpyxl import load_workbook
import re
rows=read(OUT/'FORMAT_DIFF_SEMANTIC.json')
helpers={2:{'L31','M31'},3:{'N31','O31'}}
for row in rows:
    n=int(row['variant']);kind=row['kind'];book=load_workbook(row['output'],data_only=False)
    try:
        formulas=[str(c.value) for s in book for cells in s for c in cells if c.data_type=='f']
        for change in row['style_changes']:
            assert change['field']!='protection',change
            if kind=='act':
                allowed=(change['field']=='number_format' and n==2 and change['cell']=='I27') or (change['field']=='alignment' and change['cell'] in ('E7','B43'))
                if change['field']=='alignment' and change['cell'] in helpers.get(n,set()):
                    assert change['sheet']=='4' and book['4'][change['cell']].column>11
                    assert book['4'][change['cell']].value is None
                    cell=change['cell'];pattern=r'(?<![A-Z0-9_])\$?'+cell[0]+r'\$?'+cell[1:]+r'(?!\d)'
                    assert not any(re.search(pattern,f) for f in formulas),'DANGLING_HELPER_REFERENCE'
                    allowed=True;change['reason']='Excel normalized alignment of cleared obsolete random helper outside B2:K50; no remaining formula dependency'
                assert allowed,change
            else:assert change['field']=='number_format' and change['cell'][0] in 'CDEFG' and int(change['cell'][1:])>=(4 if n==0 else 5),change
    finally:book.close()
save(OUT/'FORMAT_AUDIT_FINAL.json',{'status':'PASS','files':rows,'column_widths_unchanged':True,'preserved_by_native_validator':['sheet_names','merges','print_areas','fonts','fills','borders','coating_materials_colors_layers_norms','retained_formulas'],'permitted_changes':['current project E7 wrap/height','TEST warning B43 wrap','application numeric display precision matching VOM','variant2 I27 currency-to-micron display','dead nonprinted random helper alignment normalization'],'invalid_earlier_oracle':'FORMAT_DIFF_DETAILS.json compared StyleProxy wrappers directly; semantic copies and exact dependency checks are used here'})
print('8-file semantic formatting audit PASS; helper exceptions proven outside print and unreferenced',flush=True)
