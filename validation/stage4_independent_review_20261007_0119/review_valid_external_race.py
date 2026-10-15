"""Confirm the build race also overwrites valid external Excel/JSON edits."""
from pathlib import Path
source=Path(__file__).with_name('review_fixtures.py').read_text(encoding='utf-8')
exec(compile(source[:source.index('# Real offline renderer/build:')],str(Path(__file__).with_name('review_fixtures.py')),'exec'))
import zipfile
from xml.etree import ElementTree as ET
original_payload=s.workbook_payload(s.load_catalog())
results=[]
for f in ('Профили.xlsx',s.CATALOG,s.STATE):
    t=clone('valid_external_during_'+str(FILES.index(f)));edit={}
    def render_and_edit(payload,destination,previews=None):
        assert read(payload)==original_payload
        shutil.copy2(s.DEFAULT_WORKBOOK,destination)
        p=t.parent/f
        if f=='Профили.xlsx':
            with zipfile.ZipFile(p) as z:parts={n:z.read(n) for n in z.namelist()}
            ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            tree=ET.fromstring(parts['xl/worksheets/sheet1.xml'])
            cell=tree.find('.//{'+ns+'}c[@r="T2"]');assert cell is not None
            for child in list(cell):cell.remove(child)
            cell.set('t','inlineStr');ss=ET.SubElement(cell,'{'+ns+'}is');ET.SubElement(ss,'{'+ns+'}t').text='VALID_EXTERNAL_REVIEW_NOTE'
            parts['xl/worksheets/sheet1.xml']=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
            with zipfile.ZipFile(p,'w',zipfile.ZIP_DEFLATED) as z:
                for n,b in parts.items():z.writestr(n,b)
            from openpyxl import load_workbook
            wb=load_workbook(p,read_only=True);assert wb['Профили']['T2'].value=='VALID_EXTERNAL_REVIEW_NOTE';wb.close()
        else:
            data=read(p)
            if f==s.STATE:data['preferences'][0]['basis']='VALID_EXTERNAL_REVIEW_BASIS'
            else:data['records'][0]['notes'].append('VALID_EXTERNAL_REVIEW_NOTE')
            p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
            assert read(p)==data
        edit['sha256']=sha(p)
        return {'network_attempts':0}
    with patch.object(s,'render_workbook',render_and_edit):built=s.build(t)
    assert sha(t.parent/f)!=edit['sha256'] and sha(Path(built['backup'])/f)==edit['sha256']
    backup=Path(built['backup']);commit=read(backup/s.COMMIT)
    mismatch=[name for name,digest in commit['files'].items() if sha(backup/name)!=digest]
    assert f in mismatch
    results.append({'file':f,'external_edit_valid':True,'build_returned_success':True,
      'external_edit_silently_overwritten':True,'edited_sha256':edit['sha256'],'backup_has_edited_bytes':True,
      'backup_generation_hash_mismatches':mismatch})
    save('valid_external_race.json',results)
    print(json.dumps(results[-1],ensure_ascii=False),flush=True)
assert all(sha(p)==h for p,h in reviewed.items())
