"""Read-only Tk widget check; manual User E2E remains NOT_RUN."""
import sys, json, hashlib, os
from pathlib import Path
from unittest.mock import patch
sys.stdout.reconfigure(encoding='utf-8')
OUT=Path(__file__).resolve().parent; ROOT=OUT.parents[1]
sys.path.insert(0,str(ROOT))
expected=json.loads((OUT/'reviewed_hashes_before.json').read_text(encoding='utf-8'))
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
assert all(sha(p)==h for p,h in expected.items())
from profiles_stage4.gui import ProfileWindow
from profiles_stage4.storage import DEFAULT_WORKBOOK
app=ProfileWindow();app.withdraw(); checks=[]
try:
    app.update()
    for key,nd,status in [('35Б1','','FOUND_PREFERRED'),('35Б1','ГОСТ 26020-83','FOUND'),
                          ('20Ш1','','AMBIGUOUS'),('L125×80×7','ГОСТ 8510-86','CONFLICT'),('UNKNOWN','','NOT_FOUND')]:
        app.profile.delete(0,'end');app.profile.insert(0,key)
        app.standard.delete(0,'end');app.standard.insert(0,nd)
        app.search();app.update()
        output=app.output.get('1.0','end')
        assert app.status.cget('text')==status and status in output and 'Provenance:' in output
        if key=='35Б1' and not nd:assert '328 × 6' in output and '348 × 9' in output
        checks.append({'key':key,'standard':nd,'status':status,'widget_assertions':'PASS'})
    with patch('profiles_stage4.gui.os.startfile') as start:
        app.open_workbook();start.assert_called_once_with(str(DEFAULT_WORKBOOK))
finally:app.destroy()
assert all(sha(p)==h for p,h in expected.items())
result={'checks':checks,'open_dispatch_mocked':'PASS','native_Excel_open':'NOT_RUN','User_E2E':'NOT_RUN','originals_unchanged':True}
(OUT/'gui_independent.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
