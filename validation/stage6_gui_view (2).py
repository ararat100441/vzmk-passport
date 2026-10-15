"""Native Tk preview evidence, no generation from this validation helper."""
import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from akz_stage6.gui import Launcher
from akz_stage6.core import sha256
import tkinter as tk
E=Path((ROOT/'validation/STAGE6_PART1_CURRENT.txt').read_text(encoding='utf-8'))
preview=json.loads((E/'preview.json').read_text(encoding='utf-8'))['payload']
catalog=[x for x in preview['sources'] if 'TEST_Справочник марок.xlsx' in x][0]
product_path=ROOT/'akz_stage6/gui.py';product_before=sha256(product_path)
root=tk.Tk();gui=Launcher(root)
gui.vom.set(preview['vom']['path']);gui.catalog.set(catalog)
gui.check()
assert gui.checked is not None and str(gui.generate_button['state'])=='normal'
content=gui.result.get('1.0','end')
assert all(x in content[:1000] for x in ('Бп2-5','Бп2-18','Бп2-19','105.26','3.989','ИТОГО'))
checked=gui.checked
gui.vom.set('changed selection')
assert gui.checked is None and str(gui.generate_button['state'])=='disabled'
gui.vom.set(preview['vom']['path']);gui.check()
assert gui.checked is not None and gui.checked.seal==checked.seal
(E/'native_gui_preview.json').write_text(json.dumps({'status':'PASS','preview_seal':gui.checked.seal,
    'generation_enabled':str(gui.generate_button['state']),'sources':gui.checked.payload['sources']},ensure_ascii=False,indent=2),encoding='utf-8')
root.mainloop()
assert sha256(product_path)==product_before
(E/'native_gui_product_hashes.json').write_text(json.dumps({'BEFORE':{str(product_path):product_before},
    'AFTER':{str(product_path):sha256(product_path)},'CURRENT':{str(product_path):sha256(product_path)},
    'checks':['human preview shows all marks and totals first','change source disables generation',
              'fresh native preview enables generation and has exact same source seal'],'status':'PASS'},ensure_ascii=False,indent=2),encoding='utf-8')
