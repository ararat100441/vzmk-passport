from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
text=(ROOT/'tests/stage9_091/ownership_stop.py').read_text(encoding='utf8')
text=text.replace("Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/diagnostics'","Path(SETUP['base'])/'VZMKPassport/diagnostics'")
(ROOT/'tests/stage9_092/ownership_stop.py').write_text(text,encoding='utf8')
text=(ROOT/'tests/stage9_091/gui.py').read_text(encoding='utf8')
text=text.replace("OUT=ROOT/'validation/stage9_091_20261012_01';work=Path(read(max(OUT.glob('lifecycle_*/import_*.json'),key=lambda p:p.stat().st_mtime))['workspace']);results=[]","OUT=ROOT/'validation/stage9_092_20261012_01';work=Path(read(OUT/'LIFECYCLE_FINAL.json')['workspace']);results=[]")
text=text.replace('from stage9_app.gui import App',"import os\nos.environ['LOCALAPPDATA']=str(Path(__import__('json').loads((ROOT/'validation/stage9_092_20261012_01/SETUP.json').read_text(encoding='utf8'))['base'])/'source_gui')\nfrom stage9_app.gui import App")
(ROOT/'tests/stage9_092/gui_source.py').write_text(text,encoding='utf8')
