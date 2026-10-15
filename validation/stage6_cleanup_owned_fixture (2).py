"""Close only the exact isolated duplicate fixture left by the interrupted run."""
import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from akz_stage6.core import FIXTURES
import win32com.client,win32process
E=Path((ROOT/'validation/STAGE6_PART1_CURRENT.txt').read_text(encoding='utf-8'))
target=(FIXTURES/'stage6_akz_part1_20261009_085822_run2/TEST_duplicate_catalog.xlsx').resolve()
result={'target':str(target),'closed':False,'other_books_closed':False}
try:
    app=win32com.client.GetActiveObject('Excel.Application')
    _,pid=win32process.GetWindowThreadProcessId(app.Hwnd)
    result['pid']=pid
    if pid==30896:  # observed process from this task's interrupted native test
        for book in list(app.Workbooks):
            if Path(book.FullName).resolve()==target:
                book.Close(SaveChanges=False);result['closed']=True
        if result['closed'] and app.Workbooks.Count==0:app.Quit()
    else:result['reason']='Different Excel process; left untouched'
except Exception as exc:result['reason']=str(exc)
(E/'owned_fixture_cleanup.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(result)
