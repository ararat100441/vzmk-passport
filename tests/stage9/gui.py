import sys,tkinter as tk,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[2];sys.path.insert(0,str(ROOT))
from stage9_app.gui import App
from stage9_app.workspace import save
from stage8_closure.storage import read
OUT=ROOT/'validation/stage9_20261012_01';work=Path(read(OUT/'WORKSPACE.json')['path']);results=[]
from uuid import uuid4
OUT=OUT/('gui_'+uuid4().hex[:8]);OUT.mkdir()
for scaling in (1.0,1.5,2.0):
    window=tk.Tk();window.tk.call('tk','scaling',scaling);app=App(window,work)
    notebook=app.closure_text.master.master.master
    for size in ('720x520','1000x760','1360x900'):
        window.geometry(size);window.update()
        notebook.select(2);window.update()
        boxes={}
        for name in ('close_button','recover_button','restore_button'):
            b=getattr(app,name);boxes[name]=[b.winfo_rootx()-window.winfo_rootx(),b.winfo_rooty()-window.winfo_rooty(),b.winfo_width(),b.winfo_height()]
            x,y,w,h=boxes[name]
            assert x>=0 and y>=0 and x+w<=window.winfo_width() and y+h<=window.winfo_height(),(scaling,size,name,boxes[name],window.winfo_width(),window.winfo_height())
        results.append({'scale':scaling,'size':size,'buttons':boxes,'status':'PASS'})
    if scaling==1.5:
        # Capture only this owned test window, using PrintWindow (no desktop automation).
        import win32gui,win32ui,ctypes
        from PIL import Image
        hwnd=win32gui.GetParent(window.winfo_id());rect=win32gui.GetWindowRect(hwnd);w,h=rect[2]-rect[0],rect[3]-rect[1]
        dc=win32gui.GetWindowDC(hwnd);src=win32ui.CreateDCFromHandle(dc);dest=src.CreateCompatibleDC();bitmap=win32ui.CreateBitmap();bitmap.CreateCompatibleBitmap(src,w,h);dest.SelectObject(bitmap)
        ctypes.windll.user32.PrintWindow(hwnd,dest.GetSafeHdc(),2)
        info=bitmap.GetInfo();im=Image.frombuffer('RGB',(info['bmWidth'],info['bmHeight']),bitmap.GetBitmapBits(True),'raw','BGRX',0,1);im.save(OUT/'GUI_RECOVERY.png')
        win32gui.DeleteObject(bitmap.GetHandle());dest.DeleteDC();src.DeleteDC();win32gui.ReleaseDC(hwnd,dc)
    # The specific stale-manifest defect must remain reproducibly covered.
    app.manifest='old';app.preview='old';app.vom.set('changed');assert app.manifest is None and app.preview is None
    callbacks=[];original_work=app.work
    app.work=lambda task,args,callback,**kw:callbacks.append(callback)
    app.check();app.number.set('another passport');callbacks.pop()({'preview':'stale','summary':'stale'})
    assert app.preview is None and app.manifest is None and str(app.gen['state'])=='disabled'
    from stage9_app.gui import messagebox
    original_confirm=messagebox.askyesno;messagebox.askyesno=lambda *a,**kw:True
    try:
        app.preview='approved';app.generate();app.vom.set('another BOM');callbacks.pop()({'manifest':'stale generated manifest'})
        assert app.manifest is None and app.preview is None and str(app.gen['state'])=='disabled'
    finally:messagebox.askyesno=original_confirm;app.work=original_work
    app.shutdown()
save(OUT/'GUI_LAYOUT.json',{'pass':len(results),'cases':results,'stale_selection_invalidation':'PASS','async_check_stale_callback':'PASS','async_generate_stale_callback':'PASS'})
print('GUI',len(results),'PASS')
