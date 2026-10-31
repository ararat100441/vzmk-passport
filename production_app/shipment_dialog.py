"""One compact seven-column editor using the accepted RC3 Treeview editor."""
import tkinter as tk
from tkinter import ttk
from uuid import uuid4
from .storage import local_root
from .shipment_input import HEADERS,normalize,parse_tsv,write_internal

def replace_rows(grid,rows):
    """Ctrl+V is a complete shipment replacement, independent of selection."""
    for item in grid.get_children():grid.delete(item)
    for row in rows:
        grid.insert('','end',values=tuple('' if v is None else str(v) for v in row))

def open_shipment(app):
    if app.active:return
    window=tk.Toplevel(app.root);window.title('Отгрузочная ВОМ');app.fit(window,1180,680)
    window.transient(app.root);window.grab_set()
    ttk.Label(window,text='Вставьте строки из Excel: Ctrl+V. Двойной щелчок — изменить ячейку. Площадь можно оставить пустой.',wraplength=1120).pack(anchor='w',padx=12,pady=8)
    frame=ttk.Frame(window,padding=10);frame.pack(fill='both',expand=True)
    headings=ttk.Frame(frame);headings.pack(fill='x')
    ttk.Label(headings,text='Марка / наименование / количество / масса',anchor='center').pack(side='left',fill='x',expand=True)
    ttk.Label(headings,text='Площадь, м² — 1 шт. / Всего',width=30,anchor='center').pack(side='right')
    grid=ttk.Treeview(frame,columns=tuple(range(7)),show='headings',height=20,selectmode='extended')
    y=ttk.Scrollbar(frame,orient='vertical',command=grid.yview);y.pack(side='right',fill='y');grid.configure(yscrollcommand=y.set)
    for i,h in enumerate(HEADERS):
        grid.heading(i,text=('1 шт.' if i==5 else 'Всего' if i==6 else h));grid.column(i,width=220 if i==1 else 135,minwidth=75,stretch=True)
    grid.pack(fill='both',expand=True);app.editable(grid)
    rows=getattr(app,'shipment_rows',None)
    if rows is None and app.vom.get():
        from .model import visible_vom
        try:rows=visible_vom(app.vom.get()).rows
        except Exception:rows=[]
    for row in rows or []:grid.insert('','end',values=tuple('' if v is None else str(v) for v in row))
    status=ttk.Label(window,text='');status.pack(fill='x',padx=12,pady=5)
    def commit_cell():
        pending=getattr(grid,'_pending_editor_commit',None)
        if callable(pending):pending()
    def paste(event=None):
        try:
            commit_cell()
            try:text=window.clipboard_get()
            except tk.TclError:
                from .storage import Stop
                raise Stop('В буфере нет таблицы. Выделите строки ВОМ в Excel, нажмите Ctrl+C и повторите Ctrl+V.')
            shipment=parse_tsv(text)
            replace_rows(grid,shipment.rows)
            status.configure(text=shipment.summary());grid.focus_set()
        except Exception as exc:app.error(exc)
        return 'break'
    def remove():
        commit_cell()
        for item in grid.selection():grid.delete(item)
    def accept():
        try:
            commit_cell();shipment=normalize([grid.item(i,'values') for i in grid.get_children()])
            target=local_root()/'vom-inputs'/uuid4().hex/'shipment.xlsx';write_internal(shipment,target)
            app.vom.set(str(target));app.shipment_rows=shipment.rows
            app.vom_summary.set(shipment.summary());app.write('Отгрузочная ВОМ: '+shipment.summary());window.destroy()
        except Exception as exc:app.error(exc)
    controls=ttk.Frame(window,padding=10);controls.pack(fill='x')
    ttk.Button(controls,text='Удалить строку',command=remove).pack(side='left',padx=6)
    ttk.Button(controls,text='Отмена',command=window.destroy).pack(side='right')
    ttk.Button(controls,text='ОК',command=accept).pack(side='right',padx=6)
    window.bind('<Control-v>',paste);window.bind('<Control-V>',paste)
    # On Windows the physical V key has keycode 86 in both English and Russian
    # layouts; the latter reports Cyrillic_em instead of the keysym "v".
    window.bind('<Control-KeyPress>',lambda e:paste(e) if e.keycode==86 else None)
    window.bind('<Escape>',lambda e:window.destroy())
    grid.focus_set();return window
