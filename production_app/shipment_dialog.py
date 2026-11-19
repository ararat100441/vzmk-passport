"""One compact seven-column editor using the accepted RC3 Treeview editor."""
import tkinter as tk
from tkinter import messagebox,ttk
from uuid import uuid4
from .storage import local_root
from .shipment_input import HEADERS,normalize,numeric,parse_tsv,write_internal
from .storage import Stop

def replace_rows(grid,rows):
    """Ctrl+V is a complete shipment replacement, independent of selection."""
    for item in grid.get_children():grid.delete(item)
    for row in rows:
        grid.insert('','end',values=tuple('' if v is None else str(v) for v in row))

def open_shipment(app,*,original_rows=None,original_digest=None,draft_rows=None,draft_indices=None,on_accept=None):
    if app.active:return
    window=tk.Toplevel(app.root);window.title('Отгрузочная ВОМ');app.fit(window,1180,680)
    window.transient(app.root);window.grab_set()
    # Reserve the action bar BEFORE the expanding grid. Packing it last let
    # the 20-row Treeview consume the available height at Windows DPI >100%.
    controls=ttk.Frame(window,padding=(10,7))
    controls.pack(side='bottom',fill='x')
    status=ttk.Label(window,text='')
    status.pack(side='bottom',fill='x',padx=12,pady=(2,4))
    instruction=('Ctrl+V добавляет строки. Первоначальные строки можно удалить по одной.'
                 if original_rows is not None else
                 'Вставьте строки из Excel: Ctrl+V. Двойной щелчок — изменить ячейку. Площадь можно оставить пустой.')
    ttk.Label(window,text=instruction,wraplength=1120).pack(anchor='w',padx=12,pady=8)
    frame=ttk.Frame(window,padding=10);frame.pack(fill='both',expand=True)
    headings=ttk.Frame(frame);headings.pack(fill='x')
    ttk.Label(headings,text='Марка / наименование / количество / масса',anchor='center').pack(side='left',fill='x',expand=True)
    ttk.Label(headings,text='Площадь, м² — 1 шт. / Всего',width=30,anchor='center').pack(side='right')
    grid=ttk.Treeview(frame,columns=tuple(range(7)),show='headings',height=7,selectmode='extended')
    y=ttk.Scrollbar(frame,orient='vertical',command=grid.yview);y.pack(side='right',fill='y');grid.configure(yscrollcommand=y.set)
    for i,h in enumerate(HEADERS):
        grid.heading(i,text=('1 шт.' if i==5 else 'Всего' if i==6 else h));grid.column(i,width=220 if i==1 else 135,minwidth=75,stretch=True)
    grid.pack(fill='both',expand=True)
    def table_rows():
        return [grid.item(item,'values') for item in grid.get_children()]
    def update_status():
        # Never display the totals of an older selection after a row edit.
        # Invalid/incomplete edits must not be shown as a valid shipment.
        rows=table_rows()
        if not rows:
            status.configure(text='ВОМ пуста. Добавьте марку или вставьте строки.')
            return
        try:
            status.configure(text=normalize(rows).summary())
        except Exception as exc:
            status.configure(text='ВОМ требует исправления: '+str(exc))
    origins={}
    if original_rows is None:app.editable(grid,on_change=update_status)
    else:app.editable(grid,on_change=update_status,can_edit=lambda item,col:item not in origins)
    rows=original_rows if original_rows is not None else getattr(app,'shipment_rows',None)
    if draft_rows is not None:rows=draft_rows
    indices=draft_indices if draft_rows is not None else range(len(original_rows or ()))
    if rows is None and app.vom.get():
        from .model import visible_vom
        try:rows=visible_vom(app.vom.get()).rows
        except Exception:rows=[]
    for pos,row in enumerate(rows or []):
        item=grid.insert('','end',values=tuple('' if v is None else str(v) for v in row))
        if original_rows is not None and indices[pos] is not None:origins[item]=indices[pos]
    update_status()
    def commit_cell():
        pending=getattr(grid,'_pending_editor_commit',None)
        if callable(pending):pending()
    def paste(event=None):
        try:
            commit_cell()
            try:text=window.clipboard_get()
            except tk.TclError:
                raise Stop('В буфере нет таблицы. Выделите строки ВОМ в Excel, нажмите Ctrl+C и повторите Ctrl+V.')
            shipment=parse_tsv(text)
            if original_rows is None:replace_rows(grid,shipment.rows)
            else:
                for row in shipment.rows:grid.insert('','end',values=tuple('' if v is None else str(v) for v in row))
            update_status();grid.focus_set()
        except Exception as exc:app.error(exc)
        return 'break'
    def add():
        commit_cell()
        form=tk.Toplevel(window);form.title('Добавить марку')
        form.transient(window);form.resizable(False,False)
        fields=[]
        for line,col in enumerate((0,1,2,3,5)):
            ttk.Label(form,text=HEADERS[col]).grid(row=line,column=0,sticky='w',padx=10,pady=4)
            entry=ttk.Entry(form,width=30);entry.grid(row=line,column=1,padx=10,pady=4)
            fields.append(entry)

        def close(event=None):
            form.destroy();window.grab_set();grid.focus_set()

        def insert(event=None):
            try:
                mark,name,q,m,a=(entry.get() for entry in fields)
                q,m,a=(numeric(value,HEADERS[col]) for col,value in ((2,q),(3,m),(5,a)))
                row=normalize([(mark,name,q,m,q*m,a,q*a)]).rows[0]
            except Stop as exc:
                messagebox.showerror('Проверьте данные марки',str(exc),parent=form)
                return
            item=grid.insert('','end',values=tuple(map(str,row)))
            grid.selection_set(item);grid.see(item)
            update_status();close()

        buttons=ttk.Frame(form);buttons.grid(row=5,column=0,columnspan=2,sticky='e',padx=10,pady=8)
        ttk.Button(buttons,text='Добавить',command=insert).pack(side='left',padx=(0,8))
        ttk.Button(buttons,text='Отмена',command=close).pack(side='left')
        form.protocol('WM_DELETE_WINDOW',close)
        form.bind('<Return>',insert);form.bind('<Escape>',close)
        form.wait_visibility();form.grab_set();fields[0].focus_set()

    def remove():
        commit_cell()
        selected=grid.selection()
        if len(selected)!=1:
            messagebox.showinfo('Удалить марку','Выберите одну строку ВОМ.',parent=window)
            return
        item=selected[0]
        mark,_,count,*_=grid.item(item,'values')
        if messagebox.askyesno('Удалить марку',f'Удалить выбранную строку {mark}, {count} шт.?',parent=window):
            grid.delete(item);origins.pop(item,None);update_status()
    def accept():
        try:
            commit_cell();shipment=normalize(table_rows())
            target=local_root()/'vom-inputs'/uuid4().hex/'shipment.xlsx';write_internal(shipment,target)
            if on_accept:
                on_accept({'appendix_sha256':original_digest,
                           'indices':[origins.get(item) for item in grid.get_children()]})
            app.vom.set(str(target));app.shipment_rows=shipment.rows
            app.vom_summary.set(shipment.summary());app.write('Отгрузочная ВОМ: '+shipment.summary());window.destroy()
        except Exception as exc:app.error(exc)
    ttk.Button(controls,text='Добавить марку',command=add).pack(side='left',padx=6)
    ttk.Button(controls,text='Удалить марку',command=remove).pack(side='left',padx=6)
    ttk.Button(controls,text='Отмена',command=window.destroy).pack(side='right')
    ttk.Button(controls,text='ОК',command=accept).pack(side='right',padx=6)
    window.bind('<Control-v>',paste);window.bind('<Control-V>',paste)
    # On Windows the physical V key has keycode 86 in both English and Russian
    # layouts; the latter reports Cyrillic_em instead of the keysym "v".
    window.bind('<Control-KeyPress>',lambda e:paste(e) if e.keycode==86 else None)
    window.bind('<Escape>',lambda e:window.destroy())
    grid.focus_set();return window
