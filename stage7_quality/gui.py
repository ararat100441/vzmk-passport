"""No hidden regeneration: GUI consumes one explicitly selected frozen manifest."""
from datetime import datetime
from pathlib import Path
from uuid import uuid4
import tkinter as tk
from tkinter import filedialog, messagebox
from .core import verify_manifest, generate_quality
from .provenance import read, sha, Stage7Error
from .sources import ROOT


def launch():
    app=tk.Tk();app.title('ВЗМК — Этап 7 — DEV / TEST');app.geometry('950x680')
    state={'path':None,'sha':None,'output':None,'word':None}
    label=tk.StringVar(value='Выберите согласованный STAGE7_SELECTED_TEST_MANIFEST.json')
    tk.Label(app,textvariable=label,wraplength=920,justify='left').pack(fill='x',padx=12,pady=8)
    preview=tk.Text(app,wrap='word',font=('Segoe UI',11),height=25)
    preview.pack(fill='both',expand=True,padx=12)
    confirm=tk.BooleanVar(value=False)
    tk.Checkbutton(app,text='Подтверждаю формирование ТЕСТ / ЧЕРНОВИК. Фактический контроль не подтверждён.',variable=confirm).pack(pady=8)
    generate_button=tk.Button(app,text='Сформировать TEST DOCX',state='disabled')
    open_button=tk.Button(app,text='Открыть Word только для чтения',state='disabled')
    def choose():
        path=filedialog.askopenfilename(title='Точный выбранный manifest',filetypes=[('JSON','*.json')])
        if not path:return
        state.update(path=None,sha=None,output=None);confirm.set(False)
        generate_button.config(state='disabled');open_button.config(state='disabled')
        try:
            m=verify_manifest(read(path));state.update(path=path,sha=sha(path))
            f=m['fields'];lines=[f"Проект: {m['project']}",f"Паспорт: {m['passport']} — {f['end_date']}",
                f"Масса: {f['mass']} кг; 288 изделий; 104 видимые строки ВОМ.",
                'Состав выбранного TEST-комплекта:']
            lines += [r['display_title']+' — '+r['status'] for r in m['records']]
            lines += ['АКЗ №1/2/3 отсутствуют. АКЗ №0: только 3 назначенные марки / 3 изделия.',
                'УЗК SAMPLE_ONLY: ДНО / годен — пример заполнения.',
                'Результаты УЗК/АКЗ не подтверждены. Официальная выдача запрещена.',
                'Договор: '+f['contract'],'Заказчик: '+f['customer'],'Объект: '+f['object'],
                'КМД: '+f['cipher'],'Стали: '+f['steel'],
                'Пункт 12 сохраняется полностью; пункт 14 — короткий перечень.']
            preview.config(state='normal');preview.delete('1.0','end');preview.insert('1.0','\n\n'.join(lines));preview.config(state='disabled')
            label.set(path);generate_button.config(state='normal')
        except Exception as exc:messagebox.showerror('STOP',str(exc))
    def generate():
        try:
            if not confirm.get():raise Stage7Error('Подтвердите TEST-комплект после просмотра реквизитов.')
            if sha(state['path'])!=state['sha']:raise Stage7Error('Preview недействителен: manifest изменён. Выберите его заново.')
            output=ROOT/'DEV/Stage7_Выход'/(datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid4().hex[:8])
            reference=generate_quality(state['path'],output)
            state['output']=read(reference['path'])['output'];open_button.config(state='normal')
            messagebox.showinfo('TEST создан',state['output'])
        except Exception as exc:
            confirm.set(False);generate_button.config(state='disabled');messagebox.showerror('STOP',str(exc))
    def open_output():
        from .native import open_readonly
        try:state['word']=open_readonly(state['output'])
        except Exception as exc:messagebox.showerror('STOP',str(exc))
    tk.Button(app,text='Выбрать manifest и проверить',command=choose).pack(pady=5)
    generate_button.config(command=generate);generate_button.pack(pady=5)
    open_button.config(command=open_output);open_button.pack(pady=5)
    app.mainloop()
