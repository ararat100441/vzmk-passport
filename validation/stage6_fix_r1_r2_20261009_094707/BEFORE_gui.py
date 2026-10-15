"""Small DEV launcher. Generation requires a current in-memory verified preview."""
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from . import core

class Launcher:
    def __init__(self, root):
        self.root=root; self.checked=None; self.output=None; self.open_apps=[]
        root.title('ВЗМК — Этап 6 DEV — только TEST АКЗ №0')
        root.geometry('1000x720')
        frame=ttk.Frame(root,padding=12);frame.pack(fill='both',expand=True)
        self.project=tk.StringVar(value='Гараж Казахстан')
        self.passport=tk.StringVar(value='9/2026')
        self.vom=tk.StringVar();self.catalog=tk.StringVar()
        ttk.Label(frame,text='ТЕСТ / ЧЕРНОВИК. Официальная выдача STOP. Назначения каталога только DEV.').pack(anchor='w')
        for label,var,values in [('Проект',self.project,['Гараж Казахстан']),
                ('Паспорт из DEV',self.passport,[r['number'] for r in core.passports()])]:
            row=ttk.Frame(frame);row.pack(fill='x',pady=5)
            ttk.Label(row,text=label,width=24).pack(side='left')
            ttk.Combobox(row,textvariable=var,values=values,state='readonly',width=40).pack(side='left')
        for label,var,is_catalog in [('Явная ВОМ XLS/XLSX',self.vom,False),('DEV-каталог назначений',self.catalog,True)]:
            row=ttk.Frame(frame);row.pack(fill='x',pady=5)
            ttk.Label(row,text=label,width=24).pack(side='left')
            ttk.Entry(row,textvariable=var).pack(side='left',fill='x',expand=True)
            ttk.Button(row,text='Выбрать',command=lambda v=var,c=is_catalog:self.choose(v,c)).pack(side='left')
        self.synthetic=tk.BooleanVar(value=False)
        ttk.Checkbutton(frame,text='SYNTHETIC isolated fixture (номер акта будет SYNTHETIC-TEST)',
                        variable=self.synthetic).pack(anchor='w')
        controls=ttk.Frame(frame);controls.pack(fill='x',pady=8)
        ttk.Button(controls,text='Предварительная проверка №0',command=self.check).pack(side='left')
        self.generate_button=ttk.Button(controls,text='Сформировать TEST-пару',command=self.generate,state='disabled')
        self.generate_button.pack(side='left',padx=8)
        self.open_button=ttk.Button(controls,text='Открыть пару в Excel (read-only)',command=self.open,state='disabled')
        self.open_button.pack(side='left')
        self.result=tk.Text(frame,wrap='word');self.result.pack(fill='both',expand=True)
        for var in (self.project,self.passport,self.vom,self.catalog,self.synthetic):
            var.trace_add('write',self.invalidate)

    def invalidate(self,*args):
        self.checked=None; self.generate_button.configure(state='disabled')

    def choose(self,var,is_catalog):
        path=filedialog.askopenfilename(parent=self.root,initialdir=str(core.FIXTURES if is_catalog else core.PROJECT/'Паспорт'),
            filetypes=[('Excel','*.xlsx')] if is_catalog else [('Excel','*.xls *.xlsx')])
        if path:var.set(path)

    def check(self):
        self.invalidate();self.result.delete('1.0','end')
        try:
            self.checked=core.preview(self.project.get(),self.passport.get(),self.vom.get(),self.catalog.get(),
                                      synthetic=self.synthetic.get())
            data=self.checked.payload
            totals=data['totals'];template=data['templates']
            lines=[f'Проект: {self.project.get()}    Паспорт: {self.passport.get()}',
                   f'Акт № {data["act_number"]}    Дата АКЗ из DEV: {data["date"]}',
                   f'ЛКМ: {template["material"]}; {template["ral"]}; требование {template["required_dry_microns"]} мкм.',
                   '', 'Выбранные строки АКЗ №0 (в порядке ВОМ):']
            for row in data['selected']:
                lines.append(f'  ВОМ {row["row"]}: {row["mark"]} — {row["name"]}; '
                             f'{row["quantity"]} шт.; {row["mass_total"]} кг; {row["area_total"]} м²')
            lines += ['',f'ИТОГО: {totals["quantity"]} шт.; {totals["mass_total"]} кг; {totals["area_total"]} м²',
                      f'Видимых строк ВОМ: {len(data["vom"]["rows"])}; скрытых исключено: {len(data["vom"]["hidden_rows"])}.',
                      '', *data['warnings'], '', 'Точная пара шаблонов:',template['act'],template['application'],
                      '', 'Не включены в АКЗ №0:']
            for row in data['excluded']:
                reason='нет назначения' if row['assignment'] is None else f'№ АКЗ={row["assignment"]} (не реализован)'
                lines.append(f'  ВОМ {row["row"]}: {row["mark"]} — {reason}')
            self.result.insert('end','\n'.join(lines))
            self.generate_button.configure(state='normal')
        except Exception as exc:
            self.result.insert('end','STOP: '+str(exc))

    def generate(self):
        try:
            if self.checked is None:raise core.Stage6Error('PREVIEW_REQUIRED')
            name='TEST_'+self.passport.get().replace('/','_')+'_'+datetime.now().strftime('%Y%m%d_%H%M%S')
            self.output=core.generate(self.checked,name)
            self.open_button.configure(state='normal');self.invalidate()
            self.result.insert('end','\n\nCOMPLETE TEST PAIR: '+str(self.output))
        except Exception as exc:
            self.invalidate();messagebox.showerror('STOP',str(exc),parent=self.root)

    def open(self):
        if self.output is None:return
        import win32com.client
        app=win32com.client.DispatchEx('Excel.Application')
        app.AutomationSecurity=3;app.EnableEvents=False;app.AskToUpdateLinks=False
        try:
            for path in self.output.glob('*.xlsx'):
                core.native_open(app,path,read_only=True)
            app.Visible=True
            self.open_apps.append(app)
        except Exception as exc:
            app.Quit();messagebox.showerror('STOP',str(exc),parent=self.root)

def main():
    root=tk.Tk();Launcher(root);root.mainloop()
