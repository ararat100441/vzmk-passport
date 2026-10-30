"""Small DEV launcher. Generation requires a current in-memory verified preview."""
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from . import core, part2

class Launcher:
    def __init__(self, root):
        self.root=root; self.checked=None; self.output=None; self.outputs=[]; self.open_apps=[]
        root.title('ВЗМК — Этап 6 DEV — TEST АКЗ №0/1/2/3')
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
        ttk.Button(controls,text='Предварительная проверка №0/1/2/3',command=self.check).pack(side='left')
        self.generate_button=ttk.Button(controls,text='Сформировать выбранные TEST-пары',command=self.generate,state='disabled')
        self.generate_button.pack(side='left',padx=8)
        self.open_button=ttk.Button(controls,text='Открыть пару в Excel (read-only)',command=self.open,state='disabled')
        self.open_button.pack(side='left')
        groups=ttk.Frame(frame);groups.pack(fill='x')
        ttk.Label(groups,text='Создать группы:').pack(side='left')
        self.variant_vars={n:tk.BooleanVar(value=False) for n in range(4)}
        self.variant_buttons={}
        for n,var in self.variant_vars.items():
            button=ttk.Checkbutton(groups,text=f'№{n}',variable=var,state='disabled')
            button.pack(side='left');self.variant_buttons[n]=button
        self.result=tk.Text(frame,wrap='word');self.result.pack(fill='both',expand=True)
        for var in (self.project,self.passport,self.vom,self.catalog,self.synthetic):
            var.trace_add('write',self.invalidate)

    def invalidate(self,*args):
        self.checked=None; self.generate_button.configure(state='disabled')
        for button in self.variant_buttons.values():button.configure(state='disabled')

    def choose(self,var,is_catalog):
        path=filedialog.askopenfilename(parent=self.root,initialdir=str(core.FIXTURES if is_catalog else core.PROJECT/'Паспорт'),
            filetypes=[('Excel','*.xlsx')] if is_catalog else [('Excel','*.xls *.xlsx')])
        if path:var.set(path)

    def check(self):
        self.invalidate();self.result.delete('1.0','end')
        try:
            self.checked=part2.preview_groups(self.project.get(),self.passport.get(),self.vom.get(),self.catalog.get(),
                                             synthetic=self.synthetic.get())
            data=self.checked.payload
            lines=[f'Проект: {self.project.get()}    Паспорт: {self.passport.get()}',
                   f'Дата АКЗ из DEV: {data["date"]}. Климат: НЕТ ДАННЫХ / ТЕСТ. Официальная выдача STOP.']
            ready=False
            for n in range(4):
                group=data['groups'][str(n)];ok=group['status']=='READY'
                self.variant_vars[n].set(ok);self.variant_buttons[n].configure(state='normal' if ok else 'disabled')
                lines += ['',f'№ АКЗ={n}: {group["status"]}']
                if not ok:
                    lines.append(group['reason']);continue
                ready=True;g=group['payload'];t=g['templates'];totals=g['totals']
                lines += [f'Акт № {g["act_number"]}',t['act'],t['application']]
                if n==0:
                    lines.append(f'Слоёв: {t["layers"]}; {t["material"]}; {t["ral"]}; {t["required_dry_microns"]} мкм')
                else:
                    lines.append(f'Слоёв: {t["layer_count"]}')
                    for layer in t['layers']:
                        lines.append(f'  Слой {layer["ordinal"]}: {layer["material"]}; цвет {layer["color"]}; '
                                     f'{layer["required_raw"]} мкм; {layer["requirement_status"]}')
                for row in g['selected']:
                    lines.append(f'  ВОМ {row["row"]}: {row["mark"]} — {row["name"]}; '
                                 f'{row["quantity"]} шт.; {row["mass_total"]} кг; {row["area_total"]} м²')
                lines.append(f'Строк: {len(g["selected"])}; ИТОГО: {totals["quantity"]} шт.; '
                             f'{totals["mass_total"]} кг; {totals["area_total"]} м²')
            lines += ['',f'Без назначения: {len(data["unassigned"])} строк; blank не означает №0.']
            self.result.insert('end','\n'.join(lines))
            self.generate_button.configure(state='normal' if ready else 'disabled')
        except Exception as exc:
            self.result.insert('end','STOP: '+core.error_message(exc))

    def generate(self):
        try:
            if self.checked is None:raise core.Stage6Error('PREVIEW_REQUIRED')
            variants=[n for n,var in self.variant_vars.items() if var.get()]
            name='TEST_'+self.passport.get().replace('/','_')+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            results=part2.generate_groups(self.checked,variants,name)
            self.outputs=[Path(r['path']) for r in results['groups'].values() if r['status']=='COMPLETE_TEST_PAIR']
            self.output=self.outputs[0] if self.outputs else None
            self.open_button.configure(state='normal' if self.outputs else 'disabled');self.invalidate()
            self.result.insert('end','\n\n'+results['status'])
            for n,r in results['groups'].items():
                self.result.insert('end',f'\n№{n}: {r["status"]}\n'+r.get('path',r.get('reason','')))
            if results['status']!='COMPLETE_ALL_SELECTED':
                messagebox.showerror('STOP — не все пары созданы',
                    '\n\n'.join(f'№{n}: {r["reason"]}' for n,r in results['groups'].items() if r['status']=='STOP'),parent=self.root)
        except Exception as exc:
            self.invalidate();messagebox.showerror('STOP',core.error_message(exc),parent=self.root)

    def open(self):
        if self.output is None:return
        import win32com.client
        app=win32com.client.DispatchEx('Excel.Application')
        app.AutomationSecurity=3;app.EnableEvents=False;app.AskToUpdateLinks=False
        try:
            for output in self.outputs:
                for path in output.glob('*.xlsx'):
                    core.native_open(app,path,read_only=True)
            app.Visible=True
            self.open_apps.append(app)
        except Exception as exc:
            app.Quit();messagebox.showerror('STOP',core.error_message(exc),parent=self.root)

def main():
    root=tk.Tk();Launcher(root);root.mainloop()
