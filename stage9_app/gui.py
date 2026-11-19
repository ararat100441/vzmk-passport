"""One responsive window; Office jobs run in an owned subprocess."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from uuid import uuid4
from .workspace import load,save,PROJECT

def config_path():
    return Path(os.environ.get('LOCALAPPDATA',str(Path.home())))/'VZMKPassport'/'settings.json'

class App:
    def __init__(self,window,workspace=None):
        self.window=window;self.workspace=Path(workspace) if workspace else None
        self.job=None;self.preview=None;self.manifest=None;self.close_plan=None;self.restore_plan=None;self.request=None
        self.office=[];self.busy=False
        window.title('ВЗМК — Паспортизация металлоконструкций');window.geometry('1000x760');window.minsize(720,520)
        window.protocol('WM_DELETE_WINDOW',self.shutdown)
        style=ttk.Style();style.configure('TButton',padding=(9,6));style.configure('Title.TLabel',font=('Segoe UI',16,'bold'))
        outer=ttk.Frame(window,padding=14);outer.pack(fill='both',expand=True)
        ttk.Label(outer,text='ВЗМК — Паспортизация металлоконструкций',style='Title.TLabel',wraplength=670).pack(anchor='w')
        from . import VERSION
        self.status=tk.StringVar(value=f'ТЕСТ · версия {VERSION} · официальная выдача запрещена')
        ttk.Label(outer,textvariable=self.status,wraplength=670).pack(anchor='w',pady=(6,10))
        tools=ttk.Frame(outer);tools.pack(fill='x')
        ttk.Button(tools,text='Создать DEV-стенд из источников',command=self.import_sources).pack(side='left')
        ttk.Button(tools,text='Открыть DEV-стенд',command=self.choose_workspace).pack(side='left',padx=6)
        self.workspace_label=ttk.Label(outer,text='Сначала создайте или откройте DEV-стенд.',wraplength=670);self.workspace_label.pack(anchor='w',pady=8)
        notebook=ttk.Notebook(outer);notebook.pack(fill='both',expand=True)
        doc=ttk.Frame(notebook,padding=10);notebook.add(doc,text='Паспорт и документы')
        cat=ttk.Frame(notebook,padding=10);notebook.add(cat,text='Справочники')
        close=ttk.Frame(notebook,padding=10);notebook.add(close,text='Закрытие и восстановление TEST')
        from .akz_auto_gui import AutoAKZPanel
        self.auto_akz_panel=AutoAKZPanel(self,notebook)
        doc.columnconfigure(1,weight=1);doc.rowconfigure(5,weight=1)
        self.project=tk.StringVar(value=PROJECT);self.number=tk.StringVar();self.vom=tk.StringVar();self.new=tk.BooleanVar(value=False)
        for rn,label,var in ((0,'Проект',self.project),(1,'Паспорт',self.number)):
            ttk.Label(doc,text=label).grid(row=rn,column=0,sticky='w',pady=4)
            box=ttk.Combobox(doc,textvariable=var,state='readonly');box.grid(row=rn,column=1,columnspan=2,sticky='ew',pady=4)
            if rn==0:self.project_box=box;box.bind('<<ComboboxSelected>>',self.refresh_numbers)
            else:self.number_box=box
        ttk.Checkbutton(doc,text='Новый паспорт (только в копии реестра)',variable=self.new,command=self.refresh_numbers).grid(row=2,column=0,columnspan=3,sticky='w',pady=4)
        ttk.Label(doc,text='Отгрузочная ВОМ').grid(row=3,column=0,sticky='w');ttk.Entry(doc,textvariable=self.vom).grid(row=3,column=1,sticky='ew')
        ttk.Button(doc,text='Выбрать',command=self.choose_vom).grid(row=3,column=2,padx=6)
        variants=ttk.Frame(doc);variants.grid(row=4,column=0,columnspan=3,sticky='w',pady=8)
        ttk.Label(variants,text='Варианты АКЗ:').pack(side='left');self.variants={n:tk.BooleanVar(value=n==0) for n in range(4)}
        for n,var in self.variants.items():ttk.Checkbutton(variants,text='№'+str(n),variable=var,command=self.invalidate).pack(side='left',padx=5)
        self.text=self.textbox(doc,5,3)
        actions=ttk.Frame(doc);actions.grid(row=6,column=0,columnspan=3,sticky='ew',pady=(8,0))
        ttk.Button(actions,text='1. Проверить и просмотреть',command=self.check).grid(row=0,column=0,padx=3,pady=3,sticky='ew')
        self.gen=ttk.Button(actions,text='2. Сформировать TEST-комплект',command=self.generate,state='disabled');self.gen.grid(row=0,column=1,padx=3,pady=3,sticky='ew')
        ttk.Button(actions,text='3. Открыть Word / Excel',command=self.open_documents).grid(row=1,column=0,padx=3,pady=3,sticky='ew')
        ttk.Button(actions,text='4. Проверить полноту',command=self.verify).grid(row=1,column=1,padx=3,pady=3,sticky='ew')
        ttk.Button(actions,text='Открыть ранее сформированный TEST-комплект',command=self.choose_manifest).grid(row=2,column=0,columnspan=2,padx=3,pady=3,sticky='ew')
        for c in range(2):actions.columnconfigure(c,weight=1)
        cat.columnconfigure(0,weight=1);cat.rowconfigure(2,weight=1)
        ttk.Label(cat,text='Принятые справочники копируются в DEV-стенд. Проверка и поиск не меняют источники.',wraplength=650).grid(row=0,column=0,sticky='w')
        self.search=tk.StringVar();search=ttk.Frame(cat);search.grid(row=1,column=0,sticky='ew',pady=8);search.columnconfigure(0,weight=1)
        ttk.Entry(search,textvariable=self.search).grid(row=0,column=0,sticky='ew');ttk.Button(search,text='Найти марку / профиль',command=self.catalog).grid(row=0,column=1,padx=6)
        self.cat_text=self.textbox(cat,2,1)
        close.columnconfigure(0,weight=1);close.rowconfigure(2,weight=1)
        ttk.Label(close,text='9/2026 и документы TEST/SAMPLE_ONLY не допускаются к официальному закрытию.\nДля проверки списания создайте отдельный искусственный паспорт 8001/2099.',wraplength=650).grid(row=0,column=0,sticky='w')
        top=ttk.Frame(close);top.grid(row=1,column=0,sticky='ew',pady=8)
        ttk.Button(top,text='Создать искусственный TEST-паспорт',command=self.fixture).grid(row=0,column=0,padx=3,pady=3)
        ttk.Button(top,text='Выбрать запрос TEST',command=self.choose_request).grid(row=0,column=1,padx=3,pady=3)
        self.closure_text=self.textbox(close,2,1)
        bottom=ttk.Frame(close);bottom.grid(row=3,column=0,sticky='ew',pady=(8,0))
        self.close_button=ttk.Button(bottom,text='Закрыть TEST-паспорт',command=self.close_test,state='disabled');self.close_button.grid(row=0,column=0,sticky='ew',padx=3,pady=3)
        self.recover_button=ttk.Button(bottom,text='Проверить незавершённую операцию',command=self.recover);self.recover_button.grid(row=0,column=1,sticky='ew',padx=3,pady=3)
        self.restore_button=ttk.Button(bottom,text='Восстановить из резервной копии',command=self.restore,state='disabled');self.restore_button.grid(row=1,column=0,columnspan=2,sticky='ew',padx=3,pady=3)
        for c in range(2):bottom.columnconfigure(c,weight=1)
        for var in (self.project,self.number,self.vom,self.new):var.trace_add('write',self.invalidate)
        if not self.workspace:
            try:self.workspace=Path(json.loads(config_path().read_text(encoding='utf8'))['workspace'])
            except (OSError,ValueError,KeyError):pass
        if self.workspace:
            try:self.activate(self.workspace)
            except Exception as exc:self.workspace=None;self.show(self.text,'ОШИБКА: '+str(exc))

    def textbox(self,parent,row,span):
        frame=ttk.Frame(parent);frame.grid(row=row,column=0,columnspan=span,sticky='nsew');frame.columnconfigure(0,weight=1);frame.rowconfigure(0,weight=1)
        text=tk.Text(frame,wrap='word',height=5,font=('Segoe UI',10),padx=8,pady=8,state='disabled')
        text.grid(row=0,column=0,sticky='nsew');bar=ttk.Scrollbar(frame,command=text.yview);bar.grid(row=0,column=1,sticky='ns');text.config(yscrollcommand=bar.set);return text
    def show(self,widget,value):widget.config(state='normal');widget.delete('1.0','end');widget.insert('1.0',value);widget.config(state='disabled')
    def invalidate(self,*args):self.preview=None;self.manifest=None;self.gen.config(state='disabled')
    def activate(self,root):
        root,_=load(root);self.workspace=root;self.project_box.config(values=[PROJECT]);self.project.set(PROJECT);self.refresh_numbers()
        self.workspace_label.config(text='DEV-стенд подключён. Источники защищены от записи.');self.status.set('ТЕСТ · ГОТОВО к проверке исходных данных')
        self.auto_akz_panel.refresh()
        from . import VERSION
        cfg=config_path();cfg.parent.mkdir(parents=True,exist_ok=True);tmp=cfg.with_name('settings-'+uuid4().hex+'.tmp');save(tmp,{'workspace':str(root),'version':VERSION});os.replace(tmp,cfg)
    def refresh_numbers(self,*args):
        if not self.workspace:return
        from .core import passports
        numbers,next_number=passports(self.workspace,self.project.get());values=[next_number] if self.new.get() else numbers
        self.number_box.config(values=values);self.number.set(next_number if self.new.get() else ('9/2026' if '9/2026' in values else values[-1]))
    def choose_workspace(self):
        if self.busy:return
        path=filedialog.askdirectory(title='Папка созданного DEV-стенда')
        if path:
            try:self.activate(path)
            except Exception as exc:messagebox.showerror('ОШИБКА',str(exc),parent=self.window)
    def import_sources(self):
        if self.busy:return
        source=filedialog.askdirectory(title='Исходная папка Pasport (только чтение)')
        if source:self.work('import',['--source',source,'--parent',str(config_path().parent/'workspaces')],lambda r:self.activate(r['workspace']),needs_workspace=False)
    def choose_vom(self):
        path=filedialog.askopenfilename(title='Отгрузочная ВОМ',filetypes=[('Excel','*.xls *.xlsx')])
        if path:self.vom.set(path)
    def work(self,task,args,callback,needs_workspace=True):
        if self.busy:return
        if needs_workspace and not self.workspace:messagebox.showerror('ТРЕБУЕТ ПРОВЕРКИ','Сначала создайте или откройте DEV-стенд.',parent=self.window);return
        folder=config_path().parent/'diagnostics';folder.mkdir(parents=True,exist_ok=True);result=folder/(uuid4().hex+'.json')
        cmd=[sys.executable] if getattr(sys,'frozen',False) else [sys.executable,'-B',str(Path(__file__).resolve().parents[1]/'stage9.py')]
        cmd += ['--task',task,'--result',str(result),*args]
        if needs_workspace:cmd += ['--workspace',str(self.workspace)]
        self.busy=True;self.status.set('ТРЕБУЕТ ПРОВЕРКИ · выполняется операция…')
        self.job=subprocess.Popen(cmd,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0),stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        def poll():
            if self.job.poll() is None:self.window.after(150,poll);return
            self.busy=False;self.job=None
            try:
                r=json.loads(result.read_text(encoding='utf8'))
                if r['status']!='PASS':raise ValueError(r['error'])
                self.status.set('ТЕСТ · операция завершена');callback(r)
            except Exception as exc:
                self.invalidate();self.status.set('ОШИБКА · операция остановлена');self.show(self.text,'ОШИБКА: '+str(exc));messagebox.showerror('ОШИБКА',str(exc),parent=self.window)
        self.window.after(150,poll)
    def check(self):
        args=['--project',self.project.get(),'--passport',self.number.get(),'--vom',self.vom.get(),'--variants',','.join(str(n) for n,v in self.variants.items() if v.get())]
        if self.new.get():args+=['--new']
        self.invalidate()
        selection=(self.project.get(),self.number.get(),self.vom.get(),self.new.get(),tuple(n for n,v in self.variants.items() if v.get()))
        def done(r):
            current=(self.project.get(),self.number.get(),self.vom.get(),self.new.get(),tuple(n for n,v in self.variants.items() if v.get()))
            if current!=selection:
                self.invalidate();self.show(self.text,'ТРЕБУЕТ ПРОВЕРКИ · выбор изменился во время операции. Повторите проверку.');return
            self.preview=r['preview'];self.show(self.text,r['summary']);self.gen.config(state='normal')
        self.work('check',args,done)
    def generate(self):
        if not self.preview:return
        if not messagebox.askyesno('Формирование TEST','Создать TEST-комплект по проверенным данным?\nФактические результаты контроля не подтверждены.',default='no',parent=self.window):return
        selection=(self.project.get(),self.number.get(),self.vom.get(),self.new.get(),tuple(n for n,v in self.variants.items() if v.get()))
        def done(r):
            current=(self.project.get(),self.number.get(),self.vom.get(),self.new.get(),tuple(n for n,v in self.variants.items() if v.get()))
            if current!=selection:
                self.invalidate();self.show(self.text,'ТРЕБУЕТ ПРОВЕРКИ · выбор изменился во время формирования. Комплект сохранён; откройте его через кнопку ранее сформированного TEST-комплекта.');return
            self.invalidate();self.manifest=r['manifest'];self.show(self.text,'ТЕСТ · комплект сформирован.\nОткройте документы в Word / Excel и проверьте полноту.\nОфициальная выдача запрещена.')
        self.work('generate',['--preview',self.preview,'--confirm-test'],done)
    def choose_manifest(self):
        if self.busy:return
        path=filedialog.askopenfilename(title='MANIFEST.json из предыдущего TEST-запуска',filetypes=[('JSON','*.json')])
        if not path:return
        try:
            from .core import verify_manifest
            m=verify_manifest(path)
            if self.workspace is None or not Path(path).resolve().is_relative_to(self.workspace.resolve()/'validation'):raise ValueError('Комплект должен принадлежать подключённому DEV-стенду.')
            self.invalidate();self.manifest=path
            self.show(self.text,f"ТЕСТ · открыт комплект {m['passport']}.\nКоличество: {m['quantity']} шт.; масса: {m['mass_kg']} кг.\nОфициальная выдача запрещена.")
        except Exception as exc:messagebox.showerror('ОШИБКА',str(exc),parent=self.window)
    def verify(self):
        if self.manifest:self.work('verify',['--manifest',self.manifest],lambda r:self.show(self.text,f"ГОТОВО · TEST-комплект проверен: {r['physical_files']} файлов, {r['document_records']} записей.\nОфициальное закрытие запрещено."))
    def open_documents(self):
        if self.busy or not self.manifest:return
        try:
            from .core import verify_manifest
            m=verify_manifest(self.manifest)
            from .office import Viewer
            if not self.close_viewers():return
            viewer=Viewer();viewer.open([r['path'] for r in m['records']]);self.office.append(viewer)
            self.status.set('ТЕСТ · документы открыты только для чтения')
        except Exception as exc:
            self.status.set('ОШИБКА · документы не открыты')
            messagebox.showerror('ОШИБКА',str(exc)+'\n'+'\n'.join(getattr(exc,'__notes__',[])),parent=self.window)
    def close_viewers(self):
        errors=[];remaining=[]
        for viewer in self.office:
            try:
                result=viewer.close()
                if result['status']!='PASS':
                    details='; '.join(e['step']+': '+e['error'] for e in result['cleanup_errors'])
                    errors.append(details+'\nДиагностика: '+result['diagnostic'])
            except Exception as exc:
                errors.append(str(exc))
                if not viewer.closed:remaining.append(viewer)
        self.office=remaining
        if errors:
            self.status.set('ОШИБКА · завершение Word / Excel требует проверки')
            messagebox.showerror('Ошибка закрытия документов',
                'При закрытии Word / Excel возникла ошибка. Новый просмотр не запущен.\n'+'\n'.join(errors),parent=self.window)
            return False
        return True
    def catalog(self):
        if not self.workspace or self.busy:return
        try:
            from .workspace import bind
            root,_=bind(self.workspace)
            from uzk_stage5.inputs import mark_profiles
            from profiles_stage4.storage import load_catalog
            query=self.search.get().strip();marks=mark_profiles(root/'documents/project/Справочник марок.xlsx')
            lines=[f"{r['mark']} — {r['profile']}" for r in marks.values() if query.casefold() in r['mark'].casefold()][:80]
            if query:
                found=load_catalog(root/'DEV/Справочники/Профили.xlsx').lookup(query)
                labels={'FOUND':'Найден','FOUND_PREFERRED':'Найден по принятому НД','NOT_FOUND':'Профиль не найден','AMBIGUOUS':'Несколько вариантов — требуется проверка','CONFLICT':'Конфликт справочника — требуется проверка'}
                lines.append('\nПрофиль: '+query+' — '+labels.get(found['status'],found['status']))
                for profile in found.get('candidates',[])[:10]:
                    lines.append(f"  {profile['family']} · {profile['standard']} · "+', '.join(f'{k}={v}' for k,v in profile['geometry'].items() if v is not None))
            self.show(self.cat_text,'\n'.join(lines) or 'Совпадений нет.')
        except Exception as exc:self.show(self.cat_text,'ОШИБКА: '+str(exc))
    def fixture(self):self.work('fixture',[],lambda r:self.load_request(r['request']))
    def choose_request(self):
        path=filedialog.askopenfilename(title='Точный CLOSE_REQUEST.json из TEST-стенда',filetypes=[('JSON','*.json')])
        if path:self.load_request(path)
    def load_request(self,path):
        self.request=Path(path);self.close_plan=None;self.restore_plan=None;self.close_button.config(state='disabled');self.restore_button.config(state='disabled')
        def done(r):
            self.close_plan=r['plan'];p=self.close_plan['payload']
            if p['status']=='ALREADY_CLOSED_TEST':self.show(self.closure_text,'ЗАКРЫТО · TEST. Повторное списание запрещено.');return
            from stage8_closure.core import describe
            self.show(self.closure_text,describe(self.close_plan));self.close_button.config(state='normal')
        self.work('close-preview',['--sandbox',str(self.request.parents[1]),'--request',str(path)],done)
    def plan_file(self,plan):
        path=config_path().parent/'diagnostics'/(uuid4().hex+'.json');save(path,plan);return str(path)
    def close_test(self):
        if not self.close_plan:return
        from stage8_closure.core import confirmation,describe
        if messagebox.askyesno('Закрыть искусственный TEST',describe(self.close_plan)+'\nПодтвердить однократное списание?',default='no',parent=self.window):
            self.work('close',['--preview',self.plan_file(self.close_plan),'--confirm',confirmation(self.close_plan)],lambda r:(self.close_button.config(state='disabled'),self.show(self.closure_text,'ЗАКРЫТО · TEST. Повторное списание запрещено.')))
    def recover(self):
        if not self.request:return
        def done(r):
            self.close_button.config(state='disabled');self.restore_button.config(state='disabled');self.close_plan=None
            if 'plan' in r:
                self.restore_plan=r['plan'];self.restore_button.config(state='normal')
                from stage8_closure.recovery import describe
                self.show(self.closure_text,describe(self.restore_plan))
            else:self.show(self.closure_text,'Журнал проверен.\n'+str(r.get('operation_status',r.get('state','ТЕСТ'))))
        self.work('recover',['--sandbox',str(self.request.parents[1])],done)
    def restore(self):
        if not self.restore_plan:return
        from stage8_closure.recovery import describe,confirmation
        if messagebox.askyesno('Восстановить TEST',describe(self.restore_plan)+'\nВосстановить из этой резервной копии?',default='no',parent=self.window):
            self.work('restore',['--preview',self.plan_file(self.restore_plan),'--confirm',confirmation(self.restore_plan)],lambda r:(self.restore_button.config(state='disabled'),self.show(self.closure_text,'ГОТОВО · восстановлено и проверено Excel.\nПаспорт не закрыт. Списаний: 0. Требуется новая проверка запроса.')))
    def shutdown(self):
        if self.busy:
            messagebox.showinfo('Операция выполняется','Дождитесь завершения безопасной операции перед закрытием программы.',parent=self.window);return
        if not self.close_viewers():return
        self.window.destroy()

def launch():
    w=tk.Tk();App(w);w.mainloop()
