import os,subprocess,sys,traceback
from pathlib import Path
from uuid import uuid4
from datetime import date, datetime
from copy import deepcopy
import tkinter as tk
from tkinter import ttk,filedialog,messagebox,simpledialog
from .storage import local_root,read,save,sha,Stop
from . import model,core,akz_presets,VERSION
from .external_open import open_external
from .owner_control import STATEMENT
from .real_forms import DOC_FIELD_LABELS, UZK_COLUMNS, UZK_HEADERS

class App:
    def __init__(self,root):
        self.root=root;root.title('ВЗМК — Паспортизация '+VERSION);self.fit(root,1050,760);root.minsize(min(950,root.winfo_screenwidth()-40),min(650,root.winfo_screenheight()-80))
        self.config_path=local_root()/'rc3-marked-settings.json'
        initial=Path(sys.executable).parent/'initial_settings.json'
        config_error=None
        try:self.cfg=read(self.config_path) if self.config_path.exists() else read(initial) if getattr(sys,'frozen',False) and initial.exists() else {'profiles':{}}
        except Exception as exc:self.cfg={'profiles':{}};config_error=str(exc)
        if getattr(sys,'frozen',False):
            from .path_config import installed_assets
            self.cfg=installed_assets(self.cfg,Path(sys.executable).parent)
        self.plan=None;self.facts={};self.active=None;self.result_path=None;self.buttons=[]
        self.last_output_folder=None
        self.last_output_project=None
        self.preview_pdf=None
        self.first_reference=None  # read-only first issued passport, never copied as facts
        self.project=tk.StringVar();self.number=tk.StringVar();self.vom=tk.StringVar();self.new=tk.BooleanVar(value=True)
        self.shipment_rows=None;self.preliminary=None
        self.vom_summary=tk.StringVar(value='Строки ВОМ ещё не введены')
        self.source_number=tk.StringVar()
        self.shipment_date=tk.StringVar(value=date.today().strftime('%d.%m.%Y'))
        self.shipment_date.trace_add('write',lambda *_:self.invalidate())
        self.project.trace_add('write',lambda *_:self.inputs_changed())
        self.number.trace_add('write',lambda *_:self.inputs_changed())
        self.vom.trace_add('write',lambda *_:self.shipment_changed())
        top=ttk.Frame(root,padding=12);top.pack(fill='x')
        ttk.Label(top,text='Паспорта металлоконструкций',font=('Segoe UI',17,'bold')).pack(side='left')
        ttk.Button(top,text='Настройки путей',command=self.settings).pack(side='right')
        self.paths=ttk.Label(root,wraplength=1050,justify='left');self.paths.pack(fill='x',padx=14)
        form=ttk.Frame(root,padding=14);form.pack(fill='x');form.columnconfigure(1,weight=1)
        ttk.Label(form,text='Объект из реестра').grid(row=0,column=0,sticky='w')
        self.projects=ttk.Combobox(form,textvariable=self.project,state='readonly');self.projects.grid(row=0,column=1,sticky='ew',padx=8,pady=5)
        self.projects.bind('<<ComboboxSelected>>',lambda e:self.select_project())
        ttk.Button(form,text='Обновить реестр',command=self.refresh).grid(row=0,column=2)
        ttk.Label(form,text='№ паспорта, исходная серия').grid(row=1,column=0,sticky='w')
        self.numbers=ttk.Combobox(form,textvariable=self.number,state='readonly');self.numbers.grid(row=1,column=1,sticky='ew',padx=8,pady=5)
        self.new_widget=ttk.Checkbutton(form,text='Новый паспорт',variable=self.new,command=self.select_project);self.new_widget.grid(row=1,column=2)
        self.vom_widget=ttk.Label(form,textvariable=self.vom_summary,wraplength=620)
        self.vom_widget.grid(row=2,column=1,sticky='ew',padx=8,pady=5)
        ttk.Button(form,text='Отгрузочная ВОМ',command=self.shipment_dialog).grid(row=2,column=0,pady=5,sticky='w')
        ttk.Label(form,text='Исходный паспорт (образец документов)').grid(row=3,column=0,sticky='w')
        self.source_numbers=ttk.Combobox(form,textvariable=self.source_number,state='readonly')
        self.source_numbers.grid(row=3,column=1,sticky='ew',padx=8,pady=5)
        self.source_numbers.bind('<<ComboboxSelected>>',lambda e:self.select_source())
        self.shipment_label=ttk.Label(form,text='Дата окончания / отгрузки')
        self.shipment_label.grid(row=4,column=0,sticky='w')
        self.shipment_entry=ttk.Entry(form,textvariable=self.shipment_date,width=20)
        self.shipment_entry.grid(row=4,column=1,sticky='w',padx=8,pady=5)
        tools=ttk.Frame(root,padding=(14,0));tools.pack(fill='x')
        # The first issued passport supplies the visual reference. The clean
        # approved forms remain an internal software asset, not a daily dialog.
        for text,command in [('Фактические результаты',self.fact_dialog),('Открыть документы',self.open_documents)]:
            ttk.Button(tools,text=text,command=command).pack(side='left',padx=(0,6),pady=5)
        steps=ttk.Frame(root,padding=14);steps.pack(fill='x')
        # A single user action; native Office/registry safety stages remain.
        button=ttk.Button(steps,text='Сформировать пакет',command=self.check)
        button.pack(fill='x',expand=True)
        self.buttons.append(button)
        self.progress=ttk.Label(root,text='Выберите рабочие пути. Реестр пока не изменяется.');self.progress.pack(fill='x',padx=15)
        ttk.Label(root,text='Запись отгрузки выполняется только после просмотра PDF и вашего подтверждения.',wraplength=940).pack(side='bottom',padx=14,pady=(0,12))
        self.log=tk.Text(root,wrap='word',font=('Segoe UI',10),state='disabled',height=10);self.log.pack(fill='both',expand=True,padx=14,pady=12)
        root.protocol('WM_DELETE_WINDOW',self.close)
        self.invalidate();self.refresh()
        configured=list(self.cfg.get('profiles',{}));name=self.cfg.get('last_project') or (configured[0] if len(configured)==1 else '')
        if name in self.projects['values']:self.project.set(name);self.select_project()
        if not self.cfg.get('registry'):root.after(300,self.settings)
        if config_error:root.after(500,lambda:self.error('Не удалось загрузить настройки. Выберите пути заново. Исходники не изменены. '+config_error))

    def fit(self,window,width,height):window.geometry(f'{min(width,window.winfo_screenwidth()-40)}x{min(height,window.winfo_screenheight()-80)}')

    def scroll_content(self,dialog):
        outer=ttk.Frame(dialog);outer.pack(fill='both',expand=True)
        canvas=tk.Canvas(outer,highlightthickness=0)
        sy=ttk.Scrollbar(outer,orient='vertical',command=canvas.yview)
        sx=ttk.Scrollbar(outer,orient='horizontal',command=canvas.xview)
        canvas.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
        sy.pack(side='right',fill='y');sx.pack(side='bottom',fill='x');canvas.pack(fill='both',expand=True)
        body=ttk.Frame(canvas);canvas.create_window((0,0),window=body,anchor='nw')
        body.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')))
        return body

    def close(self):
        if self.active and self.active.poll() is None:
            messagebox.showinfo('Операция Office','Дождитесь штатного завершения Office. Программа не будет принудительно закрывать процессы.');return
        self.root.destroy()

    def write(self,text):
        self.log.configure(state='normal');self.log.insert('end',str(text)+'\n');self.log.see('end');self.log.configure(state='disabled')

    def invalidate(self):
        self.plan=None;self.preview_pdf=None;self.preliminary=None
        for b in self.buttons:
            b.configure(state='normal' if not self.active else 'disabled',text='Сформировать пакет')

    def inputs_changed(self):
        self.facts={};self.invalidate()

    def synchronized_facts(self):
        from .shared_uzk import find_shared_uzk,planned_control
        from .catalog_assignments import current_assignments
        from .shipment_sync import synchronize
        from .geometry_selection import GeometrySelection
        from .automatic_acts import find_approved_geometry
        v=model.visible_vom(self.vom.get())
        source=find_shared_uzk(self.cfg,require=True)
        plan=planned_control(source,self.project.get(),v)
        profile=model.profile(self.cfg,self.project.get());assignments={}
        if profile.get('catalog'):
            assignments=current_assignments(profile['catalog'],profile['catalog_sha256'],plan)
        data=synchronize(self.facts,self.project.get(),self.number.get(),self.cfg.get('training',False),plan,assignments)
        geometry=find_approved_geometry(self.cfg,require=False)
        if geometry:
            selection=GeometrySelection(geometry,self.cfg,self.project.get())
            for row in data['marks']:row['standard']=selection.resolve(row['mark'],row['profile'])
        self.cfg['uzk_registry']=str(source);self.cfg['use_shared_uzk']=True
        return data

    def shipment_changed(self):
        # Changing only VOM preserves entered climate ranges, while clearing
        # approvals and replacing mark identities/quantities from real sources.
        self.invalidate()
        if not self.project.get() or not self.vom.get():return
        try:self.facts=self.synchronized_facts()
        except Exception as exc:
            # Do not keep a stale mark list when a source cannot be read.
            self.facts={'akz_drafts':deepcopy(self.facts.get('akz_drafts',{}))}
            self.write('Фактические результаты требуют проверки: '+str(exc))

    def error(self,exc):self.write(str(exc));messagebox.showerror('Не удалось продолжить',str(exc))

    def choose(self,var,directory):
        if self.active:return
        path=filedialog.askdirectory() if directory else filedialog.askopenfilename(filetypes=[('Excel/Word','*.xlsx *.xlsm *.docx'),('Все файлы','*.*')])
        if path:var.set(path)

    def settings_path(self):return getattr(self,'config_path',local_root()/'settings.json')

    def shipment_dialog(self):
        from .shipment_dialog import open_shipment
        return open_shipment(self)

    def choose_vom(self):
        self.choose(self.vom,False)
        if self.vom.get():self.summarize_vom()

    def summarize_vom(self):
        try:
            v=model.visible_vom(self.vom.get())
            self.write(f'ВОМ: {len(v.rows)} марок / {v.quantity} изделий / масса {v.mass:.2f} кг / площадь {str(v.area) if v.area is not None else 'не указана'} м². Исходный файл не изменён.')
            self.progress.configure(text='ВОМ прочитана. Для выпуска внесите текущие результаты в «Фактические результаты».')
            return v
        except Exception as exc:self.error(exc)

    def settings(self):
        if self.active:return
        from .path_config import visible_settings,apply_settings
        shown=visible_settings(self.cfg)
        dialog=tk.Toplevel(self.root)
        dialog.title('Рабочие пути и справочники')
        self.fit(dialog,1100,740)
        dialog.transient(self.root)
        dialog.grab_set()
        selected={key:tk.StringVar(value=value)
                  for key,value in shown.items()}
        dialog._selected_path_vars=selected
        inputs=[
            ('registry','Рабочий реестр ВЗМК.xlsx'),
            ('output','Папка рабочих документов (первые и новые паспорта)'),
            ('uzk_registry','ОБЩИЙ реестр УЗК.xlsm — для всех проектов'),
        ]
        for row,(key,label) in enumerate(inputs):
            ttk.Label(dialog,text=label).grid(
                row=row,column=0,padx=12,pady=12,sticky='w')
            ttk.Entry(dialog,textvariable=selected[key],width=84).grid(
                row=row,column=1,sticky='ew')
            ttk.Button(dialog,text='Выбрать…',
                       command=lambda k=key:self.choose(
                           selected[k],k=='output')).grid(
                           row=row,column=2,padx=8)
        ttk.Label(
            dialog,text='Программа сама находит первый паспорт в '
            'папке проекта / Паспорт и сохраняет новый там же, '
            'отдельно по номеру и дате. Все документы берутся из явно выбранного исходного паспорта без скрытой подмены. До подтверждения ничего не списывается; выпуск обновляет остатки с резервной копией.',
            wraplength=1010,
        ).grid(row=3,column=0,columnspan=3,padx=14,pady=8)

        effective={}
        try:
            effective=model.profile(self.cfg,self.project.get()) if self.project.get() else {}
            from .automatic_acts import find_approved_geometry
            source_paths=[('Справочник марок выбранного объекта',effective.get('catalog')),
                          ('Геометрия Профили.xlsx',find_approved_geometry(self.cfg,require=False))]
        except Stop as exc:
            source_paths=[('Справочники',None)]
            self.write(str(exc))
        dialog._source_path_vars=[]
        for row,(label,path) in enumerate(source_paths,4):
            ttk.Label(dialog,text=label).grid(row=row,column=0,padx=12,pady=8,sticky='w')
            value=tk.StringVar(value=str(path or 'не найден'))
            dialog._source_path_vars.append(value)
            ttk.Entry(dialog,textvariable=value,state='readonly',width=84).grid(row=row,column=1,sticky='ew')
            ttk.Button(dialog,text='Открыть',state='normal' if path else 'disabled',
                       command=lambda p=path:open_external(str(p))).grid(row=row,column=2,padx=8)
        def migrate_catalog():
            try:
                path=effective.get('catalog')
                if not path:raise Stop('Справочник проекта не найден.')
                from .catalog_ledger import migrate_in_place
                if not messagebox.askyesno('Справочник — резервная копия',
                    'Подготовить новую структуру справочника с резервной копией? Общее количество заполняется вручную. Неполная история не даёт подтверждённого остатка.\n'+str(path)):return
                result=migrate_in_place(path,self.cfg,self.project.get())
                self.write(str(result));messagebox.showinfo('Справочник',str(result))
            except Exception as exc:self.error(exc)
        ttk.Button(dialog,text='Подготовить структуру справочника',command=migrate_catalog).grid(row=7,column=1,sticky='w',pady=6)

        def choose_current_catalog():
            try:
                project=self.project.get()
                if not project:raise Stop('Выберите зарегистрированный объект на главном экране.')
                path=filedialog.askopenfilename(parent=dialog,title='Действующий справочник выбранного проекта',filetypes=[('Справочник XLSX','*.xlsx')])
                if not path:return
                from .project_catalog import resolve_catalog
                # Resolve the project root without substituting another project.
                from .first_passport import archive_root_for_project
                archive=archive_root_for_project(selected['output'].get(),project)
                roots=[x for x in archive.iterdir() if x.is_dir() and x.name.casefold()==project.casefold()]
                if len(roots)!=1:raise Stop('Не найдена единственная существующая папка выбранного проекта.')
                active=resolve_catalog(roots[0],project,path)
                self.cfg.setdefault('catalog_choices',{})[project]=str(active)
                save(self.settings_path(),self.cfg);self.invalidate()
                effective.update(model.profile(self.cfg,project))
                dialog._source_path_vars[0].set(str(active))
                self.write('Действующий справочник: '+str(active))
            except Exception as exc:self.error(exc)
        def create_project_catalog():
            try:
                folder=filedialog.askdirectory(parent=dialog,title='Уже существующая папка проекта',initialdir=selected['output'].get())
                if not folder:return
                folder=Path(folder).resolve(strict=True)
                project=self.project.get() if folder.name.casefold()==self.project.get().casefold() else folder.name
                from .project_catalog import create_catalog
                result=create_catalog(folder,project,self.cfg)
                self.write('Создан только справочник: '+result['path'])
                messagebox.showinfo('Справочник создан',result['path']+'\n'+result['notice'])
            except Exception as exc:self.error(exc)
        ttk.Button(dialog,text='Создать справочник в существующей папке проекта',command=create_project_catalog).grid(row=10,column=1,sticky='w',pady=5)
        ttk.Button(dialog,text='Выбрать действующий справочник',command=choose_current_catalog).grid(row=11,column=1,sticky='w',pady=5)

        def store(training=False):
            try:
                if (self.cfg.get('training') and not training and
                        selected['registry'].get()==self.cfg.get('registry')):
                    raise Stop('Для выхода из учебного режима явно выберите '
                               'рабочий реестр ВЗМК, а не текущую копию.')
                installed=Path(sys.executable).resolve().parent if getattr(
                    sys,'frozen',False) else None
                cfg=apply_settings(
                    self.cfg,
                    registry=selected['registry'].get(),
                    output=selected['output'].get(),
                    uzk_registry=selected['uzk_registry'].get(),
                    installed_root=installed,
                )
                model.validate_paths(cfg)
                if training:
                    parent=filedialog.askdirectory(
                        title='Куда создать НОВУЮ учебную копию?')
                    if not parent:return
                    cfg=core.training_copy(cfg,parent)
                retained=self.facts if all(cfg.get(k)==self.cfg.get(k) for k in ('registry','output','uzk_registry')) else {}
                self.cfg=cfg
                save(self.settings_path(),cfg)
                self.invalidate()
                dialog.destroy();self.refresh();self.facts=retained
            except Exception as exc:self.error(exc)

        ttk.Button(dialog,text='Сохранить рабочие пути',
                   command=store).grid(row=8,column=1,sticky='w',pady=12)
        # Training-copy machinery remains available to automated tests, but
        # it is not a daily work button in the simplified PTO interface.
        def open_recovery():
            dialog.destroy()
            self.recovery()
        ttk.Button(dialog,text='Восстановление после сбоя',
                   command=open_recovery).grid(
                   row=9,column=1,sticky='w',pady=6)

    def show_paths(self):
        from .shared_uzk import find_shared_uzk
        try:uzk=find_shared_uzk(self.cfg,require=False)
        except Stop:uzk=None
        self.paths.configure(
            text='Реестр: '+str(self.cfg.get('registry') or 'не выбран')+'\n'
            'Документы и паспорта: '+str(self.cfg.get('output') or 'не выбрана')+'\n'
            'Общий УЗК: '+str(uzk or 'не выбран'),
        )

    def refresh(self):
        if self.active:return
        self.show_paths()
        if not self.cfg.get('registry'):return
        try:
            r=model.read_registry(self.cfg['registry'])
            self.projects['values']=sorted(r.projects)
            self.invalidate()
            self.progress.configure(
                text='Реестр прочитан. Выберите проект; запись не производилась.')
            if self.project.get():self.select_project()
        except Exception as exc:self.error(exc)

    def select_project(self):
        if self.active:return
        self.invalidate();self.facts={};self.first_reference=None;self.number.set('')
        if self.last_output_project!=self.project.get():self.last_output_folder=None
        try:
            r=model.read_registry(self.cfg['registry']);name=self.project.get();entries=r.projects[name]
            self.numbers['values']=[str(e['number']) for e in entries]
            self.source_numbers['values']=list(self.numbers['values'])
            self.number.set(str(entries[0]['number']))
            p=model.select(r,name)
            self.source_number.set((self.cfg.get('source_passports') or {}).get(name,p.first_number))
            self.shipment_entry.configure(state='normal' if self.new.get() else 'disabled')
            if self.new.get():self.shipment_label.grid();self.shipment_entry.grid()
            else:self.shipment_label.grid_remove();self.shipment_entry.grid_remove()
            from .path_config import normalize_project_roots
            normalized=normalize_project_roots(self.cfg,name)
            if normalized!=self.cfg:
                self.cfg=normalized
                self.show_paths()
                self.write('Папка источников проекта уточнена по структуре '
                           '«объект/Паспорт»: '+
                           str(self.cfg['first_passport_documents'])+
                           '. Исходные документы и реестр не изменены.')
            # The FIRST ISSUED passport is the automatic reference for THIS
            # project only. Existing approved empty forms remain separate:
            # never inject historical findings or signatures as new facts.
            from .first_passport import discover_first_passport_for_config
            self.cfg['auto_first_passport']=True
            try:
                self.first_reference=discover_first_passport_for_config(
                    self.cfg,p)
                refs=self.first_reference['sources']
                source_num=self.first_reference.get('first_number',p.first_number)
                self.write(('Автоматически найден первый паспорт: ' if source_num==p.first_number else 'Автоматически найден образец документов: ')+
                           source_num+' / '+
                           self.first_reference['passport_root'])
                for role in ('appendix','vik','uzk','quality','akz0','akz_appendix0'):
                    if role in refs:
                        self.write('Источник '+role+': '+refs[role]['path'])
                if self.first_reference['missing']:
                    self.write('Не найдены источники первого паспорта: '+
                               ', '.join(self.first_reference['missing'])+
                               '. Другие проекты не подставляются.')
                else:
                    self.write('Источники первого паспорта найдены. Первое '
                               'Приложение №1 будет использовано напрямую; '
                               'старые акты контроля остаются только образцами.')
            except Stop as exc:
                self.write('Автопоиск первого паспорта: '+str(exc)+
                           ' Шаблоны других объектов не подставляются.')
            if self.new.get():self.numbers['values']=[p.next_number];self.number.set(p.next_number)
            self.write(f'{name}: первый {p.first_number}; последний {p.last_number}; следующий {p.next_number}')
            aliases=model.approved_legacy_aliases(name,entries)
            if aliases:
                self.write('ВНИМАНИЕ: в реестре есть подтверждённые опечатки: '
                           +'; '.join(f"строка {e['row']}: {e['number']} (для расчёта серии {alias})"
                                     for e,alias in aliases)
                           +'. Исходные записи НЕ изменяются.')
        except Exception as exc:self.error(exc)

    def select_source(self):
        if self.active:return
        self.first_reference=None;self.facts={}
        try:
            self.cfg['source_passports']={**self.cfg.get('source_passports',{}),
                                          self.project.get():self.source_number.get()}
            project=model.select(model.read_registry(self.cfg['registry']),self.project.get())
            from .first_passport import discover_first_passport_for_config
            self.first_reference=discover_first_passport_for_config(self.cfg,project)
            self.invalidate()
            self.write('Образец '+self.source_number.get()+': '+self.first_reference['passport_root'])
            for role,ref in self.first_reference['sources'].items():self.write(role+': '+ref['path'])
            if self.first_reference['missing']:self.write('Нет форм выбранного образца: '+', '.join(self.first_reference['missing']))
        except Exception as exc:self.error(exc)

    def shipment_facts(self):
        facts=self.synchronized_facts() if hasattr(self,'synchronized_facts') else deepcopy(self.facts)
        if self.new.get() and hasattr(self,'shipment_date'):
            try:chosen=datetime.strptime(self.shipment_date.get().strip(),'%d.%m.%Y').date()
            except ValueError:raise Stop('Дата окончания / отгрузки: введите ДД.ММ.ГГГГ.')
            facts['shipment_date']=chosen.isoformat()
        return facts

    def profile(self):
        if not self.project.get() or self.active:return
        name=self.project.get();old=self.cfg.get('profiles',{}).get(name,{})
        d=tk.Toplevel(self.root);d.title('Собственные источники: '+name);self.fit(d,1070,720);d.transient(self.root)
        d.grab_set();d.bind('<Escape>',lambda e:d.destroy())
        root=tk.StringVar(value=old.get('root',''));fields={}
        ttk.Label(d,text='Папка подготовленных форм объекта').grid(row=0,column=0,padx=8,pady=8)
        ttk.Entry(d,textvariable=root,width=85).grid(row=0,column=1);ttk.Button(d,text='Выбрать',command=lambda:self.choose(root,True)).grid(row=0,column=2)
        roles=['appendix','vik','uzk','quality']+[r+str(n) for n in range(4) for r in ('akz','akz_appendix')]
        for row,role in enumerate(roles,1):
            original=(self.first_reference or {}).get('sources',{}).get(role)
            first_appendix=bool(role=='appendix' and original)
            # First Appendix is dynamically sourced from issued passport;
            # the other original acts are strictly read-only references.
            display=(original['path'] if first_appendix else
                     old.get('templates',{}).get(role,{}).get('path',''))
            var=tk.StringVar(value=display);fields[role]=var
            label={'appendix':('Приложение №1 (из первого паспорта)'
                              if first_appendix else 'Приложение №1'),
                   'vik':'ВИК','uzk':'УЗК','quality':'Документ о качестве'}.get(
                   role,('Приложение АКЗ №' if role.startswith('akz_appendix')
                         else 'Акт АКЗ №')+role[-1])
            ttk.Label(d,text=label).grid(row=row,column=0,padx=8,pady=5,sticky='w')
            ttk.Entry(d,textvariable=var,width=85,
                      state='readonly' if first_appendix else 'normal').grid(
                      row=row,column=1)
            ttk.Button(d,text='Авто' if first_appendix else 'Файл…',
                       state='disabled' if first_appendix else 'normal',
                       command=lambda v=var:self.choose(v,False)).grid(row=row,column=2)
        auto_catalog=(self.first_reference or {}).get('catalog')
        catalog=tk.StringVar(value=auto_catalog['path'] if auto_catalog
                             else old.get('catalog',''))
        ttk.Label(d,text='Справочник профилей / АКЗ (авто)' if auto_catalog
                  else 'Справочник марок (если есть)').grid(
                  row=13,column=0,padx=8,sticky='w')
        ttk.Entry(d,textvariable=catalog,width=85,
                  state='readonly' if auto_catalog else 'normal').grid(
                  row=13,column=1)
        ttk.Button(d,text='Авто' if auto_catalog else 'Файл…',
                   state='disabled' if auto_catalog else 'normal',
                   command=lambda:self.choose(catalog,False)).grid(
                   row=13,column=2)
        ttk.Label(d,text=('Первый паспорт найден автоматически: '+
            str((self.first_reference or {}).get('passport_root','не найден'))+
            '. Первое Приложение №1 читается оттуда напрямую. Остальные поля — очищенные рабочие формы, созданные по образцам; старые результаты, измерения и подписи не копируются.'),wraplength=990).grid(row=14,column=0,columnspan=3,pady=8)
        def save_profile():
            try:
                p=old|{'root':str(model.contained(root.get(),self.cfg['documents'])),'templates':{}}
                p.pop('catalog',None);p.pop('catalog_sha256',None)
                for role,var in fields.items():
                    if role=='appendix' and (self.first_reference or {}).get('sources',{}).get('appendix'):
                        # This visible path is deliberately outside the local
                        # CLEAN-form directory. Preserve any old cached
                        # record (for manual compatibility); the effective
                        # production source is auto-selected and SHA-pinned
                        # in model.profile, never stored as a guessed file.
                        previous=old.get('templates',{}).get('appendix')
                        if previous:p['templates']['appendix']=previous
                        continue
                    if var.get():
                        f=model.contained(var.get(),p['root'])
                        p['templates'][role]={'path':str(f),'sha256':sha(f)}
                if auto_catalog:
                    # Model.profile reads the pinned REAL project catalogue,
                    # which is outside the clean Templates folder.
                    previous=old.get('catalog')
                    if previous:
                        p['catalog']=previous
                        p['catalog_sha256']=old['catalog_sha256']
                elif catalog.get():
                    p['catalog']=str(model.contained(catalog.get(),p['root']))
                    p['catalog_sha256']=sha(p['catalog'])
                cfg=self.cfg|{'profiles':self.cfg.get('profiles',{})|{name:p}}
                effective=model.profile(cfg,name)
                from .templates import validate
                project=model.select(model.read_registry(cfg['registry']),name)
                variants=sorted({int(role[-1]) for role in effective['templates']
                                 if role.startswith('akz')})
                validate(effective,project,variants,cfg.get('training',False))
                self.cfg=cfg;save(self.settings_path(),cfg);self.facts={};self.invalidate();d.destroy()
            except Exception as exc:self.error(exc)
        def forms():
            if not root.get():self.error('Сначала выберите папку объекта.');return
            folder=Path(root.get())/('Формы_на_утверждение_'+uuid4().hex[:8])
            if not messagebox.askyesno('Новые формы','Создать отдельные незаполненные формы?\n'+str(folder)+'\nИх необходимо утвердить ответственному лицу перед рабочим применением.'):return
            d.destroy();self.run('forms',{'folder':str(folder)},lambda result:self.write('Созданы формы без результатов контроля: '+result['folder']+'. Выберите их в настройках объекта; первое приложение выберите из собственных источников.'))
        ttk.Button(d,text='Создать отдельные формы на утверждение',command=forms).grid(row=15,column=1,sticky='w',pady=10)
        ttk.Button(d,text='Сохранить привязки и SHA-256',command=save_profile).grid(row=15,column=1,sticky='e')

    def fact_dialog(self):
        if self.active:return
        try:
            v=model.visible_vom(self.vom.get())
            from .shared_uzk import find_shared_uzk,planned_control
            source=find_shared_uzk(self.cfg,require=True)
            uzk_plan=planned_control(source,self.project.get(),v)
            # An old saved settings.json may predate the third visible field.
            # Enable the single shared source in-memory for the current job;
            # persist it only if the operator saves Settings.
            self.cfg['uzk_registry']=str(source)
            self.cfg['use_shared_uzk']=True
        except Exception as exc:self.error(exc);return
        presets={'saved_at':'','variants':{}}
        if not self.cfg.get('training'):
            try:presets=akz_presets.last_for_project(self.project.get())
            except Stop as exc:
                # A damaged optional cache must never block new ACTUAL data.
                self.write('Подсказки АКЗ не загружены (ввод доступен вручную): '+str(exc))
        window=tk.Toplevel(self.root)
        window.title('Фактические результаты — '+self.project.get()+' '+self.number.get())
        self.fit(window,1160,760)
        window.transient(self.root)
        window.grab_set()
        # Footer is OUTSIDE the scrolling content, so Back stays visible.
        footer=ttk.Frame(window,padding=(10,6))
        footer.pack(side='bottom',fill='x')
        back_button=ttk.Button(footer,text='← Назад',
                               command=lambda:go_back())
        back_button.pack(side='right',padx=6)
        d=self.scroll_content(window)
        ttk.Label(d,text=(
            'Дата начала изготовления одинакова во всех паспортах проекта '
            'и берётся из его первого паспорта в реестре ВЗМК. '
            'Дата окончания и даты ВИК, УЗК, АКЗ новой отгрузки '
            'определяются по соответствующим текущим данным и формулам. '
            'Старые результаты контроля не переносятся.'
        ),wraplength=1090).pack(anchor='w',padx=10,pady=6)
        ttk.Label(d,text='Марка и расчётное количество УЗК заполняются из ВОМ и общего реестра. Профиль и № АКЗ — из справочника проекта. Редактирование: двойной щелчок. Старые заключения контроля автоматически не назначаются.').pack(pady=4)
        cols=('mark','profile','akz','uzk_quantity','standard')
        headings={'mark':'Марка','profile':'Профиль',
                  'akz':'№ АКЗ','uzk_quantity':'УЗК, шт','standard':'ГОСТ / НД'}
        table_frame=ttk.Frame(d)
        table_frame.pack(fill='both',expand=True,padx=10)
        grid=ttk.Treeview(table_frame,columns=cols,show='headings',height=12)
        for key in cols:
            grid.heading(key,text=headings[key])
            grid.column(key,width=105 if key in ('akz','uzk_quantity')
                        else 240)
        y=ttk.Scrollbar(table_frame,orient='vertical',command=grid.yview)
        grid.configure(yscrollcommand=y.set)
        y.pack(side='right',fill='y')
        grid.pack(side='left',fill='both',expand=True)
        old={str(r['mark']):dict(r) for r in self.facts.get('marks',[])}
        # Data not displayed in this simple table stays intact internally.
        # No deleted column is interpreted as a positive control finding.
        hidden_fields={}
        for mark in dict.fromkeys(str(r[0]) for r in v.rows):
            previous=old.get(mark,{})
            hidden_fields[mark]={
                k:previous[k] for k in ('paint','uzk_result','basis',
                                          'vik_result','vik_basis')
                if k in previous
            }
            values=[
                mark,
                str(previous.get('profile') or ''),
                str(previous.get('akz'))
                if previous.get('akz') not in (None,'') else '',
                str(uzk_plan[mark]['uzk_quantity']),
                str(previous.get('standard') or ''),
            ]
            grid.insert('','end',values=values)
        ttk.Label(
            d,text='Общий УЗК (только чтение): '+str(source)+
                 '. Плановый контроль по ВОМ: '+
                 str(sum(item['uzk_quantity'] for item in uzk_plan.values()))+
                 ' шт. Количество вычислено по маркам и не является '
                 'подтверждением результата контроля.',
            wraplength=1090,
        ).pack(anchor='w',padx=12,pady=4)
        # Apply CURRENT project's approved assignment catalogue when present;
        # never use another object's catalogue or invent AKZ0 for empty cells.
        from .catalog_assignments import current_assignments
        active_profile=self.cfg.get('profiles',{}).get(self.project.get(),{})
        try:
            active_profile=model.profile(self.cfg,self.project.get())
            catalogue=active_profile.get('catalog')
            if catalogue:
                assignments=current_assignments(
                    catalogue,active_profile['catalog_sha256'],
                    (r[0] for r in v.rows))
                for item in grid.get_children():
                    vals=list(grid.item(item,'values'))
                    chosen=assignments.get(str(vals[0]).strip(),{})
                    for key in ('profile','paint','akz'):
                        if key not in chosen:continue
                        if key=='paint':
                            # The paint remains an internal material mapping
                            # used by the AKZ generator, not a visible column.
                            if not str(hidden_fields[vals[0]].get('paint','')).strip():
                                hidden_fields[vals[0]]['paint']=chosen[key]
                        else:
                            index=cols.index(key)
                            if not str(vals[index]).strip():
                                vals[index]=str(chosen[key])
                    grid.item(item,values=vals)
                ttk.Label(d,text='Назначения загружены из справочника проекта: '+
                          str(catalogue)+'. Проверьте их перед выпуском.',
                          wraplength=1090).pack(anchor='w',padx=12)
        except Stop as exc:
            self.write('Справочник назначений не подставлен: '+str(exc))
            # A catalogue/archive failure must not conceal independently
            # installed approved climate forms. Keep assignments blank.
            if not active_profile.get('templates'):
                from .prepared_forms import prepared_project
                try:active_profile=prepared_project(self.cfg,self.project.get())
                except Stop as form_error:self.write(str(form_error))
        # Show only variants actually assigned by the current VOM/catalogue.
        # If the project has ONE approved AKZ form but marks are unassigned,
        # show its blank ranges as a convenience — do not assign AKZ to marks.
        from .akz_ui import climate_variants,display_variants,source_coats
        from .measurement_ranges import (measurements_from_ranges,
                                         magnus_dew_point,midrange)
        current_profile=active_profile
        variants=climate_variants(current_profile)
        if not variants and (self.first_reference or {}).get('sources',{}).get('akz0'):
            variants=[0]  # reference for display only, not confirmed assignment
        variant_forms={}
        # Proofs for АКЗ must be explicitly associated with that variant;
        # a generic VIK/UZK proof cannot be silently called АКЗ evidence.
        akz_proofs={n:list(self.facts.get('akz_proofs',{}).get(str(n),[]))
                    for n in range(4)}
        prev={int(entry['variant']):entry
              for entry in self.facts.get('climates',[])}
        akz_notice=ttk.Label(d,wraplength=1080,justify='left')
        akz_notice.pack(fill='x',padx=12,pady=5)
        for n in variants:
            formref=current_profile.get('templates',{}).get('akz'+str(n))
            if formref:
                try:coats=source_coats(formref)
                except Stop as exc:
                    self.write('Не удалось проверить число слоёв АКЗ №'+str(n)+
                               ': '+str(exc))
                    coats=None
            else:coats=None
            group=ttk.LabelFrame(d,text='АКЗ №'+str(n)+
                                 (' — '+str(coats)+' слой(я)' if coats else
                                  ' — требуется рабочая форма'),
                                 padding=8)
            old_interval=prev.get(n,{}).get('measurement_intervals',{})
            # The user may return with only SOME ranges entered. Draft values
            # win over previously validated ranges, without needing a SAVE.
            draft=(self.facts.get('akz_drafts') or {}).get(str(n),{})
            fields={}
            instructions=ttk.Label(
                group,
                text='Только фактические диапазоны текущих измерений: '
                     '«От» и «До». Расчётное среднее = (От + До) / 2. '
                     'Это не заменяет первичный протокол.',
                wraplength=1060,
            )
            instructions.grid(row=0,column=0,columnspan=3,sticky='w',pady=4)
            ttk.Label(group,text='Параметр').grid(row=1,column=0,sticky='w')
            ttk.Label(group,text='От').grid(row=1,column=1,sticky='w')
            ttk.Label(group,text='До').grid(row=1,column=2,sticky='w')
            labels={'air':'Воздух, °C','surface':'Поверхность, °C',
                    'humidity':'Влажность, %'}
            if coats==1:labels['wet']='Толщина мокрого слоя, мкм'
            measurements=list(labels.items())+[
                ('layer_'+str(i),'Толщина сухого слоя, мкм' if coats==1 else 'Фактическая толщина слоя '+str(i)+', мкм')
                for i in range(1,(coats or 0)+1)]
            for line,(key,title) in enumerate(measurements,start=2):
                stored=(draft.get('layer_intervals',old_interval.get('layers',[]))
                        if key.startswith('layer_') else
                        draft.get(key,old_interval.get(key)))
                if isinstance(stored,list):
                    number=int(key.split('_')[1])
                    stored=next((item for item in stored
                                 if item.get('layer')==number),{})
                stored=stored or {}
                ttk.Label(group,text=title).grid(
                    row=line,column=0,sticky='w',padx=4,pady=3)
                pair={}
                for column,label in ((1,'min'),(2,'max')):
                    var=tk.StringVar(value=str(stored.get(label,'')))
                    entry=ttk.Entry(group,textvariable=var,width=20)
                    entry.grid(row=line,column=column,sticky='w',padx=4,pady=3)
                    # Replacing a saved value must not append to its old text.
                    entry.bind('<FocusIn>',lambda e:e.widget.after_idle(
                        lambda w=e.widget:(w.selection_range(0,'end'),w.icursor('end')) if w.winfo_exists() else None))
                    pair[label]=var
                fields[key]=pair

            dew_status=ttk.Label(
                group,text='Точка росы рассчитывается по температуре воздуха '
                           'и влажности (формула Магнуса).',
                wraplength=1010,justify='left')
            dew_status.grid(row=len(measurements)+2,column=0,
                            columnspan=3,sticky='w',pady=8)
            def update_dew_display(*_, inputs=fields, target=dew_status):
                required=[inputs[k][edge].get().strip()
                          for k in ('air','surface','humidity')
                          for edge in ('min','max')]
                if not all(required):
                    target.configure(text='Точка росы: будет рассчитана '
                                     'после ввода воздуха и влажности.')
                    return
                try:
                    mean_air=midrange(
                        inputs['air']['min'].get(),inputs['air']['max'].get(),
                        'воздух',-60,80)
                    mean_surface=midrange(
                        inputs['surface']['min'].get(),
                        inputs['surface']['max'].get(),
                        'поверхность',-60,100)
                    mean_humidity=midrange(
                        inputs['humidity']['min'].get(),
                        inputs['humidity']['max'].get(),
                        'влажность',0,100)
                    dew=magnus_dew_point(mean_air,mean_humidity)
                    if float(mean_air)<=float(mean_surface):
                        target.configure(
                            text='Ошибка АКЗ: температура воздуха должна '
                                 'быть выше температуры поверхности.')
                    elif float(mean_surface)<float(dew)+3:
                        target.configure(
                            text='Ошибка АКЗ: поверхность должна быть '
                                 'минимум на 3 °C теплее точки росы '
                                 '('+str(dew)+' °C).')
                    else:
                        target.configure(
                            text='Точка росы автоматически: '+str(dew)+
                                 ' °C. Воздух: '+str(mean_air)+' °C; '
                                 'поверхность: '+str(mean_surface)+' °C; '
                                 'влажность: '+str(mean_humidity)+' %. '
                                 'Проверка климата — PASS.')
                except (Stop,ValueError,TypeError) as exc:
                    target.configure(text='Проверьте диапазоны: '+str(exc))
            for k in ('air','surface','humidity'):
                for edge in ('min','max'):
                    fields[k][edge].trace_add('write',update_dew_display)
            update_dew_display()
            # Detailed 3-coat material batches remain accessible ONLY when
            # an actual approved 3-layer form is selected. They are never
            # shown for the one-layer Kazakhstan act.
            details=None
            variant_forms[n]={'frame':group,'fields':fields,
                              'coats':coats,'details':details}

        def assigned_marks():
            result=[]
            for item in grid.get_children():
                vals=grid.item(item,'values')
                raw=str(vals[cols.index('akz')]).strip()
                if not raw:
                    value=None
                else:
                    try:value=int(raw)
                    except (TypeError,ValueError):value=raw
                result.append({'mark':str(vals[0]),'akz':value})
            return result

        def refresh_akz():
            try:visible=display_variants(assigned_marks(),current_profile)
            except Stop as exc:
                akz_notice.configure(text='Некорректный № АКЗ: '+str(exc))
                return
            # With a single original 1-coat form, show its blank group even
            # before the catalogue fills any current mark assignments.
            if not visible and len(variant_forms)==1:
                visible=list(variant_forms)
            for n,data in variant_forms.items():
                if n in visible:data['frame'].pack(fill='x',padx=10,pady=6)
                else:data['frame'].pack_forget()
            unsupported=[n for n in visible if n not in variant_forms]
            if unsupported:
                akz_notice.configure(text='Для текущих марок назначены АКЗ '+
                    ', '.join(map(str,unsupported))+
                    ', но нет соответствующих утверждённых рабочих форм. '
                    'Ничего не подставлено автоматически.')
            elif visible:
                akz_notice.configure(text='Показаны только применяемые '
                    'АКЗ: '+', '.join(map(str,visible))+
                    '. Пустые назначения марок не приравниваются к АКЗ №0.')
            else:
                akz_notice.configure(text='В справочнике нет подтверждённых '
                    'назначений АКЗ. Выберите действующие виды для марок.')
        from .automatic_acts import find_approved_geometry,calculate_design_characteristics
        from .geometry_selection import GeometrySelection
        geometry=find_approved_geometry(self.cfg,require=False)
        selection=GeometrySelection(geometry,self.cfg,self.project.get()) if geometry else None
        standard_status=ttk.Label(d,wraplength=1080,justify='left')
        standard_status.pack(fill='x',padx=12,pady=5)
        if geometry is None:self.write('Нет утверждённого Профили.xlsx: расчёт УЗК будет заблокирован при формировании.')
        def standard_options(item):
            vals=grid.item(item,'values')
            return selection.options(str(vals[1])) if selection else []
        def update_standards():
            errors=[]
            for item in grid.get_children():
                vals=list(grid.item(item,'values'))
                try:vals[4]=selection.resolve(str(vals[0]),str(vals[1])) if selection else ''
                except Stop as exc:errors.append(str(exc))
                grid.item(item,values=vals)
            text=('ГОСТ по умолчанию — первый допустимый в порядке строк Профили.xlsx. '
                  'Проверьте соответствие проекту. Двойной щелчок: общий выбор профиля или исключение марки.')
            if errors:text+='\n'+'\n'.join(errors)
            standard_status.configure(text=text)
            return errors
        def changed_standard(item,col,old,new):
            if col==4 and new!=old and selection:
                vals=list(grid.item(item,'values'))
                scope=messagebox.askyesnocancel('Область выбора ГОСТ',
                    'Применить «'+new+'» ко всем маркам с профилем '+str(vals[1])+'?\n'
                    'Да — общий выбор профиля (остальные индивидуальные исключения сохранятся).\n'
                    'Нет — индивидуальное исключение только для марки '+str(vals[0])+'.',parent=window)
                if scope is None:
                    vals[4]=old;grid.item(item,values=vals);return
                selection.choose(str(vals[0]),str(vals[1]),new,individual=not scope)
            if col in (1,4):update_standards()
        def recalculate(*_):
            refresh_akz()
            if not geometry:return
            current=[dict(zip(cols,grid.item(item,'values'))) for item in grid.get_children()]
            try:
                design=calculate_design_characteristics(v,{'marks':current},geometry,expected_sha=selection.pin)
                picked={str(grid.item(i,'values')[0]) for i in grid.selection()}
                parts=[r for r in design if r['mark'] in picked]
                suffix='; '.join(r['mark']+': S='+r['thickness_mm']+' мм, L='+r['length_mm']+' мм, '+r['frequency_mhz']+' МГц, '+r['angle_deg']+'°' for r in parts)
                update_standards()
                if suffix:standard_status.configure(text=standard_status.cget('text')+'\n'+suffix)
            except Stop as exc:standard_status.configure(text=standard_status.cget('text')+'\n'+str(exc))
        update_standards()
        self.editable(grid,immutable={0,3},on_change=recalculate,choices={4:standard_options},on_cell_change=changed_standard)
        grid.bind('<<TreeviewSelect>>',recalculate)
        refresh_akz()

        def go_back():
            """Save all entered values to this session, including partial
            ranges, without issuing any certificate or changing Excel.
            """
            try:
                # A visible Treeview cell may still be in an Entry editor.
                # Explicitly commit it; mouse clicks on a Tk button do not
                # consistently trigger FocusOut on every Windows theme.
                pending=getattr(grid,'_pending_editor_commit',None)
                if callable(pending):pending()
                from .simple_facts import save_screen_draft
                visible=[
                    dict(zip(cols,grid.item(item,'values')))
                    for item in grid.get_children()
                ]
                climate_drafts={}
                for n,info in variant_forms.items():
                    values={
                        'coats':info['coats'],
                        'air':{},'surface':{},'humidity':{},
                        'layer_intervals':[],
                    }
                    for k in ('air','surface','humidity',*(['wet'] if info['coats']==1 else [])):
                        values[k]={edge:info['fields'][k][edge].get().strip()
                                   for edge in ('min','max')}
                    if info['coats'] in (1,3):
                        for i in range(1,info['coats']+1):
                            key='layer_'+str(i)
                            values['layer_intervals'].append(
                                {'layer':i,**{
                                    edge:info['fields'][key][edge].get().strip()
                                    for edge in ('min','max')
                                }})
                    details=info['details']
                    if details is not None:
                        detailkeys=('layer','material','paint_batch',
                                    'hardener_batch','color',
                                    'application_date','time_start','time_end',
                                    'wet_average')
                        saved_rows=[]
                        for item in details.get_children():
                            vals=dict(zip(detailkeys,details.item(item,'values')))
                            if any(str(vals[k]).strip() for k in detailkeys[1:]):
                                vals['layer']=int(vals['layer'])
                                saved_rows.append(vals)
                        if saved_rows:values['layer_details']=saved_rows
                    climate_drafts[str(n)]=values
                draft=save_screen_draft(
                    self.facts,self.project.get(),self.number.get(),
                    self.cfg.get('training',False),
                    visible,hidden_fields,climate_drafts,
                )
                self.facts=draft
                if selection:
                    selection.verify_source()
                    if update_standards():raise Stop('Проверьте ГОСТ / НД: источник или сохранённый выбор изменился.')
                    self.cfg=selection.settings()
                save(self.settings_path(),self.cfg)
                self.invalidate()
                window.destroy()
                self.write('Введённые марки и диапазоны АКЗ сохранены '
                           'в текущей заявке. Реестр не изменён. '
                           'Проверка достоверности и комплектности '
                           'выполняется при «Сформировать пакет».')
            except Exception as exc:self.error(exc)
        window.protocol('WM_DELETE_WINDOW',go_back)
        window.bind('<Escape>',lambda event:go_back())

    def editable(self,tree,immutable=set(),on_change=None,choices=None,on_cell_change=None):
        tree._pending_editor_commit=None
        def edit(event):
            item=tree.identify_row(event.y);column=tree.identify_column(event.x)
            if not item or not column:return
            col=int(column[1:])-1
            if col in immutable:return
            x,y,w,h=tree.bbox(item,column)
            if col in (choices or {}):
                options=choices[col](item)
                if not options:self.error('Нет утверждённого ГОСТ для профиля этой марки.');return
                entry=ttk.Combobox(tree,values=options,state='readonly')
                current=tree.item(item,'values')[col]
                entry.set(current if current in options else '')
            else:
                entry=ttk.Entry(tree);entry.insert(0,tree.item(item,'values')[col])
            entry.place(x=x,y=y,width=w,height=h);entry.focus_set()
            def done(e=None):
                if not entry.winfo_exists():return
                values=list(tree.item(item,'values'))
                old=values[col]
                values[col]=entry.get()
                tree.item(item,values=values)
                entry.destroy()
                tree._pending_editor_commit=None
                if on_cell_change:on_cell_change(item,col,old,values[col])
                if on_change:on_change()
            tree._pending_editor_commit=done
            def lost_focus(e=None):
                if col not in (choices or {}):
                    done(e);return
                # A native ttk popup takes focus while the operator chooses.
                # Destroying its editor here would close the list immediately.
                def finish_unposted():
                    if not entry.winfo_exists():return
                    popup=entry.tk.call('ttk::combobox::PopdownWindow',str(entry))
                    if entry.tk.getboolean(entry.tk.call('winfo','ismapped',popup)):return
                    done()
                entry.after_idle(finish_unposted)
            entry.bind('<Return>',done);entry.bind('<FocusOut>',lost_focus)
            if col in (choices or {}):entry.bind('<<ComboboxSelected>>',done)
        tree.bind('<Double-1>',edit)

    def catalog(self):
        if self.active:return
        target=filedialog.asksaveasfilename(title='НОВЫЙ черновик справочника',defaultextension='.xlsx',initialfile='Черновик справочника марок.xlsx')
        if target:self.run('catalog',{'vom':self.vom.get(),'target':target},lambda r:self.write('Черновик: '+r['path']+'; профиль/краска/АКЗ пустые.'))

    def check(self):
        """The single working path: validate, preview, one explicit consent."""
        if self.active:return
        if self.plan and self.preview_pdf:return self.review_preview()
        if self.preliminary and self.preview_pdf:return self.review_preliminary()
        try:
            if not self.project.get() or not self.number.get() or not self.vom.get():raise Stop('Выберите объект, исходный паспорт и введите отгрузочную ВОМ.')
            current=self.shipment_facts()
            try:
                from .owner_control import engineering_facts
                current=engineering_facts(current,model.profile(self.cfg,self.project.get()))
            except Stop:
                pass  # incomplete ranges/assignments may still have a marked preview
            else:
                self.run('check',{'cfg':self.cfg,'project':self.project.get(),'number':self.number.get(),'vom':self.vom.get(),'facts':current,'new':self.new.get()},self.checked)
                return
            self.run('draft',{'cfg':self.cfg,'project':self.project.get(),'number':self.number.get(),'vom':self.vom.get(),'facts':current,'new':self.new.get()},self.preliminary_ready)
        except Exception as exc:self.error(exc)

    def preliminary_ready(self,result):
        self.preliminary=result;self.preview_pdf=result['preview_pdf']
        self.last_output_folder=result['folder'];self.last_output_project=self.project.get()
        self.write('Предварительный комплект подготовлен: '+str(result['files'])+' файлов, PDF '+str(result['pages'])+' страниц. Рабочие Excel не изменены.')
        for role,reason in result.get('blocked_documents',{}).items():self.write(role+': '+reason)
        self.buttons[0].configure(text='Просмотреть и подтвердить')
        self.review_preliminary()

    def review_preliminary(self):
        if not self.preliminary or not self.preview_pdf:return
        try:open_external(self.preview_pdf)
        except Exception as exc:self.error('Не удалось открыть предварительный PDF: '+str(exc));return
        summary=self.preliminary['summary']
        text=('Проект: '+self.project.get()+'\nПаспорт: '+self.number.get()+'\nДата отгрузки: '+self.shipment_date.get()+
              '\n'+str(summary['quantity'])+' изделий / '+summary['mass_kg']+' кг / '+summary['area_m2']+' м²'+
              '\nПапка просмотра: '+self.preliminary['folder']+
              '\n\nПроверьте все страницы предварительного PDF. Исторические заключения и подписи образца не подтверждают текущий контроль. После одного подтверждения программа независимо проверит текущие факты и право записи. При их отсутствии выпуск будет запрещён. Подтверждаете выпуск?')
        if not messagebox.askyesno('Предварительный просмотр и выпуск',text):
            self.write('Выпуск отменён. Предварительный комплект сохранён; рабочие Excel не изменены.');return
        try:
            from .owner_control import engineering_facts
            engineering_facts(self.shipment_facts(),model.profile(self.cfg,self.project.get()))
        except Stop as exc:
            self.error('Выпуск пока невозможен: '+str(exc));return
        self.preliminary=None;self.preview_pdf=None
        self.check()

    def drafted(self,result):
        self.plan=None
        self.preview_pdf=result['preview_pdf']
        self.last_output_folder=result['folder'];self.last_output_project=self.project.get()
        summary=result['summary']
        self.write('ЧЕРНОВОЙ ПАКЕТ: '+str(summary['marks'])+' марок / '+
                   str(summary['quantity'])+' изделий / '+str(summary['mass_kg'])+
                   ' кг / '+str(summary['area_m2'])+' м²; план УЗК '+
                   str(summary['uzk_quantity'])+' шт. '+str(result['files'])+
                   ' файлов, PDF '+str(result['pages'])+' стр.\n'+result['folder'])
        for note in result['notes']:self.write(note)
        self.progress.configure(text='Черновики готовы. Реестр не изменён. Официальный выпуск не выполнен.')
        message=('Подготовлены редактируемые черновики и общий PDF.\n'
                 'Результаты нового контроля не утверждены. Данные эталона '
                 'сохранены как образец и требуют проверки.\n'
                 'Реестр и общий УЗК не изменены. Это НЕ подтверждение выпуска.\n\n'+
                 result['folder']+'\n\nОткрыть черновой PDF?')
        if messagebox.askyesno('Черновой пакет — без выпуска',message):
            try:open_external(self.preview_pdf)
            except OSError as exc:self.error('Не удалось открыть PDF: '+str(exc))
        else:self.write('Просмотр отложен. Черновики сохранены; выпуск не выполнялся.')

    def check_official(self):
        """Retained strict flow for already supplied, current verified facts."""
        if self.active:return
        if self.plan and self.preview_pdf:
            self.review_preview()
            return
        if not all(self.cfg.get(k) for k in ('registry','documents','output')):
            self.error('Сначала выберите рабочие пути в «Настройки путей».');return
        if not self.project.get() or not self.number.get() or not self.vom.get():
            self.error('Выберите объект, паспорт и отгрузочную ВОМ.');return
        if not self.facts:
            if self.new.get():
                try:
                    from .registry_dates import require_new_date_formulas
                    r=model.read_registry(self.cfg['registry'])
                    require_new_date_formulas(self.cfg['registry'],r,model.select(r,self.project.get()))
                except Exception as exc:
                    self.error(str(exc));return
            self.error('Откройте «Фактические результаты» и проверьте сведения '
                       'текущей отгрузки. Старые протоколы не подставляются.')
            return
        from .simple_facts import prepare_for_check
        from .working_facts import apply_current_controls
        try:
            if not self.facts.get('evidence') and not self.facts.get('current_controls'):
                raise Stop('Официальный выпуск запрещён: отсутствуют достоверные текущие результаты ВИК/УЗК и подтверждённые измерения АКЗ. Образец исходного паспорта не является новым контролем. Предварительный PDF сохранён; три рабочих файла не изменены.')
            self.facts=self.shipment_facts()
            current=apply_current_controls(self.facts) if self.facts.get('current_controls') else dict(self.facts)
            confirmed_facts=prepare_for_check(current)
        except Exception as exc:
            self.error(exc)
            return
        self.run('check',{'cfg':self.cfg,'project':self.project.get(),
                          'number':self.number.get(),'vom':self.vom.get(),
                          'facts':confirmed_facts,
                          'new':self.new.get()},self.checked)

    def checked(self,result):
        self.plan=result['plan']
        self.facts=result['data']['facts']
        d=result['data']
        self.write(f"Данные проверены: {d['project']} | {d['number']} | "
                   f"{d['marks']} марок / {d['quantity']} изделий / "
                   f"{d['mass_kg']} кг. Дата паспорта: {d['date']}.")
        self.run('prepare',{'path':self.plan},self.prepared)
    def prepare(self):self.run('prepare',{'path':self.plan},self.prepared)
    def verify(self):self.run('verify',{'path':self.plan,'preview':True},self.verified)
    def verified(self,result):
        self.preview_pdf=result.get('preview_pdf')
        if not self.preview_pdf or not Path(self.preview_pdf).is_file():
            self.error('Не создан предварительный PDF: выпуск запрещён.')
            return
        self.buttons[0].configure(text='Просмотреть и подтвердить')
        self.review_preview()

    def open_documents(self):
        """Open the latest completed passport or the current prepared drafts."""
        candidates=[]
        if self.last_output_folder and self.last_output_project==self.project.get():
            candidates.append(Path(self.last_output_folder))
        if self.plan:
            preview=Path(self.plan).parent/'preview.json'
            if preview.is_file():
                folder=Path(read(preview)['pdf']).parent
                if folder.is_dir() and not folder.is_symlink():
                    open_external(folder);return
            candidates.append(Path(self.plan).parent/'documents')
        for folder in candidates:
            if folder.is_dir() and not folder.is_symlink():
                try:
                    open_external(str(folder))
                except OSError as exc:
                    self.error('Не удалось открыть папку документов: '+str(exc))
                return
        messagebox.showinfo('Документы','Комплект пока не сформирован. '
                            'После подготовки Word/Excel здесь откроется '
                            'папка паспорта.')

    def issue(self):
        if not self.plan or not self.preview_pdf:return
        # Only entered from explicit consent in review_preview(), with exact
        # project/date/registry/output shown. Do not ask a SECOND approval.
        self.run('confirm',{'path':self.plan,
                            'responsible':self.facts.get('responsible',''),
                            'reviewed':True,
                            'approved':True,
                            'current_control':True},
                 lambda r:self.run('release',{'path':self.plan},self.released))
    def released(self,r):self.run('commit',{'path':self.plan,'release_result':self.result_path},self.committed)
    def committed(self,r):self.run('readback',{'journal':r['journal']},lambda v:self.run('complete',{'journal':r['journal'],'readback_result':self.result_path},self.completed))
    def completed(self,r):
        if r.get('folder'):
            self.last_output_folder=str(r['folder'])
            self.last_output_project=self.project.get()
        self.write('Комплект выпущен. Документы и PDF: '+str(r.get('folder','')))
        # Store preferences only for an actually completed production issue.
        # Old protocols, measurement thicknesses, dates and approvals never
        # become facts of the next passport.
        if r.get('state')=='ГОТОВО' and not self.cfg.get('training') and self.facts.get('kind')=='ACTUAL':
            try:
                if akz_presets.remember(self.project.get(),self.facts.get('climates',[]),self.facts.get('akz_date')):
                    self.write('Последние климатические показатели и диапазоны слоёв сохранены только для объекта '+self.project.get()+'.')
            except Exception as exc:
                self.write('Выпуск завершён, но сохранить подсказки АКЗ не удалось: '+str(exc))
        messagebox.showinfo('Комплект готов','Word, Excel и общий PDF сохранены. Учёт отгрузки проверен.\n\n'+str(r.get('folder','')));self.facts={};self.invalidate();self.refresh()
    def recovery(self):
        if self.active:return
        try:
            paths=core.pending(self.cfg['registry'])
            if not paths:messagebox.showinfo('Восстановление','Незавершённых операций нет.');return
            journal=str(paths[0])
            if messagebox.askyesno('Откат незавершённой операции',journal+'\nВосстановить только при совпадении SHA? Изменённые пользователем документы/реестр не будут перезаписаны.'):
                self.run('recover',{'journal':journal},lambda r:self.write(r))
        except Exception as exc:self.error(exc)

    def run(self,task,request,callback):
        if self.active:return
        root=local_root()/'gui-jobs'/uuid4().hex;root.mkdir(parents=True)
        save(root/'request.json',request)
        prefix=[sys.executable] if getattr(sys,'frozen',False) else [sys.executable,'-B',str(Path(__file__).resolve().parents[1]/'production.py')]
        self.active=subprocess.Popen(prefix+['--task',task,'--request',str(root/'request.json'),'--result',str(root/'result.json')],creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
        for widget in (self.projects,self.numbers,self.new_widget,self.vom_widget,self.source_numbers,self.shipment_entry):widget.configure(state='disabled')
        states=[b['state'] for b in self.buttons]
        for b in self.buttons:b.configure(state='disabled')
        labels={'check':'Проверяются исходные данные','prepare':'Формируются Word и Excel','verify':'Проверяется комплект и создаётся PDF','confirm':'Сохраняется ваше подтверждение','release':'Готовятся окончательные документы','commit':'Сохраняется отгрузка с резервными копиями','readback':'Проверяется запись отгрузки','complete':'Завершается сохранение','recover':'Восстанавливается незавершённая операция'}
        self.progress.configure(text=labels.get(task,'Выполняется операция')+'. Дождитесь завершения.');self.root.after(500,lambda:self.poll(root,task,callback,states))
    def poll(self,root,task,callback,states):
        if self.active.poll() is None:self.root.after(500,lambda:self.poll(root,task,callback,states));return
        code=self.active.returncode;self.active=None
        self.projects.configure(state='readonly');self.numbers.configure(state='readonly');self.new_widget.configure(state='normal');self.vom_widget.configure(state='normal')
        self.source_numbers.configure(state='readonly');self.shipment_entry.configure(state='normal' if self.new.get() else 'disabled')
        try:
            r=read(root/'result.json');self.result_path=str(root/'result.json')
            if code or r['status']!='PASS':raise Stop(str(r.get('error',r))+'\nЖурнал: '+str(root))
            for b,state in zip(self.buttons,states):b.configure(state=state)
            self.progress.configure(text='Проверка завершена.');callback(r)
        except Exception as exc:self.invalidate();self.progress.configure(text='Операция остановлена. Откройте сообщение об ошибке.');self.error(exc)

    def prepared(self,result):
        self.write('Word и Excel подготовлены: '+str(result['files'])+' файлов.')
        self.run('verify',{'path':self.plan,'preview':True},self.verified)


    def review_preview(self):
        if not self.preview_pdf or not self.plan:return
        try:
            d=read(self.plan)
            open_external(self.preview_pdf)
        except Exception as exc:
            self.error('Не удалось открыть подготовленный PDF: '+str(exc))
            return
        message=('Выпуск комплекта документов\n'
                 'Проект: '+str(d['project'])+'\n'
                 'Номер: '+str(d['number'])+'\n'
                 'Дата паспорта: '+str(d['date'])+'\n'+
                 str(d['marks'])+' марок / '+str(d['quantity'])+' изделий / '+str(d['mass_kg'])+' кг / '+str(d['area_m2'])+' м²\n'+
                 'Реестр: '+str(d['registry'])+'\n'
                 'Папка результата: '+str(d['target'])+'\n'+
                 (('Справочник отгрузок: '+d['accounting']['catalog']+'\nОбщий УЗК: '+d['accounting']['uzk']+
                   '\nСписание УЗК: '+str(sum(d['accounting']['controls'].values()))+' шт. Три файла — с резервными копиями.\n') if d.get('accounting') else '')+
                 ('Будет добавлена одна строка в реестр.\n' if d['new']
                  else 'Новых строк реестра не будет.\n')+
                 '\nОткрылся PDF «Предварительный просмотр — не выпущено». '
                 'Проверьте ВСЕ страницы, марки, фактические результаты '
                 'контроля, материалы, толщины и даты.\n\n'
                 'Это ЕДИНСТВЕННОЕ подтверждение официального выпуска.\n'
                 'Нажимая «Да», вы подтверждаете, что текущие результаты '
                 'контроля достоверны, утверждённые формы действуют, '
                 'вы вправе выпустить документы и проверили все страницы.\n'
                 'После «Да» будут выполнены итоговые проверки и запись '
                 'в реестр с резервной копией.\n'
                 +STATEMENT+'\nЕсли контроль не выполнен или результаты неизвестны, нажмите «Нет».\nПодтверждаете текущий контроль и выпуск?')
        if not messagebox.askyesno('Предварительный просмотр и выпуск',message):
            self.write('Подтверждение отложено; предварительный PDF доступен. '
                       'Реестр не изменён.')
            return
        self.issue()


def launch():
    root=tk.Tk();ttk.Style(root).theme_use('clam');App(root);root.mainloop()
