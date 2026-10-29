import os,subprocess,sys,traceback
from pathlib import Path
from uuid import uuid4
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk,filedialog,messagebox,simpledialog
from .storage import local_root,read,save,sha,Stop
from . import model,core,akz_presets,VERSION
from .external_open import open_external
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
        self.source_number=tk.StringVar()
        self.shipment_date=tk.StringVar(value=date.today().strftime('%d.%m.%Y'))
        self.shipment_date.trace_add('write',lambda *_:self.invalidate())
        self.project.trace_add('write',lambda *_:self.inputs_changed())
        self.number.trace_add('write',lambda *_:self.inputs_changed())
        self.vom.trace_add('write',lambda *_:self.inputs_changed())
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
        ttk.Label(form,text='Отгрузочная ВОМ').grid(row=2,column=0,sticky='w')
        self.vom_widget=ttk.Entry(form,textvariable=self.vom);self.vom_widget.grid(row=2,column=1,sticky='ew',padx=8,pady=5)
        ttk.Button(form,text='Выбрать ВОМ…',command=self.choose_vom).grid(row=2,column=2)
        self.vom_widget.bind('<FocusOut>',lambda e:self.summarize_vom() if self.vom.get() and not self.active else None)
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
        self.plan=None;self.preview_pdf=None
        for b in self.buttons:
            b.configure(state='normal' if not self.active else 'disabled',text='Сформировать пакет')

    def inputs_changed(self):
        self.facts={};self.invalidate()

    def error(self,exc):self.write(str(exc));messagebox.showerror('Не удалось продолжить',str(exc))

    def choose(self,var,directory):
        if self.active:return
        path=filedialog.askdirectory() if directory else filedialog.askopenfilename(filetypes=[('Excel/Word','*.xlsx *.xlsm *.docx'),('Все файлы','*.*')])
        if path:var.set(path)

    def settings_path(self):return getattr(self,'config_path',local_root()/'settings.json')

    def choose_vom(self):
        self.choose(self.vom,False)
        if self.vom.get():self.summarize_vom()

    def summarize_vom(self):
        try:
            v=model.visible_vom(self.vom.get())
            self.write(f'ВОМ: {len(v.rows)} марок / {v.quantity} изделий / масса {v.mass:.2f} кг / площадь {v.area:.3f} м². Исходный файл не изменён.')
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
            'отдельно по номеру и дате. Очищенные технические формы '
            'хранятся внутри установленной программы и не выбираются '
            'при обычной работе. До подтверждения ничего не списывается; выпуск обновляет остатки с резервной копией.',
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
        standard_grid=None
        if self.project.get() and self.vom.get() and effective.get('catalog'):
            try:
                from .catalog_assignments import read_assignments
                from .automatic_acts import profile_standards
                assignments=read_assignments(effective['catalog'])
                geometry=find_approved_geometry(self.cfg,require=True)
                frame=ttk.LabelFrame(dialog,text='ГОСТ / НД текущих марок — двойной щелчок для выбора')
                frame.grid(row=6,column=0,columnspan=3,padx=12,pady=6,sticky='ew')
                standard_grid=ttk.Treeview(frame,columns=('mark','profile','standard'),show='headings',height=4)
                for key,title,width in [('mark','Марка',200),('profile','Профиль',220),('standard','ГОСТ / НД',560)]:
                    standard_grid.heading(key,text=title);standard_grid.column(key,width=width)
                standard_grid.pack(fill='x')
                current={r['mark']:r for r in self.facts.get('marks',[])};cache={}
                for mark in dict.fromkeys(str(r[0]) for r in model.visible_vom(self.vom.get()).rows):
                    profile=current.get(mark,{}).get('profile') or assignments.get(mark,{}).get('profile','')
                    if profile not in cache:cache[profile]=profile_standards(geometry,profile)
                    options=cache[profile]
                    standard_choice=self.cfg.get('geometry_choices',{}).get(self.project.get(),{}).get(mark) or current.get(mark,{}).get('standard') or (options[0] if len(options)==1 else '')
                    standard_grid.insert('','end',values=(mark,profile,standard_choice))
                self.editable(standard_grid,immutable={0,1},choices={2:lambda item:cache.get(standard_grid.item(item,'values')[1],[])})
            except Stop as exc:self.write('ГОСТ / НД: '+str(exc))

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
                if standard_grid is not None:
                    pending=getattr(standard_grid,'_pending_editor_commit',None)
                    if callable(pending):pending()
                    choices=dict(cfg.get('geometry_choices',{}).get(self.project.get(),{}))
                    for item in standard_grid.get_children():
                        mark,profile,standard=standard_grid.item(item,'values')
                        if standard:
                            if standard not in profile_standards(geometry,profile):raise Stop('Выбран неутверждённый ГОСТ для '+mark)
                            choices[mark]=standard
                        else:choices.pop(mark,None)
                    cfg['geometry_choices']={**cfg.get('geometry_choices',{}),self.project.get():choices}
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
        facts=dict(self.facts)
        if self.new.get() and hasattr(self,'shipment_date'):
            try:chosen=datetime.strptime(self.shipment_date.get().strip(),'%d.%m.%Y').date()
            except ValueError:raise Stop('Дата окончания / отгрузки: введите ДД.ММ.ГГГГ.')
            facts['shipment_date']=chosen.isoformat()
        for row in facts.get('marks',[]):
            choice=self.cfg.get('geometry_choices',{}).get(self.project.get(),{}).get(row['mark'])
            if choice:row['standard']=choice
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
                str(self.cfg.get('geometry_choices',{}).get(self.project.get(),{}).get(mark) or previous.get('standard') or ''),
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
        from .working_facts import form_material,save_controls,required_doc_fields
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
        # The current control is entered here once. These entries describe
        # this shipment, never a finding copied from the source passport.
        control_vars={};reselected=set()
        def control_row(frame,role,row,*,verdict=False,material=''):
            previous=self.facts.get('current_controls',{}).get(role,{})
            fields={'path':tk.StringVar(value=previous.get('path',''))}
            if verdict:
                fields['result']=tk.StringVar(value=previous.get('result',''))
                ttk.Label(frame,text=('ВИК' if role=='VIK' else 'УЗК')+' — заключение по текущей отгрузке').grid(row=row,column=0,sticky='w',padx=4,pady=4)
                ttk.Combobox(frame,textvariable=fields['result'],values=('', 'ГОДЕН','НЕ ГОДЕН'),state='readonly',width=18).grid(row=row,column=1,sticky='w')
                row+=1
            if role.startswith('AKZ'):
                fields['material']=tk.StringVar(value=material or previous.get('material',''))
                ttk.Label(frame,text='ЛКМ по выбранной форме').grid(row=row,column=0,sticky='w',padx=4,pady=4)
                ttk.Entry(frame,textvariable=fields['material'],width=66,state='readonly' if material else 'normal').grid(row=row,column=1,columnspan=2,sticky='ew')
                row+=1
            ttk.Label(frame,text='Первичный протокол '+('АКЗ №'+role[3:] if role.startswith('AKZ') else 'ВИК' if role=='VIK' else 'УЗК')).grid(row=row,column=0,sticky='w',padx=4,pady=4)
            ttk.Entry(frame,textvariable=fields['path'],width=66).grid(row=row,column=1,sticky='ew')
            def choose_protocol():
                path=filedialog.askopenfilename(parent=window,title='Текущий первичный протокол',filetypes=[('Протоколы','*.pdf *.docx *.xlsx *.txt'),('Все файлы','*.*')])
                if path:fields['path'].set(path);reselected.add(role)
            ttk.Button(frame,text='Выбрать…',command=choose_protocol).grid(row=row,column=2,padx=6)
            control_vars[role]=fields
            return row+1
        controls_frame=ttk.LabelFrame(d,text='Текущий контроль отгрузки',padding=8)
        controls_frame.pack(fill='x',padx=10,pady=6)
        ttk.Label(controls_frame,text='Заключение относится ко всем маркам текущей ВОМ; УЗК — только к маркам с ненулевым планом. При разных результатах выпуск общего комплекта запрещён. Протоколы проверяются перед формированием.',wraplength=1050).grid(row=0,column=0,columnspan=3,sticky='w')
        control_line=control_row(controls_frame,'VIK',1,verdict=True)
        if any(item['uzk_quantity'] for item in uzk_plan.values()):
            control_row(controls_frame,'UZK',control_line,verdict=True)
        extra_vars={}
        for line,key in enumerate(required_doc_fields(current_profile),start=control_line+3):
            var=tk.StringVar(value=self.facts.get('doc_fields',{}).get(key,''))
            ttk.Label(controls_frame,text=DOC_FIELD_LABELS[key]).grid(row=line,column=0,sticky='w',padx=4,pady=3)
            ttk.Entry(controls_frame,textvariable=var,width=66).grid(row=line,column=1,columnspan=2,sticky='ew')
            extra_vars[key]=var
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
                    pair[label]=var
                fields[key]=pair

            dew_status=ttk.Label(
                group,text='Точка росы рассчитывается по температуре воздуха '
                           'и влажности (формула Магнуса).',
                wraplength=1010,justify='left')
            dew_status.grid(row=len(measurements)+2,column=0,
                            columnspan=3,sticky='w',pady=8)
            control_row(group,'AKZ'+str(n),len(measurements)+5,
                        material=form_material(current_profile,n))
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
            if coats==3:
                ttk.Label(group,text='Дополнительные факты трёхслойного акта '
                          '(ЛКМ, партии, цвет, дата и время) — по текущим '
                          'первичным протоколам.',wraplength=1020).grid(
                          row=len(measurements)+3,column=0,columnspan=3,
                          sticky='w')
                detailkeys=('layer','material','paint_batch','hardener_batch',
                            'color','application_date','time_start','time_end',
                            'wet_average')
                detail_frame=ttk.Frame(group)
                detail_frame.grid(row=len(measurements)+4,column=0,
                                  columnspan=3,sticky='ew')
                details=ttk.Treeview(
                    detail_frame,columns=detailkeys,show='headings',height=3)
                for key in detailkeys:
                    details.heading(key,text=key)
                    details.column(key,width=75 if key=='layer' else 150)
                details.pack(fill='x')
                self.editable(details,immutable={0})
                previous={int(r['layer']):r
                          for r in draft.get('layer_details',
                                             prev.get(n,{}).get('layer_details',[]))}
                for layer in (1,2,3):
                    record=previous.get(layer,{})
                    details.insert('','end',values=(layer,)+tuple(
                        record.get(k,'') for k in detailkeys[1:]))
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
        from .automatic_acts import find_approved_geometry,profile_standards
        geometry=find_approved_geometry(self.cfg,require=False)
        standard_cache={}
        if geometry is None:self.write('Нет утверждённого Профили.xlsx: расчёт УЗК будет заблокирован при формировании.')
        def standard_options(item):
            vals=grid.item(item,'values')
            key=str(vals[1])
            if key not in standard_cache:standard_cache[key]=profile_standards(geometry,key) if geometry else []
            return standard_cache[key]
        for item in grid.get_children():
            vals=list(grid.item(item,'values'));options=standard_options(item)
            if not vals[4] and len(options)==1:vals[4]=options[0]
            grid.item(item,values=vals)
        self.editable(grid,immutable={0,3},on_change=refresh_akz,choices={4:standard_options})
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
                previous_controls=dict(self.facts.get('current_controls',{}))
                for role in reselected:previous_controls.pop(role,None)
                draft['current_controls']=save_controls(
                    {role:{key:var.get().strip() for key,var in fields.items()}
                     for role,fields in control_vars.items()},previous_controls)
                draft['doc_fields']={**self.facts.get('doc_fields',{}),**{key:var.get().strip() for key,var in extra_vars.items()}}
                self.facts=draft
                self.cfg.setdefault('geometry_choices',{})[self.project.get()]={r['mark']:r['standard'] for r in draft['marks'] if r.get('standard')}
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

    def editable(self,tree,immutable=set(),on_change=None,choices=None):
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
                values[col]=entry.get()
                tree.item(item,values=values)
                entry.destroy()
                tree._pending_editor_commit=None
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
        return self.check_official()

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
            self.facts=self.shipment_facts()
            current=apply_current_controls(self.facts)
            confirmed_facts=prepare_for_check(current)
        except Stop as exc:
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
                            'approved':True},
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
                 'Подтверждаете выпуск?')
        if not messagebox.askyesno('Предварительный просмотр и выпуск',message):
            self.write('Подтверждение отложено; предварительный PDF доступен. '
                       'Реестр не изменён.')
            return
        self.issue()


def launch():
    root=tk.Tk();ttk.Style(root).theme_use('clam');App(root);root.mainloop()
