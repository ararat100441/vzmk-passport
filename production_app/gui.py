import os,subprocess,sys,traceback
from pathlib import Path
from uuid import uuid4
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from .storage import local_root,read,save,sha,Stop
from . import model,core,akz_presets,VERSION
from .real_forms import DOC_FIELD_LABELS, UZK_COLUMNS, UZK_HEADERS

class App:
    def __init__(self,root):
        self.root=root;root.title('ВЗМК — Паспортизация '+VERSION);self.fit(root,1050,760);root.minsize(min(950,root.winfo_screenwidth()-40),min(650,root.winfo_screenheight()-80))
        config_error=None
        try:self.cfg=read(local_root()/'settings.json') if (local_root()/'settings.json').exists() else {'profiles':{}}
        except Exception as exc:self.cfg={'profiles':{}};config_error=str(exc)
        self.plan=None;self.facts={};self.active=None;self.result_path=None;self.buttons=[]
        self.preview_pdf=None
        self.first_reference=None  # read-only first issued passport, never copied as facts
        self.project=tk.StringVar();self.number=tk.StringVar();self.vom=tk.StringVar();self.new=tk.BooleanVar(value=False)
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
        ttk.Button(form,text='Выбрать ВОМ…',command=lambda:self.choose(self.vom,False)).grid(row=2,column=2)
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
        ttk.Label(root,text='TEST/SAMPLE_ONLY не являются результатами контроля. Выпуск требует протоколов, проверки Office и отдельного подтверждения.',wraplength=940).pack(side='bottom',padx=14,pady=(0,12))
        self.log=tk.Text(root,wrap='word',font=('Segoe UI',10),state='disabled',height=10);self.log.pack(fill='both',expand=True,padx=14,pady=12)
        root.protocol('WM_DELETE_WINDOW',self.close)
        self.invalidate();self.refresh()
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

    def error(self,exc):self.write(str(exc));messagebox.showerror('Остановлено — исходники не исправляются автоматически',str(exc))

    def choose(self,var,directory):
        if self.active:return
        path=filedialog.askdirectory() if directory else filedialog.askopenfilename(filetypes=[('Excel/Word','*.xlsx *.xlsm *.docx'),('Все файлы','*.*')])
        if path:var.set(path)

    def settings(self):
        if self.active:return
        from .path_config import visible_settings,apply_settings
        shown=visible_settings(self.cfg)
        dialog=tk.Toplevel(self.root)
        dialog.title('Рабочие пути — три источника')
        self.fit(dialog,1060,380)
        dialog.transient(self.root)
        dialog.grab_set()
        selected={key:tk.StringVar(value=value)
                  for key,value in shown.items()}
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
            'при обычной работе. УЗК.xlsm используется только для чтения.',
            wraplength=1010,
        ).grid(row=3,column=0,columnspan=3,padx=14,pady=8)

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
                self.cfg=cfg
                save(local_root()/'settings.json',cfg)
                self.invalidate();self.facts={}
                dialog.destroy();self.refresh()
            except Exception as exc:self.error(exc)

        ttk.Button(dialog,text='Сохранить рабочие пути',
                   command=store).grid(row=4,column=1,sticky='w',pady=12)
        ttk.Button(dialog,text='Создать учебную копию',
                   command=lambda:store(True)).grid(
                   row=4,column=1,sticky='e',pady=12)
        def open_recovery():
            dialog.destroy()
            self.recovery()
        ttk.Button(dialog,text='Восстановление после сбоя',
                   command=open_recovery).grid(
                   row=5,column=1,sticky='w',pady=6)

    def refresh(self):
        if self.active:return
        from .shared_uzk import find_shared_uzk
        try:uzk=find_shared_uzk(self.cfg,require=False)
        except Stop:uzk=None
        mode=('УЧЕБНАЯ КОПИЯ' if self.cfg.get('training')
              else 'РАБОЧИЙ РЕЖИМ — запись только после подтверждения')
        self.paths.configure(
            text=mode+'\n'
            'Реестр: '+str(self.cfg.get('registry') or 'не выбран')+'\n'
            'Документы и паспорта: '+str(self.cfg.get('output') or 'не выбрана')+'\n'
            'Общий УЗК: '+str(uzk or 'не выбран'),
        )
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
        self.invalidate();self.facts={};self.first_reference=None
        try:
            r=model.read_registry(self.cfg['registry']);name=self.project.get();entries=r.projects[name]
            self.numbers['values']=[str(e['number']) for e in entries]
            self.number.set(str(entries[0]['number']))
            p=model.select(r,name)
            # The FIRST ISSUED passport is the automatic reference for THIS
            # project only. Existing approved empty forms remain separate:
            # never inject historical findings or signatures as new facts.
            from .first_passport import discover_first_passport_for_config
            self.cfg['auto_first_passport']=True
            try:
                self.first_reference=discover_first_passport_for_config(
                    self.cfg,p)
                refs=self.first_reference['sources']
                self.write('Автоматически найден первый паспорт: '+
                           p.first_number+' / '+
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
                p={'root':str(model.contained(root.get(),self.cfg['documents'])),'templates':{}}
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
                self.cfg=cfg;save(local_root()/'settings.json',cfg);self.facts={};self.invalidate();d.destroy()
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
        d=tk.Toplevel(self.root);d.title('Фактические результаты — '+self.project.get()+' '+self.number.get());self.fit(d,1160,760)
        d.transient(self.root);d.grab_set()
        window=d;window.bind('<Escape>',lambda e:window.destroy());d=self.scroll_content(window)
        ttk.Label(d,text=(
            'Реквизиты и даты берутся из реестра ВЗМК. '
            'Для нового паспорта его текущие даты должны быть '
            'получены по утверждённым формулам/источникам, '
            'старые даты первого паспорта не копируются.'
        ),wraplength=1090).pack(anchor='w',padx=10,pady=6)
        ttk.Label(d,text='Марка и расчётное количество УЗК заполняются из ВОМ и общего реестра. Профиль и № АКЗ — из справочника проекта. Редактирование: двойной щелчок. Старые заключения контроля автоматически не назначаются.').pack(pady=4)
        cols=('mark','profile','akz','uzk_quantity')
        headings={'mark':'Марка','profile':'Профиль',
                  'akz':'№ АКЗ','uzk_quantity':'УЗК, шт'}
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
        def import_assignments():
            try:
                profile=model.profile(self.cfg,self.project.get())
                path=profile.get('catalog')
                if not path:
                    raise Stop('Справочник не найден у выбранного проекта.')
                mapping=current_assignments(
                    path,profile['catalog_sha256'],
                    (grid.item(item,'values')[0] for item in grid.get_children()))
                for item in grid.get_children():
                    vals=list(grid.item(item,'values'))
                    source=mapping.get(str(vals[0]).strip(),{})
                    for key in ('profile','paint','akz'):
                        if key not in source:continue
                        if key=='paint':
                            hidden_fields[vals[0]]['paint']=source[key]
                        else:
                            vals[cols.index(key)]=str(source[key])
                    grid.item(item,values=vals)
                refresh_akz()
                self.write('Назначения перечитаны из '+path)
            except Exception as exc:self.error(exc)
        ttk.Button(d,text='Обновить назначения из справочника',
                   command=import_assignments).pack(pady=3)
        # Show only variants actually assigned by the current VOM/catalogue.
        # If the project has ONE approved AKZ form but marks are unassigned,
        # show its blank ranges as a convenience — do not assign AKZ to marks.
        from .akz_ui import available_variants,display_variants,source_coats
        from .measurement_ranges import measurements_from_ranges
        current_profile=active_profile if 'active_profile' in locals() else (
            self.cfg.get('profiles',{}).get(self.project.get(),{}))
        variants=available_variants(current_profile)
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
                    'humidity':'Влажность, %','dew':'Точка росы, °C'}
            measurements=list(labels.items())+[
                ('layer_'+str(i),'Фактическая толщина слоя '+str(i)+', мкм')
                for i in range(1,(coats or 0)+1)]
            for line,(key,title) in enumerate(measurements,start=2):
                stored=(old_interval.get('layers',[]) if key.startswith('layer_')
                        else old_interval.get(key))
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
            # Detailed 3-coat material batches remain accessible ONLY when
            # an actual approved 3-layer form is selected. They are never
            # shown for the one-layer Kazakhstan act.
            details=None
            if coats==3:
                ttk.Label(group,text='Дополнительные факты трёхслойного акта '
                          '(ЛКМ, партии, цвет, дата и время) — по текущим '
                          'первичным протоколам.',wraplength=1020).grid(
                          row=len(measurements)+2,column=0,columnspan=3,
                          sticky='w')
                detailkeys=('layer','material','paint_batch','hardener_batch',
                            'color','application_date','time_start','time_end',
                            'wet_average')
                detail_frame=ttk.Frame(group)
                detail_frame.grid(row=len(measurements)+3,column=0,
                                  columnspan=3,sticky='ew')
                details=ttk.Treeview(
                    detail_frame,columns=detailkeys,show='headings',height=3)
                for key in detailkeys:
                    details.heading(key,text=key)
                    details.column(key,width=75 if key=='layer' else 150)
                details.pack(fill='x')
                self.editable(details,immutable={0})
                previous={int(r['layer']):r
                          for r in prev.get(n,{}).get('layer_details',[])}
                for layer in (1,2,3):
                    record=previous.get(layer,{})
                    details.insert('','end',values=(layer,)+tuple(
                        record.get(k,'') for k in detailkeys[1:]))
            attachments=ttk.Frame(group)
            attachments.grid(row=len(measurements)+5,column=0,
                             columnspan=3,sticky='w',pady=5)
            count_label=ttk.Label(
                attachments,text='Протоколов АКЗ №'+str(n)+': '+
                                 str(len(akz_proofs[n])))
            count_label.pack(side='left',padx=5)
            def pick_akz_proof(variant=n,display=count_label):
                try:
                    for raw in filedialog.askopenfilenames(
                            title='Первичный протокол АКЗ №'+str(variant)):
                        path=str(Path(raw).resolve())
                        proof={'path':path,'sha256':sha(path),
                               'role':'AKZ','variant':variant}
                        if not any(x.get('path')==path
                                   for x in akz_proofs[variant]):
                            akz_proofs[variant].append(proof)
                        if not any(x.get('path')==path for x in proofs):
                            proofs.append(proof)
                    display.configure(text='Протоколов АКЗ №'+str(variant)+
                                      ': '+str(len(akz_proofs[variant])))
                except Exception as exc:self.error(exc)
            ttk.Button(attachments,text='Приложить протокол АКЗ…',
                       command=pick_akz_proof).pack(side='left',padx=5)
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
        self.editable(grid,immutable={0,3},on_change=refresh_akz)
        refresh_akz()

        # Actual per-weld inspection evidence is entered in this existing
        # screen when required by the approved detailed UZK form.
        # No separate CSV/second workbook creation is needed.
        ttk.Label(
            d,text='УЗК: объём рассчитан по общему УЗК.xlsm. '
                   'Сведения о реально проверенных соединениях (сварщик, '
                   'дата, прибор, оценка и протокол) вводите здесь только '
                   'по текущим первичным результатам, если этого требует '
                   'утверждённый акт. Не переносите заключения первого паспорта.',
            wraplength=1080,
        ).pack(padx=12,pady=5)
        details_frame=ttk.Frame(d)
        details_frame.pack(fill='x',padx=10)
        uzk_details=ttk.Treeview(
            details_frame,columns=UZK_COLUMNS,show='headings',height=5)
        for key,label in zip(UZK_COLUMNS,UZK_HEADERS):
            uzk_details.heading(key,text=label)
            uzk_details.column(
                key,width=95 if key in ('mark','date','frequency_mhz',
                                        'angle_deg') else 175)
        uzk_scroll=ttk.Scrollbar(details_frame,orient='horizontal',
                                 command=uzk_details.xview)
        uzk_details.configure(xscrollcommand=uzk_scroll.set)
        uzk_details.pack(fill='x',expand=True)
        uzk_scroll.pack(fill='x')
        self.editable(uzk_details)
        for record in self.facts.get('uzk_details',[]):
            uzk_details.insert('','end',
                               values=tuple(record.get(k,'')
                                            for k in UZK_COLUMNS))
        action_frame=ttk.Frame(d)
        action_frame.pack(anchor='w',padx=12,pady=4)
        def add_uzk_record():
            uzk_details.insert('','end',values=('',)*len(UZK_COLUMNS))
        def remove_uzk_record():
            for item in uzk_details.selection():
                uzk_details.delete(item)
        ttk.Button(action_frame,text='Добавить фактическое соединение',
                   command=add_uzk_record).pack(side='left',padx=4)
        ttk.Button(action_frame,text='Удалить выбранное',
                   command=remove_uzk_record).pack(side='left',padx=4)
        ttk.Label(d,text='Дополнительные реквизиты утверждённых форм (ВИК, качество, УЗК): заполните только необходимые поля по подтверждённым источникам. Старые сведения не подставляются.',wraplength=1080).pack(pady=5)
        extra_columns=('key','description','value')
        extra_frame=ttk.Frame(d);extra_frame.pack(fill='both',padx=10,pady=4)
        extra=ttk.Treeview(extra_frame,columns=extra_columns,show='headings',height=8)
        for col,heading,width in [('key','Поле',210),('description','Наименование',430),('value','Подтверждённое значение',470)]:
            extra.heading(col,text=heading);extra.column(col,width=width)
        extra_scroll=ttk.Scrollbar(extra_frame,orient='vertical',command=extra.yview)
        extra.configure(yscrollcommand=extra_scroll.set)
        extra_scroll.pack(side='right',fill='y');extra.pack(fill='both',expand=True)
        self.editable(extra,immutable={0,1})
        previous_fields=self.facts.get('doc_fields',{})
        for key,label in DOC_FIELD_LABELS.items():
            extra.insert('','end',values=(key,label,previous_fields.get(key,'')))
        proofs=list(self.facts.get('evidence',[]));proof_label=ttk.Label(d,text='Протоколов: '+str(len(proofs)));proof_label.pack()
        def add():
            for f in filedialog.askopenfilenames(title='Первичные протоколы фактического контроля'):
                proofs.append({'path':str(Path(f).resolve()),'sha256':sha(f)})
            proof_label.configure(text='Протоколов: '+str(len(proofs)))
        ttk.Button(d,text='Приложить первичные протоколы…',command=add).pack()
        approved=tk.BooleanVar(value=self.facts.get('confirmed',False));rights=tk.BooleanVar(value=self.facts.get('authorized',False));forms=tk.BooleanVar(value=self.facts.get('forms_approved',False))
        for variable,text in [(approved,'Подтверждаю фактические результаты; это не TEST/SAMPLE_ONLY'),(rights,'Подтверждаю право на выпуск и достоверность данных'),(forms,'Формы документов утверждены ответственным лицом')]:ttk.Checkbutton(d,text=text,variable=variable).pack(anchor='w',padx=12)
        def store():
            try:
                data={
                    'project':self.project.get(),'number':self.number.get(),
                    'kind':'TRAINING' if self.cfg.get('training') else 'ACTUAL',
                    'confirmed':approved.get(),'authorized':rights.get(),
                    'forms_approved':forms.get(),'evidence':proofs,
                    'marks':[],'climates':[],
                    'auto_registry_dates':True,
                    'no_responsible_input':True,
                    'responsible':'',
                }
                for item in grid.get_children():
                    shown=dict(zip(cols,grid.item(item,'values')))
                    mark=str(shown['mark'])
                    row={**hidden_fields.get(mark,{}),**shown}
                    raw=str(row['akz']).strip()
                    row['akz']=int(raw) if raw else None
                    # Empty results stay unconfirmed. Removing table columns
                    # is never permission to silently set VIK/UZK to ГОДЕН.
                    data['marks'].append(row)
                used={r['akz'] for r in data['marks'] if r['akz'] is not None}
                for variant in sorted(used):
                    if variant not in variant_forms:
                        raise Stop('Для АКЗ №'+str(variant)+
                                   ' нет утверждённой рабочей формы.')
                    info=variant_forms[variant]
                    coats=info['coats']
                    if coats not in (1,3):
                        raise Stop('Нет подтверждённого формата 1/3 слоя для АКЗ №'+
                                   str(variant))
                    ranges={}
                    for key in ('air','surface','humidity','dew'):
                        ranges[key]={edge:info['fields'][key][edge].get().strip()
                                     for edge in ('min','max')}
                    ranges['layer_intervals']=[]
                    for index in range(1,coats+1):
                        key='layer_'+str(index)
                        row={'layer':index}
                        row.update({edge:info['fields'][key][edge].get().strip()
                                    for edge in ('min','max')})
                        ranges['layer_intervals'].append(row)
                    climate=measurements_from_ranges(
                        variant,ranges,coats=coats)
                    colors=sorted({str(mark.get('paint','')).strip()
                                   for mark in data['marks']
                                   if mark.get('akz')==variant and
                                      str(mark.get('paint','')).strip()})
                    climate['material']=', '.join(colors)
                    # A protocol reference is never generated from a filename
                    # unless the operator actually attached that real file.
                    # Only proof files explicitly selected for this АКЗ
                    # variant can supply the coating's protocol reference.
                    climate['basis']='; '.join(
                        Path(proof['path']).name for proof in akz_proofs[variant])
                    details=info['details']
                    if details is not None:
                        detailkeys=('layer','material','paint_batch','hardener_batch',
                                    'color','application_date','time_start',
                                    'time_end','wet_average')
                        rows=[]
                        for item in details.get_children():
                            vals=dict(zip(detailkeys,details.item(item,'values')))
                            if any(str(vals[k]).strip() for k in detailkeys[1:]):
                                vals['layer']=int(vals['layer'])
                                rows.append(vals)
                        if rows:climate['layer_details']=rows
                    data['climates'].append(climate)
                data['akz_proofs']={str(n):list(akz_proofs[n])
                                    for n in used if akz_proofs[n]}
                data['uzk_details']=[
                    dict(zip(UZK_COLUMNS,uzk_details.item(item,'values')))
                    for item in uzk_details.get_children()
                ]
                data['doc_fields']={}
                for item in extra.get_children():
                    key,_,value=extra.item(item,'values')
                    if str(value).strip():data['doc_fields'][key]=str(value).strip()
                from .facts import validate
                validate(data,v,self.project.get(),self.number.get(),
                         self.cfg.get('training',False),
                         defer_registry_dates=True)
                self.facts=data;self.invalidate();window.destroy();self.write('Фактические данные сохранены локально в текущей заявке. Исходники не изменены.')
            except Exception as exc:self.error(exc)
        ttk.Button(d,text='Проверить и принять введённые данные',command=store).pack(pady=10)

    def editable(self,tree,immutable=set(),on_change=None):
        def edit(event):
            item=tree.identify_row(event.y);column=tree.identify_column(event.x)
            if not item or not column:return
            col=int(column[1:])-1
            if col in immutable:return
            x,y,w,h=tree.bbox(item,column);entry=ttk.Entry(tree);entry.place(x=x,y=y,width=w,height=h)
            entry.insert(0,tree.item(item,'values')[col]);entry.focus_set()
            def done(e=None):
                if not entry.winfo_exists():return
                values=list(tree.item(item,'values'));values[col]=entry.get();tree.item(item,values=values);entry.destroy()
                if on_change:on_change()
            entry.bind('<Return>',done);entry.bind('<FocusOut>',done)
        tree.bind('<Double-1>',edit)

    def catalog(self):
        if self.active:return
        target=filedialog.asksaveasfilename(title='НОВЫЙ черновик справочника',defaultextension='.xlsx',initialfile='Черновик справочника марок.xlsx')
        if target:self.run('catalog',{'vom':self.vom.get(),'target':target},lambda r:self.write('Черновик: '+r['path']+'; профиль/краска/АКЗ пустые.'))

    def check(self):
        if self.active:return
        if self.plan and self.preview_pdf:
            self.review_preview()
            return
        if not all(self.cfg.get(k) for k in ('registry','documents','output')):
            self.error('Сначала выберите рабочие пути в «Настройки путей».');return
        if not self.project.get() or not self.number.get() or not self.vom.get():
            self.error('Выберите объект, паспорт и отгрузочную ВОМ.');return
        if not self.facts:
            self.error('Откройте «Фактические результаты» и проверьте сведения '
                       'текущей отгрузки. Старые протоколы не подставляются.')
            return
        self.run('check',{'cfg':self.cfg,'project':self.project.get(),
                          'number':self.number.get(),'vom':self.vom.get(),
                          'facts':self.facts,'new':self.new.get()},self.checked)

    def checked(self,result):
        self.plan=result['plan']
        self.facts=result['data']['facts']
        d=result['data']
        self.write(f"Данные проверены: {d['project']} | {d['number']} | "
                   f"{d['marks']} марок / {d['quantity']} изделий / "
                   f"{d['mass_kg']} кг. Дата паспорта: {d['date']}.")
        self.run('prepare',{'path':self.plan},self.prepared)

    def prepared(self,result):
        self.write('Word и Excel подготовлены: '+str(result['files'])+' файлов.')
        self.run('verify',{'path':self.plan,'preview':True},self.verified)

    def verified(self,result):
        self.preview_pdf=result.get('preview_pdf')
        if not self.preview_pdf or not Path(self.preview_pdf).is_file():
            self.error('Не создан черновой PDF: выпуск запрещён.')
            return
        self.buttons[0].configure(text='Просмотреть и подтвердить')
        self.review_preview()

    def review_preview(self):
        if not self.preview_pdf or not self.plan:return
        try:
            d=read(self.plan)
            os.startfile(self.preview_pdf)
        except Exception as exc:
            self.error('Не удалось открыть подготовленный PDF: '+str(exc))
            return
        mode='УЧЕБНЫЙ РЕЖИМ' if self.cfg.get('training') else 'РАБОЧИЙ ВЫПУСК'
        message=(mode+'\n'
                 'Проект: '+str(d['project'])+'\n'
                 'Номер: '+str(d['number'])+'\n'
                 'Дата паспорта: '+str(d['date'])+'\n'
                 'Реестр: '+str(d['registry'])+'\n'
                 'Папка результата: '+str(d['target'])+'\n'+
                 ('Будет добавлена одна строка в реестр.\n' if d['new']
                  else 'Новых строк реестра не будет.\n')+
                 '\nОткрылся предварительный PDF с пометкой ЧЕРНОВИК. '
                 'Проверьте ВСЕ страницы, марки, фактические результаты '
                 'контроля, материалы, толщины и даты.\n\n'
                 'Это ЕДИНСТВЕННОЕ подтверждение официального выпуска. '
                 'После «Да» будут подготовлены итоговые документы, '
                 'выполнены проверки и запись в реестр с резервной копией.\n'
                 'Подтверждаете выпуск?')
        if not messagebox.askyesno('Предварительный просмотр и выпуск',message):
            self.write('Подтверждение отложено; предварительный PDF доступен. '
                       'Реестр не изменён.')
            return
        self.issue()

    def open_documents(self):
        if self.plan:
            folder=Path(self.plan).parent/'documents'
            if folder.is_dir():os.startfile(folder)
    def issue(self):
        if not self.plan or not self.preview_pdf:return
        # Only entered from explicit consent in review_preview(), with exact
        # project/date/registry/output shown. Do not ask a SECOND approval.
        self.run('confirm',{'path':self.plan,
                            'responsible':self.facts.get('responsible',''),
                            'reviewed':True},
                 lambda r:self.run('release',{'path':self.plan},self.released))
    def released(self,r):self.run('commit',{'path':self.plan,'release_result':self.result_path},self.committed)
    def committed(self,r):self.run('readback',{'journal':r['journal']},lambda v:self.run('complete',{'journal':r['journal'],'readback_result':self.result_path},self.completed))
    def completed(self,r):
        self.write('ГОТОВО. '+str(r))
        # Store preferences only for an actually completed production issue.
        # Old protocols, measurement thicknesses, dates and approvals never
        # become facts of the next passport.
        if r.get('state')=='ГОТОВО' and not self.cfg.get('training') and self.facts.get('kind')=='ACTUAL':
            try:
                if akz_presets.remember(self.project.get(),self.facts.get('climates',[]),self.facts.get('akz_date')):
                    self.write('Последние климатические показатели и диапазоны слоёв сохранены только для объекта '+self.project.get()+'.')
            except Exception as exc:
                self.write('Выпуск завершён, но сохранить подсказки АКЗ не удалось: '+str(exc))
        messagebox.showinfo('Проверенный результат',str(r));self.invalidate();self.refresh()
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
        for widget in (self.projects,self.numbers,self.new_widget,self.vom_widget):widget.configure(state='disabled')
        states=[b['state'] for b in self.buttons]
        for b in self.buttons:b.configure(state='disabled')
        self.progress.configure(text='Выполняется: '+task+'. Дождитесь штатного закрытия Office.');self.root.after(500,lambda:self.poll(root,task,callback,states))
    def poll(self,root,task,callback,states):
        if self.active.poll() is None:self.root.after(500,lambda:self.poll(root,task,callback,states));return
        code=self.active.returncode;self.active=None
        self.projects.configure(state='readonly');self.numbers.configure(state='readonly');self.new_widget.configure(state='normal');self.vom_widget.configure(state='normal')
        try:
            r=read(root/'result.json');self.result_path=str(root/'result.json')
            if code or r['status']!='PASS':raise Stop(str(r.get('error',r))+'\nЖурнал: '+str(root))
            for b,state in zip(self.buttons,states):b.configure(state=state)
            self.progress.configure(text='PASS: '+task);callback(r)
        except Exception as exc:self.invalidate();self.progress.configure(text='FAIL: '+task);self.error(exc)

def launch():
    root=tk.Tk();ttk.Style(root).theme_use('clam');App(root);root.mainloop()
