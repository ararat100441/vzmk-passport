"""Editable TEST ranges in the existing unified Stage 9 window."""
import tkinter as tk
from tkinter import ttk,messagebox
from uuid import uuid4
from . import akz_auto
from .workspace import save

class AutoAKZPanel:
    def __init__(self,app,notebook):
        self.app=app;self.cfg=None;self.vars={}
        pane=ttk.Frame(notebook,padding=10);notebook.add(pane,text='АКЗ: TEST-диапазоны')
        ttk.Label(pane,text='Климат и толщина моделируются только для TEST. Официальная выдача запрещена.\nОдин климатический набор сохраняется на дату нанесения слоя. Повторный ввод для каждого акта не нужен.',wraplength=840).grid(row=0,column=0,columnspan=4,sticky='w',pady=6)
        self.enabled=tk.BooleanVar(value=False);ttk.Checkbutton(pane,text='Автоматически заполнять TEST-климат и толщины',variable=self.enabled,command=app.invalidate).grid(row=1,column=0,columnspan=4,sticky='w')
        for col,label in enumerate(('Параметр','Минимум','Максимум','Норматив из образца')):ttk.Label(pane,text=label).grid(row=2,column=col,sticky='w',padx=6,pady=6)
        keys=[('air_c','Температура воздуха, °C'),('surface_c','Температура поверхности, °C'),('relative_humidity','Влажность, %')]
        keys += [(f'{n}:{i}',f'АКЗ №{n} / слой {i+1}, мкм') for n,count in ((0,1),(1,1),(2,2),(3,3)) for i in range(count)]
        self.norms={}
        for row,(key,label) in enumerate(keys,3):
            ttk.Label(pane,text=label).grid(row=row,column=0,sticky='w',padx=6,pady=3)
            pair=[tk.StringVar(),tk.StringVar()];self.vars[key]=pair
            for col,var in enumerate(pair,1):
                ttk.Entry(pane,textvariable=var,width=12).grid(row=row,column=col,padx=6,pady=3)
                var.trace_add('write',app.invalidate)
            norm=ttk.Label(pane,text='');norm.grid(row=row,column=3,sticky='w',padx=6);self.norms[key]=norm
        ttk.Label(pane,text='Точка росы: Магнус (17,62; 243,12), расчёт по воздуху и влажности.\nКлимат: шаг 0,1; толщина: целые мкм. Поверхность ≥ Td + 3 °C. Климат даты не перезаписывается.\nУ №2 второй слой: индивидуальная норма не задана; сохраняется ∑1+2=180. Его диапазон — только TEST-параметр.',wraplength=840).grid(row=13,column=0,columnspan=4,sticky='w',pady=8)
        ttk.Button(pane,text='Сохранить TEST-настройки',command=self.store).grid(row=14,column=0,columnspan=3,sticky='w',pady=6)
        self.notice=tk.StringVar(value='Откройте DEV-стенд.');ttk.Label(pane,textvariable=self.notice,wraplength=840).grid(row=15,column=0,columnspan=4,sticky='w')
    def refresh(self):
        if not self.app.workspace:return
        try:
            cfg=akz_auto.read_settings(self.app.workspace,required=False)
            self.cfg=cfg or (akz_auto.defaults(self.app.workspace)|{'enabled':False})
            self.enabled.set(self.cfg['enabled']);p=akz_auto.profiles(self.app.workspace)
            for key,pair in self.vars.items():
                if ':' in key:
                    n,i=key.split(':');value=self.cfg['thickness'][n][int(i)];layer=p[n]['layers'][int(i)]
                    self.norms[key].config(text=str(layer['required_raw'])+' мкм')
                else:value=self.cfg['climate'][key]
                for var,item in zip(pair,value):var.set(str(item))
            self.notice.set('Настройки сохранены отдельно от шаблонов и реестров.' if cfg else 'Показаны примерные TEST-диапазоны. Сохраните их для включения моделирования.')
        except Exception as exc:self.notice.set('ОШИБКА: '+str(exc));self.cfg=None
    def store(self):
        if self.app.busy or not self.app.workspace or self.cfg is None:return
        try:
            import copy
            cfg=copy.deepcopy(self.cfg);cfg['enabled']=bool(self.enabled.get())
            for key,pair in self.vars.items():
                from .akz_settings import gui_number
                values=[gui_number(var.get()) for var in pair]
                if ':' in key:
                    n,i=key.split(':');cfg['thickness'][n][int(i)]=values
                else:cfg['climate'][key]=values
            akz_auto.validate_settings(cfg,akz_auto.profiles(self.app.workspace))
            from .gui import config_path
            incoming=config_path().parent/'diagnostics'/('akz_settings_'+uuid4().hex+'.json');save(incoming,cfg)
            self.app.invalidate()
            self.app.work('akz-settings-save',['--akz-settings',str(incoming)],lambda result:self.refresh())
        except Exception as exc:messagebox.showerror('Настройки АКЗ',str(exc),parent=self.app.window)
