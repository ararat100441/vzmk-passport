"""Simple isolated DEV UI, with a frozen preview and explicit confirmation."""
from datetime import datetime
import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from uuid import uuid4
from .storage import ROOT, Stop
from . import core, recovery


class App:
    def __init__(self, window):
        self.window = window
        self.plan = None
        self.restore_plan = None
        self.sandbox = None
        self.request = None
        window.title('ВЗМК — Этап 8 — безопасное закрытие DEV / TEST')
        window.geometry('1000x760')
        tk.Label(window,text='ИЗОЛИРОВАННЫЙ TEST. Официальное закрытие и реальные списания запрещены.',
                 fg='#9a271d',wraplength=950,font=('Segoe UI',11,'bold')).pack(pady=10)
        bar = tk.Frame(window);bar.pack(fill='x',padx=12)
        tk.Button(bar,text='Создать искусственный TEST-стенд',command=self.create).pack(side='left',padx=3)
        tk.Button(bar,text='Выбрать запрос / manifest Этапа 7',command=self.choose).pack(side='left',padx=3)
        self.text = tk.Text(window,wrap='word',font=('Segoe UI',11),height=27)
        self.text.pack(fill='both',expand=True,padx=12,pady=10)
        self.show('Выберите точный CLOSE_REQUEST.json тестового стенда или manifest Этапа 7.\n\n'
                  'Паспорт 9/2026 с SAMPLE_ONLY и TEST-контролем закрывать нельзя.\n'
                  'Искусственный стенд использует отдельную копию XLSM и вымышленные подтверждения.')
        self.confirm = tk.BooleanVar(value=False)
        tk.Checkbutton(window,text='Я проверил проект, паспорт, отгрузку, документы и остатки. Подтверждаю закрытие TEST.',
                       variable=self.confirm).pack(pady=5)
        self.close_button = tk.Button(window,text='Закрыть паспорт',state='disabled',command=self.close)
        self.close_button.pack(pady=5)
        self.recover_button = tk.Button(window,text='Проверить / восстановить незавершённую операцию',command=self.recover)
        self.recover_button.pack(pady=5)
        self.restore_button = tk.Button(window,text='Восстановить из резервной копии',state='disabled',command=self.restore)
        self.restore_button.pack(pady=5)

    def show(self, text):
        self.text.config(state='normal');self.text.delete('1.0','end');self.text.insert('1.0',text);self.text.config(state='disabled')

    def reset(self):
        self.plan = None;self.confirm.set(False);self.close_button.config(state='disabled')
        self.restore_plan = None;self.restore_button.config(state='disabled')

    def load(self, path):
        from .storage import read
        self.reset();self.sandbox=None;self.request=None
        q=read(path)
        if q.get('schema')=='STAGE7_SELECTED_TEST_MANIFEST_V1':
            result=core.inspect_stage7(path)
            lines=[f"Проект: {result['project']}",f"Паспорт: {result['passport']}",
                f"Выбранная отгрузка: {result['shipment']}",'Документы:']
            lines += [f"  {d['kind']} — {d['status']}" for d in result['documents']]
            lines += [f"Количество изделий УЗК: {result['controlled_units']}",'Статус проверки: STOP',
                result['reason'],result['balances'],'Обязательные официальные документы и фактический контроль не подтверждены.']
            self.show('\n'.join(lines));return
        self.request=Path(path).resolve();self.sandbox=self.request.parents[1]
        try:self.plan=core.preview(self.sandbox,self.request)
        except Stop:
            from .storage import history
            if history(self.sandbox)['active'] is None:raise
            self.pending();return
        self.show(core.describe(self.plan))
        if self.plan['payload']['status']=='READY_TEST':self.close_button.config(state='normal')

    def choose(self):
        p=filedialog.askopenfilename(title='Точный запрос Stage 8 или manifest Stage 7',filetypes=[('JSON','*.json')])
        if not p:return
        try:self.load(p)
        except Exception as exc:self.reset();self.show('STOP: '+str(exc));messagebox.showerror('STOP',str(exc))

    def create(self):
        from .fixtures import create
        self.reset()
        try:
            parent=ROOT/'validation'/('stage8_ui_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid4().hex[:6])/'sandboxes'
            result=create(parent);self.load(result['request'])
        except Exception as exc:self.show('STOP: '+str(exc));messagebox.showerror('STOP',str(exc))

    def close(self):
        try:
            if not self.plan or not self.confirm.get():raise Stop('Требуется явное подтверждение после проверки.')
            if not messagebox.askyesno('Подтвердить закрытие TEST',core.describe(self.plan)+'\n\nЗакрыть этот тестовый паспорт?',default='no'):
                return
            result=core.close(self.plan,core.confirmation(self.plan))
            self.reset();self.show('TEST закрыт. Повторного списания не будет.\nОперация: '+result['operation']['operation_id'])
        except Exception as exc:self.reset();self.show('STOP: '+str(exc));messagebox.showerror('STOP',str(exc))

    def recover(self):
        try:
            if self.sandbox is None:raise Stop('Сначала выберите запрос из тестового стенда.')
            result=core.recover(self.sandbox)
            self.reset();self.show(json.dumps(result,ensure_ascii=False,indent=2))
        except Exception as exc:
            self.reset()
            if self.sandbox is not None:self.pending(str(exc))
            else:self.show('STOP: '+str(exc));messagebox.showerror('STOP',str(exc))

    def pending(self, reason=''):
        self.reset()
        try:
            self.restore_plan=recovery.preview(self.sandbox)
            self.show(recovery.describe(self.restore_plan));self.restore_button.config(state='normal')
        except Exception as exc:
            self.show('Обнаружена незавершённая операция / STOP.\n'+reason+'\n'
                      'Восстановление из backup недоступно: '+str(exc)+'\n'
                      'Закрытие и повторное списание запрещены. Кнопка проверки журнала доступна.')

    def restore(self):
        try:
            if not self.restore_plan:raise Stop('Сначала проверьте аварийное состояние и backup.')
            p=self.restore_plan
            if not messagebox.askyesno('Подтвердить восстановление TEST',
                    recovery.describe(p)+'\n\nВосстановить этот реестр из указанного backup?',default='no'):
                return
            result=recovery.restore(p,recovery.confirmation(p))
            self.reset();self.show('Восстановление успешно проверено native Excel.\n'
                'Статус: '+result['status']+'\nПаспорт не закрыт. Списаний: 0.\n'
                'Заново выберите запрос для нового preview и нового подтверждения.\n'
                'Evidence: '+result['native_evidence'])
        except Exception as exc:
            self.reset();self.show('STOP: '+str(exc));messagebox.showerror('STOP',str(exc))


def launch():
    root=tk.Tk();App(root);root.mainloop()
