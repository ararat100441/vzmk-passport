"""Small DEV UI: selection -> read-only preview -> verified TEST document."""
import os
from pathlib import Path
from queue import Empty, Queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from profiles_stage4.model import fmt
from .core import generate, preview
from .inputs import ROOT, passports, verify_pins


class App:
    def __init__(self, root):
        self.root, self.plan, self.output = root, None, None
        self.queue = Queue()
        self.busy = False
        root.title('Этап 5 DEV — УЗК (ТЕСТ / ЧЕРНОВИК)')
        root.geometry('1050x640')
        root.minsize(880, 520)
        self.project = tk.StringVar()
        self.passport = tk.StringVar()
        self.vom = tk.StringVar()
        self.checked = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value='Выберите проект, паспорт и ВОМ. Источники читаются без изменения.')
        panel = ttk.Frame(root, padding=14)
        panel.pack(fill='both', expand=True)
        ttk.Label(panel, text='ТЕСТ / ОБРАЗЕЦ — результаты контроля не подтверждены', font=('Segoe UI', 13, 'bold')).pack(anchor='w')
        form = ttk.Frame(panel)
        form.pack(fill='x', pady=12)
        ttk.Label(form, text='Проект').grid(row=0, column=0, sticky='w')
        self.project_box = ttk.Combobox(form, textvariable=self.project, state='readonly', width=38)
        self.project_box.grid(row=0, column=1, sticky='ew', padx=8)
        ttk.Label(form, text='Паспорт').grid(row=0, column=2)
        self.passport_box = ttk.Combobox(form, textvariable=self.passport, state='readonly', width=18)
        self.passport_box.grid(row=0, column=3, padx=8)
        ttk.Label(form, text='ВОМ XLS / XLSX').grid(row=1, column=0, sticky='w', pady=10)
        ttk.Entry(form, textvariable=self.vom, state='readonly').grid(row=1, column=1, columnspan=3, sticky='ew', padx=8)
        self.select_button = ttk.Button(form, text='Выбрать…', command=self.select_vom)
        self.select_button.grid(row=1, column=4)
        form.columnconfigure(1, weight=1)
        buttons = ttk.Frame(panel)
        buttons.pack(fill='x', pady=(0, 10))
        self.preview_button = ttk.Button(buttons, text='Предварительный расчёт', command=self.do_preview)
        self.preview_button.pack(side='left')
        ttk.Checkbutton(buttons, text='Марки и количества проверены', variable=self.checked, command=self.update_buttons).pack(side='left', padx=12)
        self.generate_button = ttk.Button(buttons, text='Создать ТЕСТ Word', command=self.do_generate, state='disabled')
        self.generate_button.pack(side='left')
        self.open_button = ttk.Button(buttons, text='Открыть Word', command=self.open_word, state='disabled')
        self.open_button.pack(side='left', padx=8)
        self.details = tk.Text(panel, wrap='word', font=('Consolas', 11), state='disabled')
        self.details.pack(fill='both', expand=True)
        ttk.Label(panel, textvariable=self.status, wraplength=980).pack(anchor='w', pady=(10, 0))
        ttk.Label(panel, text='Выход: DEV\\Stage5_Выход. Остатки УЗК и статус паспорта не изменяются.').pack(anchor='w')
        verify_pins()
        self.rows = passports()
        projects = sorted({r['project'] for r in self.rows})
        self.project_box['values'] = projects
        self.project_box.bind('<<ComboboxSelected>>', self.project_changed)
        for variable in (self.project, self.passport, self.vom):
            variable.trace_add('write', self.invalidate)
        if 'Гараж Казахстан' in projects:
            self.project.set('Гараж Казахстан')
            self.project_changed()
        root.after(100, self.poll)

    def display(self, value):
        self.details.configure(state='normal')
        self.details.delete('1.0', 'end')
        self.details.insert('1.0', value)
        self.details.configure(state='disabled')

    def invalidate(self, *_):
        self.plan = self.output = None
        self.checked.set(False)
        self.update_buttons()

    def update_buttons(self):
        self.generate_button.configure(state='normal' if self.plan and self.plan.data['marks'] and self.checked.get() and not self.busy else 'disabled')
        self.open_button.configure(state='normal' if self.output and not self.busy else 'disabled')

    def project_changed(self, *_):
        numbers = sorted({r['number'] for r in self.rows if r['project'] == self.project.get()},
                         key=lambda n: tuple(map(int, reversed(n.split('/')))))
        self.passport_box['values'] = numbers
        self.passport.set(numbers[-1] if numbers else '')
        self.vom.set('')

    def select_vom(self):
        value = filedialog.askopenfilename(title='ВОМ выбранной отгрузки — только видимые строки',
                                          initialdir=str(ROOT.parent / 'Гараж казахстан/Паспорт'),
                                          filetypes=[('Excel ВОМ', '*.xls *.xlsx')])
        if value:
            self.vom.set(value)

    def work(self, operation, function):
        if self.busy:
            return
        self.busy = True
        self.preview_button.configure(state='disabled')
        self.select_button.configure(state='disabled')
        self.project_box.configure(state='disabled')
        self.passport_box.configure(state='disabled')
        self.update_buttons()
        self.status.set('Чтение и проверка…' if operation == 'preview' else 'Создание тестового документа…')
        def job():
            try:
                self.queue.put((operation, function(), None))
            except Exception as exc:
                self.queue.put((operation, None, str(exc)))
        Thread(target=job, daemon=True).start()

    def do_preview(self):
        values = (self.project.get(), self.passport.get(), self.vom.get())
        self.invalidate()
        self.work('preview', lambda: preview(*values))

    def do_generate(self):
        if self.plan and self.checked.get():
            plan = self.plan
            self.work('generate', lambda: generate(plan))

    def poll(self):
        try:
            operation, value, error = self.queue.get_nowait()
        except Empty:
            pass
        else:
            self.busy = False
            self.preview_button.configure(state='normal')
            self.select_button.configure(state='normal')
            self.project_box.configure(state='readonly')
            self.passport_box.configure(state='readonly')
            if error:
                self.invalidate()
                self.display('STOP\n' + error)
                self.status.set('Генерация остановлена. Требуется устранить указанную причину.')
            elif operation == 'preview':
                self.plan = value
                p = value.data
                lines = [f'Паспорт {p["passport"]["number"]}; УЗК дата {p["passport"]["date"]}',
                         f'Видимые строки ВОМ: {p["vom"]["visible_count"]}; масса: {p["vom"]["mass_kg"]} кг',
                         f'Марок УЗК: {len(p["marks"])}; единиц: {p["controlled_units"]}', '']
                for m in p['marks']:
                    segments = '; '.join(f'{fmt(s["L"])} × {fmt(s["S"])} мм, {fmt(s["frequency_mhz"])} МГц / {fmt(s["angle_degrees"])}°' for s in m['segments'])
                    lines.extend([f'{m["mark"]} | {m["raw_profile"]} | {m["lookup"]["status"]}',
                                  f'  ВОМ {m["vom_quantity"]}, остаток {m["remaining"]}, Q={m["Q"]}: {segments}', ''])
                self.display('\n'.join(lines))
                self.status.set('Проверьте марки и количества.' if p['marks'] else 'Для текущего паспорта УЗК не требуется')
            else:
                self.output = Path(value) if value else None
                self.status.set(str(value) if value else 'Для текущего паспорта УЗК не требуется')
            self.update_buttons()
        self.root.after(100, self.poll)

    def open_word(self):
        if self.output and self.output.is_file():
            os.startfile(str(self.output))


def run():
    root = tk.Tk()
    try:
        App(root)
    except Exception as exc:
        messagebox.showerror('Stage 5 — STOP', str(exc), parent=root)
        root.destroy()
        return
    root.mainloop()
