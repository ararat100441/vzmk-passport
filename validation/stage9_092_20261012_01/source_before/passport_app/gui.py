import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .config import PassportError
from .registry import read_registry, select_project
from .service import validate, generate
from .vom import read_vom


class Application(tk.Tk):
    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self.title('ВЗМК — паспорта качества. Этап 1')
        self.geometry('850x640')
        self.minsize(760, 560)
        self.plan = None
        self.vom_path = ''
        self.events = queue.Queue()
        self.busy = False
        self.project_var = tk.StringVar()
        self.passports_var = tk.StringVar(value='Выберите проект')
        self.file_var = tk.StringVar(value='Отгрузочный ВОМ не выбран')
        self.vom_var = tk.StringVar()
        frame = ttk.Frame(self, padding=18)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text='Проект:').grid(row=0, column=0, sticky='w', padx=(0, 12))
        self.combo = ttk.Combobox(frame, textvariable=self.project_var, state='readonly')
        self.combo.grid(row=0, column=1, sticky='ew')
        self.combo.bind('<<ComboboxSelected>>', self.choose_project)
        ttk.Label(frame, textvariable=self.passports_var, wraplength=790).grid(row=1, column=0, columnspan=2, sticky='w', pady=(12, 18))
        self.choose_button = ttk.Button(frame, text='Выбрать отгрузочный ВОМ', command=self.choose_vom)
        self.choose_button.grid(row=2, column=0, columnspan=2, sticky='w')
        ttk.Label(frame, textvariable=self.file_var, wraplength=790).grid(row=3, column=0, columnspan=2, sticky='w', pady=8)
        ttk.Label(frame, textvariable=self.vom_var, wraplength=790).grid(row=4, column=0, columnspan=2, sticky='w', pady=8)
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=2, sticky='w', pady=12)
        self.check_button = ttk.Button(buttons, text='Проверить', command=self.check)
        self.check_button.pack(side='left', padx=(0, 12))
        self.generate_button = ttk.Button(buttons, text='Сформировать блок 1', command=self.form, state='disabled')
        self.generate_button.pack(side='left')
        self.output = tk.Text(frame, wrap='word', height=15, state='disabled', font=('Segoe UI', 10))
        self.output.grid(row=6, column=0, columnspan=2, sticky='nsew')
        frame.rowconfigure(6, weight=1)
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.after(80, self.poll)
        self.run(lambda: read_registry(settings.registry_path), self.loaded)

    def show(self, text):
        self.output.configure(state='normal')
        self.output.delete('1.0', 'end')
        self.output.insert('1.0', text)
        self.output.configure(state='disabled')

    def loaded(self, registry):
        self.combo['values'] = tuple(registry.projects)
        self.show('Выберите проект и отгрузочный ВОМ, затем нажмите «Проверить».')

    def run(self, operation, callback):
        if self.busy:
            return
        self.busy = True
        self.combo.configure(state='disabled')
        for button in (self.choose_button, self.check_button, self.generate_button):
            button.configure(state='disabled')
        def worker():
            try:
                self.events.put((callback, operation(), None))
            except Exception as exc:
                self.events.put((callback, None, str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            callback, result, error = self.events.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            self.combo.configure(state='readonly')
            self.choose_button.configure(state='normal')
            self.check_button.configure(state='normal')
            if error:
                self.plan = None
                self.show('Ошибок: 1\n' + error)
            else:
                callback(result)
            self.generate_button.configure(state='normal' if self.plan else 'disabled')
        self.after(80, self.poll)

    def choose_project(self, event=None):
        self.plan = None
        name = self.project_var.get()
        def completed(project):
            self.passports_var.set(f'Первый: {project.first_number} (строка {project.first_row})   Последний: {project.last_number}   Новый: {project.next_number}')
            self.show('Проект выбран. Перед формированием выполните «Проверить».')
        self.passports_var.set('Проверка паспортов проекта…')
        self.run(lambda: select_project(read_registry(self.settings.registry_path), name), completed)

    def choose_vom(self):
        path = filedialog.askopenfilename(title='Выберите текущий отгрузочный ВОМ',
                                         filetypes=[('Книга Excel', '*.xlsx')], initialdir=str(self.settings.documents_root))
        if not path:
            return
        self.plan = None
        self.vom_path = path
        self.file_var.set(path)
        self.vom_var.set('Чтение ВОМ…')
        def completed(vom):
            self.vom_var.set(f'Строк: {len(vom.rows)}   Изделий: {vom.quantity:.0f}   Масса: {vom.mass:.2f} кг   Площадь: {vom.area:.3f} м²\nСоединения: {vom.connections}')
            self.show('ВОМ прочитан. Выполните «Проверить».')
        self.run(lambda: read_vom(path), completed)

    def check(self):
        self.plan = None
        project, path = self.project_var.get(), self.vom_path
        self.show('Проверка…')
        def completed(plan):
            self.plan = plan
            self.show(plan.summary())
        self.run(lambda: validate(self.settings, project, path), completed)

    def form(self):
        if not self.plan:
            return
        plan = self.plan
        self.plan = None
        self.show('Формирование блока 1 через Microsoft Excel…')
        def completed(result):
            self.show(f'Блок 1 сформирован.\nПаспорт: {result["number"]}\nСтрока реестра: {result["row"]}\nПапка: {result["folder"]}\nПриложение: {result["appendix"]}\nРезервная копия: {result["backup"]}')
        self.run(lambda: generate(self.settings, plan), completed)

    def close(self):
        if self.busy:
            messagebox.showinfo('Операция выполняется', 'Дождитесь окончания операции перед закрытием программы.')
            return
        self.destroy()
