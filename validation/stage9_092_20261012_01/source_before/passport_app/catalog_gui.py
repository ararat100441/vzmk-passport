"""Separate Stage 3 window; accepted Stage 1/2 GUI stays intact."""
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import traceback

from .catalog import CatalogService, LABELS, CATALOG, project_folder
from .config import PassportError


class CatalogApplication(tk.Tk):
    def __init__(self, settings):
        service = CatalogService(settings)
        super().__init__()
        self.title('ВЗМК — Справочник марок. Этап 3')
        self.geometry('1150x820')
        self.minsize(900, 650)
        self.plan = None
        self.confirmations = set()
        self.service = service
        self.project_var = tk.StringVar()
        self.path_var = tk.StringVar(value='Выберите проект из ВЗМК.xlsx')
        self.result_var = tk.StringVar(value='«Проверить импорт» ничего не записывает. № АКЗ сохраняется.')
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(4, weight=2)
        frame.rowconfigure(7, weight=3)
        line = ttk.Frame(frame)
        line.grid(row=0, column=0, sticky='ew')
        line.columnconfigure(1, weight=1)
        ttk.Label(line, text='Проект:').grid(row=0, column=0, padx=(0, 10))
        self.combo = ttk.Combobox(line, textvariable=self.project_var, state='readonly',
                                  values=tuple(self.service.registry.projects))
        self.combo.grid(row=0, column=1, sticky='ew')
        self.combo.bind('<<ComboboxSelected>>', self.choose_project)
        ttk.Button(line, text='Открыть справочник марок', command=self.open_catalog).grid(row=0, column=2, padx=(12, 0))
        ttk.Label(frame, textvariable=self.path_var, wraplength=1060).grid(row=1, column=0, sticky='w', pady=8)
        ttk.Label(frame, text='Импорт Марка → Профиль', font=('Segoe UI', 12, 'bold')).grid(row=2, column=0, sticky='w')
        ttk.Label(frame, text='Вставьте таблицу из NotebookLM (Марка | Профиль) или откройте текстовый файл UTF-8.').grid(row=3, column=0, sticky='w', pady=6)
        input_frame = ttk.Frame(frame)
        input_frame.grid(row=4, column=0, sticky='nsew')
        self.input = tk.Text(input_frame, wrap='none', height=9, undo=True, font=('Segoe UI', 11))
        scroll = ttk.Scrollbar(input_frame, command=self.input.yview)
        self.input.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.input.pack(fill='both', expand=True)
        self.input.bind('<<Modified>>', self.input_changed)
        self.input.bind('<Control-KeyPress>', self.edit_shortcut)
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, sticky='w', pady=10)
        self.open_file_button = ttk.Button(buttons, text='Открыть файл', command=self.open_file)
        self.open_file_button.pack(side='left', padx=(0, 12))
        self.check_button = ttk.Button(buttons, text='Проверить импорт', command=self.check)
        self.check_button.pack(side='left', padx=(0, 12))
        self.apply_button = ttk.Button(buttons, text='Применить импорт', command=self.apply, state='disabled')
        self.apply_button.pack(side='left')
        ttk.Label(frame, text='Предварительная проверка. Для конфликта или ПРОВЕРИТЬ отметьте конкретную строку в столбце «Разрешить».',
                  wraplength=1060).grid(row=6, column=0, sticky='w', pady=(0, 8))
        preview_frame = ttk.Frame(frame)
        preview_frame.grid(row=7, column=0, sticky='nsew')
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(preview_frame, columns=('confirm', 'mark', 'current', 'incoming', 'status', 'note'),
                                 show='headings', selectmode='browse', height=10)
        for name, title, width in zip(self.tree['columns'], ('Разрешить', 'Марка', 'Сейчас', 'NotebookLM', 'Статус', 'Предупреждение'),
                                      (85, 130, 140, 160, 255, 350)):
            self.tree.heading(name, text=title)
            self.tree.column(name, width=width, minwidth=60, stretch=name in ('status', 'note'))
        self.tree.grid(row=0, column=0, sticky='nsew')
        vertical = ttk.Scrollbar(preview_frame, command=self.tree.yview)
        horizontal = ttk.Scrollbar(preview_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.grid(row=0, column=1, sticky='ns')
        horizontal.grid(row=1, column=0, sticky='ew')
        self.tree.bind('<Button-1>', self.click_confirmation)
        self.tree.bind('<space>', self.toggle_selected)
        ttk.Button(frame, text='Подтвердить / снять подтверждение выбранной строки', command=self.toggle_selected).grid(row=8, column=0, sticky='w', pady=8)
        ttk.Label(frame, textvariable=self.result_var, wraplength=1060).grid(row=9, column=0, sticky='w', pady=(6, 0))

    def invalidate(self):
        self.plan = None
        self.confirmations.clear()
        self.apply_button.configure(state='disabled')
        self.tree.delete(*self.tree.get_children())

    def input_changed(self, event=None):
        if self.input.edit_modified():
            self.invalidate()
            self.result_var.set('Входная таблица изменена. Нажмите «Проверить импорт».')
            self.input.edit_modified(False)

    def edit_shortcut(self, event):
        # Windows reports Cyrillic keysyms as ?? in this Tk version. Use virtual
        # key codes for the four standard editing shortcuts; delegate to Tk.
        if self.tk.call('tk', 'windowingsystem') == 'win32':
            action = {65: '<<SelectAll>>', 67: '<<Copy>>', 86: '<<Paste>>', 88: '<<Cut>>'}.get(event.keycode)
        else:
            action = {'a': '<<SelectAll>>', 'c': '<<Copy>>', 'v': '<<Paste>>', 'x': '<<Cut>>'}.get(event.keysym.lower())
        if action:
            if action == '<<SelectAll>>':
                self.input.tag_add('sel', '1.0', 'end-1c')
            else:
                self.input.event_generate(action)
            return 'break'

    def input_error(self, message):
        self.invalidate()
        self.result_var.set(message)
        messagebox.showerror('Ошибка импорта', message, parent=self)

    def open_file(self):
        filename = filedialog.askopenfilename(parent=self, title='Открыть файл импорта',
            initialdir=str(self.service.settings.registry_path.parent),
            filetypes=(('Текст UTF-8', '*.txt'), ('Markdown', '*.md'), ('Таблица с табуляцией', '*.tsv')))
        if not filename:
            return
        self.invalidate()
        try:
            path = Path(filename)
            if path.suffix.lower() not in ('.txt', '.md', '.tsv'):
                raise PassportError('Выберите текстовый файл TXT, MD или TSV в кодировке UTF-8.')
            text = path.read_bytes().decode('utf-8-sig')
            self.input.delete('1.0', 'end')
            self.input.insert('1.0', text)
            self.input.edit_modified(False)
            self.result_var.set(f'Загружен файл: {path}. Нажмите «Проверить импорт».')
            self.input.focus_set()
        except UnicodeError:
            self.input_error('Не удалось прочитать файл как UTF-8. Выберите текстовый файл в кодировке UTF-8.')
        except (PassportError, OSError) as exc:
            self.input_error(f'Не удалось загрузить текстовый файл: {exc}')

    def choose_project(self, event=None):
        self.invalidate()
        try:
            self.service.assert_registry()
            path = project_folder(self.service.settings, self.project_var.get()) / CATALOG
            self.path_var.set(str(path))
            self.result_var.set('Вставьте таблицу и выполните предварительную проверку.')
        except (PassportError, OSError) as exc:
            self.path_var.set('Папка проекта не определена')
            self.result_var.set(str(exc))

    def open_catalog(self):
        try:
            self.service.assert_registry()
            path = project_folder(self.service.settings, self.project_var.get()) / CATALOG
            if not path.is_file():
                raise PassportError('Справочника ещё нет. Он будет создан после проверки и применения импорта.')
            os.startfile(path)
        except (PassportError, OSError) as exc:
            self.result_var.set(str(exc))

    def check(self):
        self.invalidate()
        # A pending <<Modified>> from the text just read must not erase the
        # preview after this callback returns. Subsequent edits still invalidate.
        text = self.input.get('1.0', 'end-1c')
        self.input.edit_modified(False)
        try:
            self.plan = self.service.preview(self.project_var.get(), text)
            self.path_var.set(str(self.plan.path))
            for i, row in enumerate(self.plan.rows):
                self.tree.insert('', 'end', iid=str(i), values=('☐' if row.confirmable else '—', row.mark,
                    row.current, row.incoming, LABELS[row.status], row.warning))
            error = any(r.status == 'ERROR' for r in self.plan.rows)
            self.apply_button.configure(state='disabled' if error else 'normal')
            self.result_var.set('Исправьте ошибки входа. Ничего не записано.' if error else
                                'Preview готов. Конфликты без отметки будут пропущены. Ничего не записано.')
        except (PassportError, OSError) as exc:
            self.input_error(f'Не удалось проверить импорт: {exc}')
        except Exception:
            traceback.print_exc()
            self.input_error('Не удалось проверить импорт из-за внутренней ошибки. '
                             'Предыдущая проверка сброшена. Ничего не записано.')

    def click_confirmation(self, event):
        if self.tree.identify_column(event.x) == '#1':
            item = self.tree.identify_row(event.y)
            if item:
                self.tree.selection_set(item)
                self.toggle_selected()

    def toggle_selected(self, event=None):
        selection = self.tree.selection()
        if not self.plan or not selection:
            return
        item = selection[0]
        row = self.plan.rows[int(item)]
        if not row.confirmable:
            return
        if row.token in self.confirmations:
            self.confirmations.remove(row.token)
        else:
            self.confirmations.add(row.token)
        self.tree.set(item, 'confirm', '☑' if row.token in self.confirmations else '☐')

    def apply(self):
        if not self.plan:
            return
        plan = self.plan
        if (self.project_var.get() != plan.project or self.input.get('1.0', 'end-1c') != plan.text):
            self.invalidate()
            self.result_var.set('Проект или вход изменился. Повторите проверку.')
            return
        if not messagebox.askyesno('Применить импорт',
                f'Проект: {plan.project}\nПрименить новые марки, пустые профили и отмеченные строки?\n№ АКЗ сохраняется.', parent=self):
            return
        try:
            result = self.service.apply(plan, self.confirmations)
            self.invalidate()
            self.result_var.set(f'Добавлено: {result["added"]}; заполнено: {result["filled"]}; изменено: {result["changed"]}; '
                                f'без изменений: {result["unchanged"]}; требует проверки: {result["check_required"]}; '
                                f'пропущено: {result["skipped"]}.' +
                                (f' Резервная копия: {result["backup"]}' if result['backup'] else ' Запись не потребовалась.'))
        except (PassportError, OSError) as exc:
            self.invalidate()
            self.result_var.set(str(exc))
