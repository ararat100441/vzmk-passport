"""Small read-only catalogue viewer; user-state edits use the explicit CLI/API."""
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox
from .model import calculate_uzk
from .storage import DEFAULT_WORKBOOK, lookup


class ProfileWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Справочник профилей — Этап 4 DEV')
        self.geometry('1000x690')
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Профиль').grid(row=0, column=0, sticky='w')
        self.profile = ttk.Entry(frame, width=30)
        self.profile.grid(row=0, column=1, sticky='ew', padx=8)
        self.profile.insert(0, '35Б1')
        ttk.Label(frame, text='НД (необязательно)').grid(row=0, column=2)
        self.standard = ttk.Entry(frame, width=28)
        self.standard.grid(row=0, column=3, sticky='ew', padx=8)
        ttk.Button(frame, text='Найти', command=self.search).grid(row=0, column=4)
        ttk.Button(frame, text='Открыть DEV Профили.xlsx', command=self.open_workbook).grid(row=1, column=0, columnspan=3, sticky='w', pady=12)
        self.status = ttk.Label(frame, text='Геометрия в мм. УЗК — для одного изделия.')
        self.status.grid(row=2, column=0, columnspan=5, sticky='w')
        self.output = tk.Text(frame, wrap='word', font=('Consolas', 11))
        self.output.grid(row=3, column=0, columnspan=5, sticky='nsew', pady=8)
        scroll = ttk.Scrollbar(frame, orient='vertical', command=self.output.yview)
        scroll.grid(row=3, column=5, sticky='ns')
        self.output.configure(yscrollcommand=scroll.set, state='disabled')
        frame.rowconfigure(3, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)
        self.bind('<Return>', lambda e: self.search())
        self.profile.focus_set()

    def show(self, text):
        self.output.configure(state='normal')
        self.output.delete('1.0', 'end')
        self.output.insert('1.0', text)
        self.output.configure(state='disabled')

    def search(self):
        try:
            result = lookup(self.profile.get(), self.standard.get().strip() or None)
            self.status.configure(text=result['status'])
            lines = [f"Статус: {result['status']}", f"Ключ: {result['key']}"]
            chosen = result.get('profile')
            if chosen and result['status'] in ('FOUND', 'FOUND_PREFERRED'):
                lines += [f"Тип: {chosen['family']}", f"НД: {chosen['standard']}", f"Источник/статус: {chosen['status']}",
                          'Размеры, мм: ' + json.dumps(chosen['geometry'], ensure_ascii=False)]
                uzk = calculate_uzk(result)
                lines.append('УЗК: ' + uzk['status'])
                for r in uzk['rows']:
                    lines.append(f"{r['L']} × {r['S']} мм; {r['frequency_mhz']} МГц / {r['angle_degrees']}°")
            else:
                for row in result['candidates']:
                    lines.append(f"{row['standard']}: {row['status']}; {row['geometry']}; {row['conflict_id']}")
            lines.append('\nProvenance:')
            rows = [chosen] if chosen else result['candidates']
            for row in rows:
                for p in row['provenance']:
                    lines.append(json.dumps({'source': p['source'], 'record_id': p['record_id'], 'НД': p['standard'],
                                             'date': p['date'], 'evidence': p['evidence']}, ensure_ascii=False))
                lines.extend(row['notes'])
            self.show('\n'.join(lines))
        except (ValueError, OSError, KeyError) as exc:
            self.status.configure(text='ОШИБКА: каталог требует проверки')
            self.show(str(exc))

    def open_workbook(self):
        try:
            # Hash preflight prevents opening a stale/unbound generation through this launcher.
            from .storage import load_catalog
            load_catalog()
            os.startfile(str(DEFAULT_WORKBOOK))
        except (ValueError, OSError, KeyError) as exc:
            messagebox.showerror('DEV справочник', str(exc), parent=self)
