"""Stage 3 launcher. Reading the registry never modifies it."""
import argparse
from pathlib import Path
import tkinter as tk
from tkinter import messagebox
from passport_app.config import Settings, PassportError
from passport_app.catalog_gui import CatalogApplication


def main():
    parser = argparse.ArgumentParser(description='Этап 3 — справочник марок проекта')
    parser.add_argument('--config', type=Path, default=Path(__file__).with_name('settings.json'))
    args = parser.parse_args()
    try:
        CatalogApplication(Settings.load(args.config)).mainloop()
    except (PassportError, OSError) as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror('Этап 3: запуск остановлен', str(exc), parent=root)
        root.destroy()
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
