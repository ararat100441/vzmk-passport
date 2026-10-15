import argparse
from pathlib import Path
from passport_app.config import Settings
from passport_app.gui import Application


def main():
    parser = argparse.ArgumentParser(description='ВЗМК — паспорта качества. Этап 1')
    parser.add_argument('--config', type=Path, default=Path(__file__).with_name('settings.json'))
    args = parser.parse_args()
    Application(Settings.load(args.config)).mainloop()


if __name__ == '__main__':
    main()
