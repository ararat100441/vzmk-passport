"""Separate DEV entry point. Stages 1–4 remain unchanged."""
import argparse
import json
import sys

from uzk_stage5.core import generate, preview
from uzk_stage5.inputs import Stage5Error, verify_pins


def main():
    parser = argparse.ArgumentParser(description='Stage 5 — TEST/DRAFT УЗК, DEV only')
    parser.add_argument('command', choices=['gui', 'check', 'preview', 'generate'])
    parser.add_argument('--project', default='Гараж Казахстан')
    parser.add_argument('--passport', default='9/2026')
    parser.add_argument('--vom')
    parser.add_argument('--output-dir')
    args = parser.parse_args()
    try:
        if args.command == 'gui':
            from uzk_stage5.gui import run
            run()
        elif args.command == 'check':
            import tkinter
            import win32com.client
            import lxml.etree
            verify_pins()
            print('STAGE5_DEV_CHECK_OK; registry writes disabled')
        else:
            if not args.vom:
                parser.error('--vom is required')
            plan = preview(args.project, args.passport, args.vom)
            result = {'preview': plan.data, 'preview_seal': plan.seal}
            if args.command == 'generate':
                result['output'] = str(generate(plan, args.output_dir)) if plan.data['marks'] else None
            print(json.dumps(result, ensure_ascii=False, indent=2))
    except (Stage5Error, OSError, ValueError) as exc:
        print(f'STOP: {exc}', file=sys.stderr)
        return 2
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
