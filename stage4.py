"""Standalone Stage 4 CLI and optional viewer; fixed DEV target, offline only."""
import argparse
import json
import sys
from profiles_stage4.model import calculate_uzk
from profiles_stage4.storage import (build, lookup, load_catalog, add_manual, set_preference,
                                     resolve_conflict, confirm_manual_uzk)


def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description='Этап 4 — локальный DEV справочник профилей')
    sub = parser.add_subparsers(dest='command', required=True)
    sub.add_parser('build', help='Offline rebuild; preserves user state')
    p = sub.add_parser('lookup'); p.add_argument('profile'); p.add_argument('--standard')
    sub.add_parser('counts')
    sub.add_parser('gui')
    p = sub.add_parser('manual-add'); p.add_argument('json_file', help='UTF-8 JSON: profile,family,standard,geometry,note')
    p = sub.add_parser('preference'); p.add_argument('profile'); p.add_argument('standard'); p.add_argument('--basis', required=True)
    p = sub.add_parser('resolve'); p.add_argument('conflict_id'); p.add_argument('record_id'); p.add_argument('--reason', required=True)
    p = sub.add_parser('confirm-manual-uzk'); p.add_argument('record_id'); p.add_argument('--basis', required=True)
    args = parser.parse_args()
    try:
        if args.command == 'gui':
            from profiles_stage4.gui import ProfileWindow
            ProfileWindow().mainloop()
            return 0
        if args.command == 'build':
            result = build()
        elif args.command == 'lookup':
            result = lookup(args.profile, args.standard)
            if result['status'] in ('FOUND', 'FOUND_PREFERRED'):
                result['uzk'] = calculate_uzk(result)
        elif args.command == 'counts':
            result = load_catalog().counts()
        elif args.command == 'manual-add':
            with open(args.json_file, encoding='utf-8-sig') as f:
                data = json.load(f)
            result = add_manual(data['profile'], data['family'], data['standard'], data['geometry'], data.get('note', ''))
        elif args.command == 'preference':
            result = set_preference(args.profile, args.standard, args.basis)
        elif args.command == 'resolve':
            result = resolve_conflict(args.conflict_id, args.record_id, args.reason)
        else:
            result = confirm_manual_uzk(args.record_id, args.basis)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (ValueError, OSError, KeyError) as exc:
        print('STOP: ' + str(exc), file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
