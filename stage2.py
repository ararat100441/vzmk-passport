"""Minimal Stage 2 invocation; production generation is unavailable."""
import argparse
import json
import sys
from passport_app.config import Settings, PassportError
from passport_app.vik import generate_vik


def main():
    parser = argparse.ArgumentParser(description='Этап 2 — ВИК только на Stage 2 validation-копии')
    parser.add_argument('--settings', required=True)
    parser.add_argument('--test-root', required=True)
    parser.add_argument('--project', required=True)
    parser.add_argument('--number', required=True)
    fallback = parser.add_mutually_exclusive_group()
    fallback.add_argument('--fallback-template', help='Точный общий ВИК.docx; по умолчанию Шаблоны/ВИК.docx')
    fallback.add_argument('--fallback-root', help='Общая папка; используется только файл ВИК.docx в её корне')
    args = parser.parse_args()
    try:
        result = generate_vik(Settings.load(args.settings), args.project, args.number,
                              test_root=args.test_root, fallback_root=args.fallback_root,
                              fallback_template=args.fallback_template)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except PassportError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
