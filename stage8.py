"""Stage 8 standalone DEV launcher. Entirely local; no API or AI dependency."""
import argparse
import json
from pathlib import Path
import sys
sys.dont_write_bytecode=True


def main():
    from stage8_closure import core
    from stage8_closure.storage import read, Stop
    parser=argparse.ArgumentParser(description='Этап 8 — безопасное закрытие, только изолированный TEST')
    sub=parser.add_subparsers(dest='command')
    p=sub.add_parser('fixture');p.add_argument('--parent',required=True);p.add_argument('--no-uzk',action='store_true')
    p=sub.add_parser('preview');p.add_argument('request');p.add_argument('--save-preview')
    p=sub.add_parser('close');p.add_argument('preview');p.add_argument('--confirm',required=True)
    p=sub.add_parser('recover');p.add_argument('sandbox')
    p=sub.add_parser('restore-preview');p.add_argument('sandbox');p.add_argument('--save-preview',required=True)
    p=sub.add_parser('restore');p.add_argument('preview');p.add_argument('--confirm',required=True)
    p=sub.add_parser('inspect-stage7');p.add_argument('manifest')
    args=parser.parse_args()
    try:
        if args.command=='fixture':
            from stage8_closure.fixtures import create
            result=create(args.parent,uzk=not args.no_uzk)
        elif args.command=='preview':
            request=Path(args.request).resolve();result=core.preview(request.parents[1],request)
            if args.save_preview:
                from stage8_closure.storage import within,write
                target=within(args.save_preview,request.parents[1])
                write(target,result)
            result={'preview':result,'confirmation':core.confirmation(result),'summary':core.describe(result)}
        elif args.command=='close':result=core.close(read(args.preview),args.confirm)
        elif args.command=='recover':result=core.recover(args.sandbox)
        elif args.command=='restore-preview':
            from stage8_closure import recovery
            from stage8_closure.storage import within,write
            result=recovery.preview(args.sandbox)
            write(within(args.save_preview,Path(result['payload']['sandbox'])),result)
            result={'preview':result,'confirmation':recovery.confirmation(result),'summary':recovery.describe(result)}
        elif args.command=='restore':
            from stage8_closure import recovery
            result=recovery.restore(read(args.preview),args.confirm)
        elif args.command=='inspect-stage7':result=core.inspect_stage7(args.manifest)
        else:
            from stage8_closure.gui import launch
            launch();return
        print(json.dumps(result,ensure_ascii=False,indent=2))
    except Exception as exc:
        print('STOP: '+str(exc),file=sys.stderr);sys.exit(2)


if __name__=='__main__':main()
