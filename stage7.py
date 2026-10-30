"""Stage 7 DEV: explicit manifest preview and TEST quality generation."""
import argparse
import sys
sys.dont_write_bytecode = True


def main():
    parser=argparse.ArgumentParser(description='Этап 7 — только DEV/TEST')
    sub=parser.add_subparsers(dest='command')
    assemble=sub.add_parser('assemble',help='New approved UZK + VIK + selected manifest + quality, in validation only')
    assemble.add_argument('--output',required=True)
    preview=sub.add_parser('preview')
    preview.add_argument('manifest')
    generate=sub.add_parser('generate')
    generate.add_argument('manifest')
    generate.add_argument('--output',required=True)
    generate.add_argument('--confirm-test',action='store_true',required=True)
    args=parser.parse_args()
    from stage7_quality.core import assemble as run_assembly, verify_manifest, generate_quality
    from stage7_quality.provenance import read
    if args.command=='assemble':
        print(run_assembly(args.output))
    elif args.command=='preview':
        m=verify_manifest(read(args.manifest))
        import json
        print(json.dumps({'fields':m['fields'],'records':m['records'],'warning':'ТЕСТ: фактический контроль не подтверждён. АКЗ №0: 3 марки / 3 изделия.'},ensure_ascii=False,indent=2))
    elif args.command=='generate':
        print(generate_quality(args.manifest,args.output))
    else:
        from stage7_quality.gui import launch
        launch()


if __name__=='__main__':
    main()
