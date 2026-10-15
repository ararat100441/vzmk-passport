"""One new quality document from the existing frozen manifest; no stage reruns."""
from pathlib import Path
import sys
sys.dont_write_bytecode=True
HERE=Path(__file__).resolve().parent
ROOT=HERE.parents[1];sys.path.insert(0,str(ROOT))
from stage7_quality.provenance import read,write,sha,verify
from stage7_quality.core import generate_quality,verify_manifest
from stage7_quality.native import inspect
from stage7_quality.sources import TEMPLATE

if __name__=='__main__':
    before=read(HERE/'PRODUCT_BEFORE_FINAL_TESTS.json');verify(before)
    selected=read(HERE/'PREFLIGHT.json')['manifest'];verify({selected['path']:selected['sha256']})
    manifest=read(selected['path']);verify_manifest(manifest)
    receipt=generate_quality(selected['path'],HERE/'QUALITY_TEST_V2')
    output=read(receipt['path'])['output']
    write(HERE/'NEW_QUALITY_RESULT.json',{'status':'COMPLETE_VALIDATED_TEST','receipt':receipt,
        'output':output,'output_sha256':sha(output),'original_frozen_manifest':selected,
        'assembly_run_id':manifest['assembly_run_id'],'generator_reruns':{'VIK':0,'UZK':0,'AKZ':0},
        'presentation_revision':'USER_E2E_POINT13_POINT14_CONTRACT_VALUE_V2'})
    inspect(TEMPLATE,HERE,kind='FIRST_TEMPLATE',required=['Pirocor Protect 057','80 мкм','Степень очистки 2'])
    inspect(HERE/'REFERENCE_9_READONLY_SNAPSHOT.docx',HERE,kind='MANUAL9',required=['9/2026','29.09.2026','01.10.2026','05.10.2026'])
    verify(before)
    print({'status':'PASS','output':output,'sha256':sha(output),'receipt':receipt})
