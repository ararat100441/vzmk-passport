"""Read exactly the ten bounded source acts; preserve literal cell text."""
from collections import Counter
import csv
import json
from pathlib import Path
import sys
from zipfile import ZipFile
from lxml import etree as E

ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from uzk_stage5.inputs import sha
RUN=Path(__file__).resolve().parent
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def cell_text(cell):
    lines=[]
    for p in cell.findall('w:p',NS):
        parts=[]
        for item in p.iter():
            if item.tag=='{'+NS['w']+'}t':parts.append(item.text or '')
            elif item.tag in ('{'+NS['w']+'}br','{'+NS['w']+'}cr'):parts.append('\n')
            elif item.tag=='{'+NS['w']+'}tab':parts.append('\t')
        lines.append(''.join(parts))
    return '\n'.join(lines)


baseline=json.loads((RUN/'baseline.json').read_text(encoding='utf8'))
assert len(baseline['historical_paths'])==10
records=[]; sources=[]
project=ROOT.parent/'Гараж казахстан'
for source in sorted(map(Path,baseline['historical_paths']),key=lambda p:int(p.name.split('_')[0])):
    expected=baseline['protected_before'][str(source)]
    assert sha(source)==expected
    with ZipFile(source) as z:
        doc=E.fromstring(z.read('word/document.xml'))
    matches=[]
    for ti,table in enumerate(doc.findall('w:body/w:tbl',NS),1):
        rows=table.findall('w:tr',NS)
        for hi,row in enumerate(rows):
            cells=row.findall('w:tc',NS)
            if len(cells)==8:
                values=[cell_text(c) for c in cells]
                if ('Обнаруженные дефекты' in values[6].replace('\n',' ')
                        and 'Заключение' in values[7]):
                    matches.append((ti,hi,rows))
    assert len(matches)==1,(source,matches)
    ti,hi,rows=matches[0]
    count=0
    for ri,row in enumerate(rows[hi+1:],hi+2):
        cells=row.findall('w:tc',NS);assert len(cells)==8
        values=[cell_text(c) for c in cells]
        if any(v.strip() for v in values):
            records.append({'act':source.name,'source':str(source),'table':ti,'row':ri,
                            'mark':values[0],'graph7':values[6],'graph8':values[7]})
            count+=1
    assert sha(source)==expected
    sources.append({'relative_path':str(source.relative_to(project)).replace('\\','/'),
                    'sha256':expected,'rows':count})
pairs=Counter((r['graph7'],r['graph8']) for r in records)
assessment={'source_count':len(sources),'row_count':len(records),
    'raw_pairs':[{'graph7':a,'graph8':b,'count':c} for (a,b),c in pairs.items()],
    'empty_graph7':sum(not r['graph7'].strip() for r in records),
    'empty_graph8':sum(not r['graph8'].strip() for r in records),
    'incompatible_variants':len(pairs)>1,'negative_conclusions':[],
    'meaning':'Historical inspection outcomes, not confirmed outcomes for a new shipment.',
    'mode':'SAMPLE_ONLY; official mode disabled; per-mark executor confirmation required in future.'}
(RUN/'historical_rows.json').write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding='utf8')
(RUN/'historical_sources.json').write_text(json.dumps(sources,ensure_ascii=False,indent=2),encoding='utf8')
(RUN/'pattern_assessment.json').write_text(json.dumps(assessment,ensure_ascii=False,indent=2),encoding='utf8')
with (RUN/'HISTORICAL_TEXT.csv').open('w',encoding='utf-8-sig',newline='') as f:
    w=csv.writer(f);w.writerow(['АКТ','МАРКА','ГРАФА 7','ГРАФА 8'])
    for r in records:w.writerow([r['act'],r['mark'],r['graph7'],r['graph8']])
lines=['# Точные формулировки десяти прежних актов УЗК','',
       '| АКТ | МАРКА | ГРАФА 7 | ГРАФА 8 |','|---|---|---|---|']
for r in records:
    values=[r[k].replace('|','\\|').replace('\n','<br>') for k in ('act','mark','graph7','graph8')]
    lines.append('| '+' | '.join(values)+' |')
(RUN/'HISTORICAL_TEXT.md').write_text('\n'.join(lines)+'\n',encoding='utf8')
print(json.dumps({'sources':sources,'assessment':assessment},ensure_ascii=False,indent=2))
