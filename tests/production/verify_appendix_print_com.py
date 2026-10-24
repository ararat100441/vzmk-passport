r"""Windows Excel Gate B: print-title pagination on disposable copies.

Usage:
  py -3.12 -B tests/production/verify_appendix_print_com.py ^
      --appendix-one "C:\...\1_1_2026 - Приложение № 1.xlsx" ^
      --akz-appendix "C:\...\approved-blank-AKZ-appendix.xlsx"

The function intentionally inserts SYNTHETIC table rows only in temporary
copies to force two or more pages. No real registry or original is edited.
Review the exported PDFs visually before accepting any production template.
"""
import argparse
import hashlib
import json
import shutil
import sys
import tempfile
from pathlib import Path

from pypdf import PdfReader
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT))
from passport_app.excel import excel_session
from production_app.office import set_appendix_print_titles


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def export_copy(app,source,folder,label):
    original=sha(source)
    scratch=folder/(label+'.xlsx')
    shutil.copyfile(source,scratch)
    result=folder/(label+'.pdf')
    book=app.Workbooks.Open(Filename=str(scratch),ReadOnly=False,
                            UpdateLinks=0,AddToMru=False,
                            IgnoreReadOnlyRecommended=True)
    try:
        if book.ReadOnly or book.Worksheets.Count!=1:
            raise RuntimeError('Неожиданная защищённая Excel книга: '+str(scratch))
        s=book.Worksheets(1)
        first=5
        if 'отправочная марка' not in str(s.Cells(3,1).Value2 or '').casefold():
            raise ValueError('Требуется двухуровневая шапка с маркировкой в A3.')
        # Leave the source title, long object and two-level headers in rows1:4.
        # Extend only the COPY with 110 explicitly synthetic data rows.
        data=[('SYNTHETIC-'+str(i).zfill(3),'ТЕСТОВЫЙ ЭЛЕМЕНТ',1,
               float(i),float(i),float(i),float(i))
              for i in range(1,111)]
        s.Range('A5:G114').Value2=tuple(data)
        s.Range('A5:G114').RowHeight=18
        s.Range('A5:G114').Borders.LineStyle=1
        s.PageSetup.PrintArea='$A$1:$G$114'
        set_appendix_print_titles(s,first)
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    # Reopen and compare the native persisted print titles before PDF export.
    proof=app.Workbooks.Open(Filename=str(scratch),ReadOnly=True,
                             UpdateLinks=0,AddToMru=False)
    try:
        s=proof.Worksheets(1)
        actual=str(s.PageSetup.PrintTitleRows or '').replace('$','')
        if '3:4' not in actual:
            raise RuntimeError('После Save/Open Excel не повторяет компактную шапку.')
        proof.ExportAsFixedFormat(
            Type=0,Filename=str(result),Quality=0,
            IncludeDocProperties=True,IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
    finally:
        proof.Close(SaveChanges=False)
    pdf=PdfReader(str(result),strict=False)
    pages=len(pdf.pages)
    if pages<2:
        raise RuntimeError('Для проверки второй страницы нужен PDF минимум из 2 страниц.')
    if sha(source)!=original:
        raise RuntimeError('Оригинальный Excel изменён: '+str(source))
    return {'source':str(source),'original_sha256':original,'pages':pages,
            'print_title_rows':'$3:$4','print_pdf':str(result)}


def main(first,akz,output):
    first=Path(first).resolve(strict=True)
    akz=Path(akz).resolve(strict=True)
    if first==akz or any(p.suffix.lower()!='.xlsx' for p in (first,akz)):
        raise ValueError('Нужны два разных подтверждённых XLSX источника.')
    destination=Path(output).resolve() if output else None
    if destination is not None:
        destination.mkdir(parents=True,exist_ok=False)
    with tempfile.TemporaryDirectory(prefix='vzmk-print-com-') as temporary:
        folder=destination if destination is not None else Path(temporary)
        with excel_session() as app:
            results=[
                export_copy(app,first,folder,'synthetic-appendix1'),
                export_copy(app,akz,folder,'synthetic-akz-appendix'),
            ]
        # Without --output, temporary diagnostic PDFs are deleted. With
        # --output they remain ONLY in a newly created diagnostic folder.
        return {'status':'PASS','print_title_rows':'$3:$4',
                'originals_unchanged':True,'reports':results,
                'diagnostics_retained':destination is not None}


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--appendix-one',required=True)
    parser.add_argument('--akz-appendix',required=True)
    parser.add_argument('--output',help='NEW diagnostic directory for visual PDF inspection')
    args=parser.parse_args()
    try:
        print(json.dumps(main(args.appendix_one,args.akz_appendix,args.output),
                         ensure_ascii=False,indent=2))
    except BaseException as exc:
        print('FAIL '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        raise
