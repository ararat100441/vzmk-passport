"""Read-only preflight. Historical issued facts are never used as blank forms."""
import re,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
from passport_app.appendix import inspect_template
from passport_app.excel import load_book
from .storage import Stop

WORD_REQUIRED={
    'vik':{'NUMBER','CONTRACT','OBJECT','CIPHER','VIK_DATE','RESPONSIBLE','EVIDENCE','VIK_TABLE'},
    'uzk':{'NUMBER','OBJECT','CIPHER','UZK_DATE','RESPONSIBLE','EVIDENCE','UZK_TABLE'},
    'quality':{'NUMBER','CONTRACT','CUSTOMER','OBJECT','CIPHER','MASS','STEEL','KM','KMD_COMPANY','RESPONSIBLE','EVIDENCE'},
}
AKZ_REQUIRED={'NUMBER','OBJECT','AKZ_DATE','AIR','SURFACE','HUMIDITY','DEW','THICKNESS','MATERIAL','BASIS','RESPONSIBLE','EVIDENCE'}
TOKEN=re.compile(r'\{\{([A-Z_0-9]+)\}\}')

def validate(profile,project,variants,training=False):
    errors=[]
    for role in ['appendix','vik','uzk','quality',*[r+str(n) for n in variants for r in ('akz','akz_appendix')]]:
        ref=profile['templates'].get(role)
        if not ref:
            errors.append(role+': не выбран собственный файл');continue
        path=Path(ref['path'])
        try:
            if role=='appendix':
                inspect_template(path,project);continue
            if role in WORD_REQUIRED:
                if path.suffix.lower()!='.docx':raise Stop('Нужна форма DOCX.')
                with zipfile.ZipFile(path) as z:
                    root=E.fromstring(z.read('word/document.xml'))
                    text=''.join(root.itertext())
                required=WORD_REQUIRED[role]-({'VIK_TABLE'} if training else set())
            else:
                if path.suffix.lower()!='.xlsx':raise Stop('Нужна форма XLSX.')
                book=load_book(path)
                try:
                    if len(book.worksheets)!=1 or book.active.protection.sheet:raise Stop('Нужен один незащищённый лист формы АКЗ.')
                    sheet=book.active;text=' '.join(str(c.value or '') for row in sheet for c in row)
                    if role.startswith('akz_appendix'):
                        markers=[c for row in sheet for c in row if c.value=='{{AKZ_MARKS}}']
                        if len(markers)!=1 or markers[0].column!=1 or sheet.cell(markers[0].row+1,1).value!='Итого':raise Stop('Нужна одна строка {{AKZ_MARKS}} в A и следующая строка «Итого».')
                finally:book.close()
                required={'NUMBER','OBJECT','AKZ_DATE','AKZ_MARKS'} if role.startswith('akz_appendix') else AKZ_REQUIRED
            missing=required-set(TOKEN.findall(text))
            if missing:raise Stop('Нет полей '+', '.join('{{'+k+'}}' for k in sorted(missing))+'. Оформленный исторический акт не является пустой формой для новых результатов.')
        except Exception as exc:errors.append(role+' — '+str(path)+'\n'+str(exc))
    if errors:raise Stop('Формы объекта не готовы для выпуска:\n\n'+'\n\n'.join(errors))
    return True
