"""Read-only preflight. Historical issued facts are never used as blank forms."""
import re
from pathlib import Path
from passport_app.appendix import inspect_template
from passport_app.excel import load_book
from .storage import Stop
from .word_controls import inspect_word_fields

WORD_REQUIRED={
    'vik':{'NUMBER','CONTRACT','OBJECT','CIPHER','VIK_DATE','RESPONSIBLE','EVIDENCE'},
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
                tokens, controls = inspect_word_fields(path)
                # Historical approved VIK layouts refer to the appendix and protocols
                # rather than including a per-mark table. If an explicit {{VIK_TABLE}}
                # marker exists, the renderer still adds one; otherwise no new table
                # is imposed on the approved two-act form.
                required=WORD_REQUIRED[role]
                provided = tokens | controls
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
                provided = set(TOKEN.findall(text))
                # An approved three-layer act may display ranges instead of
                # the old single THICKNESS field. Never accept a partial set
                # of layer placeholders or substitute one layer for another.
                if role.startswith('akz') and not role.startswith('akz_appendix'):
                    layer_tokens={key for key in provided if key.startswith('LAYER_')}
                    if layer_tokens:
                        expected={f'LAYER_{n}_RANGE' for n in (1,2,3)}
                        numeric={f'LAYER_{n}_{edge}' for n in (1,2,3) for edge in ('MIN','MAX')}
                        means={f'LAYER_{n}_DRY_AVG' for n in (1,2,3)}
                        totals={f'LAYER_{n}_{edge}' for n in (1,2,3) for edge in ('TOTAL_MIN','TOTAL_MAX')}
                        if not (expected.issubset(provided) or numeric.issubset(provided)):
                            raise Stop('Форма трёхслойного АКЗ должна содержать диапазоны всех трёх слоёв.')
                        if not means.issubset(provided):
                            raise Stop('Форма трёхслойного АКЗ должна содержать среднюю сухую толщину после всех трёх слоёв.')
                        if layer_tokens-(expected|numeric|means|totals):
                            raise Stop('Неподдержанное поле трёхслойного АКЗ.')
                        required = (AKZ_REQUIRED-{'THICKNESS'})|means
            missing=required-provided
            if missing:raise Stop('Нет полей '+', '.join('{{'+k+'}}' for k in sorted(missing))+'. Оформленный исторический акт не является пустой формой для новых результатов.')
        except Exception as exc:errors.append(role+' — '+str(path)+'\n'+str(exc))
    if errors:raise Stop('Формы объекта не готовы для выпуска:\n\n'+'\n\n'.join(errors))
    return True
