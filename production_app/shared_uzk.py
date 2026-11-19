"""One shared VZMK ultrasonic inspection XLSM for all projects.

Reading/planning/drafts never write. Official accounting creates a separate
validated candidate; the existing transaction coordinator alone publishes it.

The shared workbook contains *remaining quantities*, NOT inspected weld
outcomes. It can populate the planned UZK quantity for current VOM marks,
but cannot assert GОДЕН, dates, actual weld length or personnel.

Contract from the approved Stage 5 project:
 sheet 'УЗК', first-row columns 'Заказ', 'Марка элемента', 'Кол-во (шт)'.
 Blank numeric cell -> 1; exact identical duplicate doesn't increase stock;
 conflicting duplicate blocks. Selected project's rows only.
"""
from decimal import Decimal, InvalidOperation
from collections import Counter
from bisect import bisect_left
from pathlib import Path
from openpyxl import load_workbook
from .storage import Stop, sha, safe, post_commit_path


HEADERS=('Заказ','Марка элемента','Кол-во (шт)')


def _key(raw):
    return ' '.join(str(raw or '').replace('ё','е').casefold().split())


def _qty(raw,row):
    if raw is None or isinstance(raw,str) and not raw.strip():
        return 1
    if isinstance(raw,bool) or isinstance(raw,str) and raw.startswith('='):
        raise Stop('Общий УЗК: неверное количество в строке '+str(row))
    try:
        q=Decimal(str(raw).strip().replace(',','.'))
    except (ValueError,InvalidOperation):
        raise Stop('Общий УЗК: нечисловое количество в строке '+str(row))
    if not q.is_finite() or q<=0 or q!=q.to_integral_value():
        raise Stop('Общий УЗК: количество должно быть положительным целым, '
                   'строка '+str(row))
    return int(q)


def find_shared_uzk(cfg,*,require=True):
    """A single explicitly selected file, or the unambiguous standard sibling.

    Its location does not depend on the currently selected project, and
    nothing is saved or written during discovery.
    """
    explicit=str(cfg.get('uzk_registry') or '').strip()
    if not explicit:
        registry=cfg.get('registry')
        if not registry:
            if require:raise Stop('Выберите реестр ВЗМК перед поиском общего УЗК.')
            return None
        candidate=Path(registry).parent/'УЗК.xlsm'
    else:
        candidate=Path(explicit)
    candidate=safe(candidate)
    if not candidate.is_file():
        if require:
            raise Stop('Общий реестр УЗК.xlsm не найден: '+str(candidate)+
                       '. Выберите его один раз в «Настройки путей».')
        return None
    if candidate.suffix.lower()!='.xlsm':
        raise Stop('Общий реестр УЗК должен быть файлом *.xlsm.')
    return candidate


def read_shared_uzk(path,*,project=None,marks=None,post_commit=False):
    """Read only values; reject formula quantities, duplicate contradictory rows."""
    path=post_commit_path(path) if post_commit else safe(path)
    before=sha(path)
    book=load_workbook(path,read_only=True,data_only=False,keep_links=False)
    rows={}
    project_key=_key(project) if project is not None else None
    selected_marks={_key(mark) for mark in marks} if marks is not None else None
    try:
        if 'УЗК' not in book.sheetnames:
            raise Stop('В общем XLSM отсутствует лист «УЗК».')
        sheet=book['УЗК']
        header=next(sheet.iter_rows(min_row=1,max_row=1,max_col=3,
                                    values_only=True),())
        if tuple(_key(h) for h in header)!=tuple(_key(h) for h in HEADERS):
            raise Stop('Неверные заголовки УЗК: ожидаются Заказ | Марка '
                       'элемента | Кол-во (шт).')
        for no,values in enumerate(
            sheet.iter_rows(min_row=2,max_col=3,values_only=True),start=2
        ):
            order,mark,qty=values
            if project_key is not None:
                if isinstance(order,str) and order.strip() and not order.startswith('='):
                    if _key(order)!=project_key:continue
                elif selected_marks is not None and _key(mark) not in selected_marks:
                    continue  # Unassigned rows of other marks remain opaque and untouched.
            if all(x is None or isinstance(x,str) and not x.strip()
                   for x in values):
                continue
            if (not isinstance(order,str) or not order.strip() or
                    not isinstance(mark,str) or not mark.strip() or
                    order.startswith('=') or mark.startswith('=')):
                raise Stop('Общий УЗК: заказ/марка отсутствует или неверен '
                           'в строке '+str(no))
            count=_qty(qty,no)
            key=(_key(order),_key(mark))
            if key in rows:
                if rows[key]['remaining']!=count:
                    raise Stop('Общий УЗК: противоречивые остатки у марки '+
                               mark+', строки '+str(rows[key]['rows'][0])+
                               ' и '+str(no))
                rows[key]['rows'].append(no)
            else:
                rows[key]={'project':order,'mark':mark,'remaining':count,
                           'rows':[no]}
    finally:
        book.close()
    if sha(path)!=before:
        raise Stop('Общий реестр УЗК изменился во время чтения.')
    return rows


def planned_control(path,project,vom):
    """Compute min(shipped, remaining) independently for every VOM mark.

    The first encounter determines display name; repeated same mark in VOM
    accumulates quantity. Never sum different marks or duplicated registry
    rows. Missing registry mark => no planned UZK, NOT a negative verdict.
    """
    shipped={}
    labels={}
    for row in vom.rows:
        mark=str(row[0])
        key=_key(mark)
        if not key:
            raise Stop('Марка ВОМ не определена.')
        quantity=_qty(row[2], 'ВОМ '+mark)
        shipped[key]=shipped.get(key,0)+quantity
        labels.setdefault(key,mark)
    rows=read_shared_uzk(path,project=project,marks=labels.values())
    output={}
    for key,quantity in shipped.items():
        rec=rows.get((_key(project),key))
        required=min(quantity,rec['remaining']) if rec else 0
        output[labels[key]]={
            'mark':labels[key],
            'vom_quantity':quantity,
            'registry_remaining':rec['remaining'] if rec else None,
            'uzk_quantity':required,
            'matched':bool(rec),
            'registry_rows':list(rec['rows']) if rec else [],
        }
    return output


def verify_planned_quantities(facts,plan):
    """Cannot change a computed quantity to skip selected current inspection."""
    expected={name:int(x['uzk_quantity']) for name,x in plan.items()}
    marks=[str(r.get('mark','')) for r in facts.get('marks',[])]
    missing=[m for m in expected if m not in marks]
    extra=list(dict.fromkeys(m for m in marks if m not in expected))
    frequencies=Counter(marks)
    duplicates=list(dict.fromkeys(m for m in marks if frequencies[m]>1))
    if missing or extra or duplicates:
        raise Stop('УЗК: состав марок не совпадает с текущей ВОМ; ожидалось '+str(len(expected))+
                   ', найдено '+str(len(marks))+'.\nОтсутствуют: '+(', '.join(missing) or 'нет')+
                   '.\nЛишние: '+(', '.join(extra) or 'нет')+'.\nПовторы: '+(', '.join(duplicates) or 'нет')+'.')
    for row in facts.get('marks',[]):
        mark=str(row.get('mark',''))
        if mark not in expected:
            raise Stop('УЗК: неизвестная марка текущей ВОМ '+mark)
        raw=row.get('uzk_quantity')
        if raw is None or isinstance(raw,bool) or isinstance(raw,str) and not raw.strip():
            raise Stop('УЗК: не заполнено действующее количество для марки '+mark)
        observed=_qty(raw,'ВОМ '+mark) if raw not in (0,'0') else 0
        if observed!=expected[mark]:
            raise Stop('УЗК: количество по марке '+mark+' должно быть '+
                       str(expected[mark])+' согласно общему реестру. '
                       'Изменять его вручную для выпуска нельзя.')
    return True

def decrement_candidate(source,target,project,quantities):
    """Patch ONLY the data sheet XML, preserving VBA and all other ZIP parts.

    Do not round-trip an XLSM through openpyxl (which can strip extensions).
    Complex references needing cross-sheet edits are refused, never guessed.
    """
    import re,posixpath
    from zipfile import ZipFile
    from lxml import etree as E
    source=safe(source);target=safe(target)
    if source==target or target.exists():raise Stop('УЗК: нужен отдельный новый кандидат XLSM.')
    original=sha(source);rows=read_shared_uzk(source,project=project,marks=quantities);updates={};deletions=[]
    for mark,value in quantities.items():
        if value in (0,'0'):continue
        count=_qty(value,mark);rec=rows.get((_key(project),_key(mark)))
        if not rec:raise Stop('УЗК: отсутствует списываемая марка '+mark)
        if len(rec['rows'])!=1:raise Stop('УЗК: дубли строк марки '+mark+'; запись запрещена.')
        if count>rec['remaining']:raise Stop('УЗК: недостаточный остаток '+mark)
        row=rec['rows'][0];left=rec['remaining']-count
        if left:updates[row]=left
        else:deletions.append(row)
    if not updates and not deletions:
        # A zero plan changes neither quantities nor any opaque XLSM part.
        import shutil
        shutil.copyfile(source,target)
        if sha(source)!=original or sha(target)!=original:raise Stop('УЗК изменился при копировании нулевого плана.')
        return {'sha256':original,'deleted_rows':[],'updated_rows':{},'sheet_xml':None,'changed_parts':[]}
    X='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    with ZipFile(source) as z:
        book=E.fromstring(z.read('xl/workbook.xml'))
        sheets=[s for s in book.find(X+'sheets') if s.get('name')=='УЗК']
        rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        ref=next(r.get('Target') for r in rels if r.get('Id')==sheets[0].get(R+'id'))
        sheet_path=ref.lstrip('/') if ref.startswith('/') else posixpath.normpath('xl/'+ref)
        tree=E.fromstring(z.read(sheet_path));data=tree.find(X+'sheetData')
        if deletions:
            unsupported=('mergeCells','tableParts','drawing','legacyDrawing','hyperlinks')
            complex_parts={t for t in unsupported if tree.find(X+t) is not None}
            if complex_parts:
                if not complex_parts <= {'drawing','legacyDrawing'}:
                    raise Stop('УЗК: удаление строк в сложной структуре требует отдельной проверки; исходник не изменён.')
                return _native_drawing_candidate(source,target,original,sheet_path,updates,deletions)
            for cf in tree.findall(X+'conditionalFormatting'):
                if any(E.QName(node).localname=='formula' for node in cf.iter()):raise Stop('УЗК: формулы условного форматирования зависят от строк; запись остановлена.')
            validations=tree.find(X+'dataValidations')
            if validations is not None:
                for validation in validations:
                    for node in validation:
                        if E.QName(node).localname.startswith('formula') and not str(node.text or '').startswith('"'):
                            raise Stop('УЗК: проверка данных содержит ссылки на строки; запись остановлена.')
            for name in z.namelist():
                if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                    other=E.fromstring(z.read(name))
                    for formula in other.iter(X+'f'):
                        if name==sheet_path or 'УЗК' in (formula.text or '').upper():
                            raise Stop('УЗК: найдены формулы со ссылками на удаляемые строки; запись остановлена.')
            for defined in book.iter(X+'definedName'):
                if 'УЗК' in (defined.text or '').upper():
                    if defined.get('name') not in ('_xlnm._FilterDatabase','_xlnm.Print_Area') or not re.fullmatch(r"(?:'УЗК'|УЗК)![A-Z$0-9:]+",defined.text or ''):
                        raise Stop('УЗК: именованные диапазоны зависят от удаляемых строк; запись остановлена.')
        # Sorting once avoids a full scan of all deleted rows for each cell.
        removed=sorted(deletions)
        removed_set=set(deletions)
        def shifted(no):return no-bisect_left(removed,no)
        for row in list(data):
            no=int(row.get('r'))
            if no in deletions:data.remove(row);continue
            if no in updates:
                cell=next((c for c in row if c.get('r')=='C'+str(no)),None)
                if cell is None:cell=E.SubElement(row,X+'c',r='C'+str(no))
                cell.attrib.pop('t',None)
                for element in list(cell):
                    if element.tag in (X+'f',X+'v',X+'is'):cell.remove(element)
                E.SubElement(cell,X+'v').text=str(updates[no])
            row.set('r',str(shifted(no)))
            for c in row:
                if c.get('r'):c.set('r',re.sub(r'\d+$',str(shifted(no)),c.get('r')))
        def range_shift(value):
            result=[]
            for token in value.split():
                match=re.fullmatch(r'(\$?[A-Z]+)(\$?)([0-9]+)(?::(\$?[A-Z]+)(\$?)([0-9]+))?',token)
                if not match:raise Stop('УЗК: сложный диапазон строки '+token)
                low=int(match[3]);high=int(match[6] or low)
                if low==1 and high==1048576:result.append(token);continue
                # A sparse set of removed rows must not trigger an O(range)
                # walk (Excel ranges can span over one million rows).
                new_low=low
                while new_low<=high and new_low in removed_set:new_low+=1
                if new_low>high:continue
                new_high=high
                while new_high>=new_low and new_high in removed_set:new_high-=1
                first=match[1]+match[2]+str(shifted(new_low))
                last=(match[4] or match[1])+(match[5] or '')+str(shifted(new_high))
                result.append(first+(':'+last if match[6] else ''))
            return ' '.join(result)
        for tag in ('dimension','autoFilter'):
            el=tree.find(X+tag)
            if el is not None and el.get('ref'):el.set('ref',range_shift(el.get('ref')))
        for cf in list(tree.findall(X+'conditionalFormatting')):
            ref=range_shift(cf.get('sqref',''))
            if ref:cf.set('sqref',ref)
            else:tree.remove(cf)
        validations=tree.find(X+'dataValidations')
        if validations is not None:
            for validation in list(validations):
                ref=range_shift(validation.get('sqref',''))
                if ref:validation.set('sqref',ref)
                else:validations.remove(validation)
            validations.set('count',str(len(validations)))
            if not len(validations):tree.remove(validations)
        changed_parts={sheet_path:E.tostring(tree,encoding='UTF-8',xml_declaration=True)}
        for defined in book.iter(X+'definedName'):
            if deletions and 'УЗК' in (defined.text or '').upper():
                prefix,ref=defined.text.split('!',1);changed=prefix+'!'+range_shift(ref)
                if changed!=defined.text:
                    defined.text=changed;changed_parts['xl/workbook.xml']=E.tostring(book,encoding='UTF-8',xml_declaration=True)
        with ZipFile(target,'w') as out:
            for item in z.infolist():out.writestr(item,changed_parts.get(item.filename,z.read(item.filename)))
    if sha(source)!=original:raise Stop('УЗК изменился при подготовке списания.')
    after=read_shared_uzk(target,project=project,marks=quantities)
    for mark,value in quantities.items():
        if value in (0,'0'):continue
        key=(_key(project),_key(mark));left=rows[key]['remaining']-int(value)
        if left and (key not in after or after[key]['remaining']!=left) or not left and key in after:
            raise Stop('УЗК: не подтверждено списание '+mark)
    for key,rec in rows.items():
        if key[0]!=_key(project) or not any(_key(m)==key[1] and q not in (0,'0') for m,q in quantities.items()):
            if key not in after or after[key]['remaining']!=rec['remaining']:raise Stop('УЗК: изменена другая марка.')
    return {'sha256':sha(target),'deleted_rows':deletions,'updated_rows':updates,'sheet_xml':sheet_path,'changed_parts':list(changed_parts)}

def _updated_native_filter(workbook_xml, old_sheet_xml, new_sheet_xml, removed):
    """Allow only the expected shrink of the УЗК sheet's built-in auto-filter."""
    import re
    from lxml import etree

    x='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    old_filter=etree.fromstring(old_sheet_xml).find(x+'autoFilter')
    new_filter=etree.fromstring(new_sheet_xml).find(x+'autoFilter')
    if old_filter is None or new_filter is None:
        return None
    extent=re.fullmatch(r'A1:C([1-9][0-9]*)',old_filter.get('ref',''))
    if extent is None or 1 in removed:
        return None
    end=int(extent[1])
    new_end=end-bisect_left(removed,end+1)
    if new_end==end or new_end<1:
        return None
    old_filter.set('ref',f'A1:C{new_end}')
    if etree.tostring(old_filter,method='c14n',exclusive=True)!=etree.tostring(
            new_filter,method='c14n',exclusive=True):
        return None

    workbook=etree.fromstring(workbook_xml)
    scope=next((str(i) for i,sheet in enumerate(workbook.find(x+'sheets'))
                if sheet.get('name')=='УЗК'),None)
    names=[node for node in workbook.findall(f'{x}definedNames/{x}definedName')
           if node.get('name')=='_xlnm._FilterDatabase' and node.get('localSheetId')==scope]
    if scope is None or len(names)!=1 or names[0].get('hidden')!='1':
        return None
    name=names[0]
    formula=re.fullmatch(r"((?:'УЗК'|УЗК)!\$A\$1:\$C\$)([1-9][0-9]*)",name.text or '')
    if formula is None or int(formula[2])!=end:
        return None
    name.text=formula[1]+str(new_end)
    return etree.tostring(workbook,encoding='UTF-8',xml_declaration=True)


def _native_drawing_candidate(source,target,original,sheet_path,updates,deletions):
    """Delete rows on a private Excel copy and accept only the verified delta."""
    from io import BytesIO
    from tempfile import TemporaryDirectory
    from zipfile import ZipFile
    from openpyxl.utils.cell import coordinate_from_string
    from lxml import etree
    from olefile import OleFileIO
    from passport_app.excel import excel_session,open_excel

    def cells(path):
        workbook=load_workbook(path,read_only=True,data_only=False,keep_links=False)
        try:
            return {sheet.title:{cell.coordinate:(cell.data_type,cell.value)
                                for row in sheet.iter_rows() for cell in row
                                if cell.value is not None} for sheet in workbook}
        finally:
            workbook.close()

    def vba_streams(data):
        with OleFileIO(BytesIO(data)) as ole:
            return {tuple(name):ole.openstream(name).read() for name in ole.listdir()}

    def workbook_semantics(data):
        root=etree.fromstring(data)
        # Excel rewrites location, collaboration metadata and window position
        # on Save As a private copy. Preserve the original bytes of all three.
        names={'absPath','revisionPtr','bookViews'}
        return [etree.tostring(node,method='c14n',exclusive=True)
                for node in root if etree.QName(node).localname not in names
                and not (etree.QName(node).localname=='AlternateContent'
                         and [etree.QName(child).localname for child in node.iter()][1:]
                         ==['Choice','absPath'])]

    before=cells(source)
    removed=sorted(deletions)
    removed_set=set(removed)
    def shifted(row):
        return row-bisect_left(removed,row)

    expected={}
    for address, value in before['УЗК'].items():
        column,row=coordinate_from_string(address)
        if row not in removed_set:
            expected[column+str(shifted(row))]=value
    for row,value in updates.items():
        expected['C'+str(shifted(row))]=('n',value)
    before['УЗК']=expected

    with TemporaryDirectory(prefix='uzk-',dir=target.parent) as folder:
        native=Path(folder)/'native.xlsm'
        verified=Path(folder)/'verified.xlsm'
        from shutil import copyfile
        copyfile(source,native)
        if sha(native)!=original:
            raise Stop('УЗК изменился при копировании для Excel.')
        app=None
        try:
            with excel_session() as app:
                book=open_excel(app,native)
                sheet=None
                try:
                    sheet=book.Worksheets('УЗК')
                    for row,left in updates.items():
                        sheet.Cells(row,3).Value2=left
                    for row in sorted(deletions,reverse=True):
                        sheet.Rows(row).Delete()
                    book.Save()
                finally:
                    sheet=None
                    book.Close(SaveChanges=False)
                    book=None
        finally:
            app=None

        if cells(native)!=before:
            raise Stop('УЗК: Excel изменил неожиданные ячейки, формулы или листы.')
        allowed={sheet_path,'xl/sharedStrings.xml'}
        with ZipFile(source) as old, ZipFile(native) as new:
            if set(old.namelist())!=set(new.namelist()) or new.testzip() is not None:
                raise Stop('УЗК: Excel изменил состав частей XLSM.')
            changed={n for n in old.namelist() if old.read(n)!=new.read(n)}
            incidental={'docProps/core.xml','xl/workbook.xml','xl/vbaProject.bin'}
            if changed-allowed-incidental:
                raise Stop('УЗК: изменились графика, элементы управления или другие части XLSM: '
                           +', '.join(sorted(changed-allowed-incidental)))
            workbook='xl/workbook.xml'
            workbook_edit=None
            if workbook in changed:
                workbook_before=old.read(workbook)
                workbook_after=workbook_semantics(new.read(workbook))
                if workbook_semantics(workbook_before)!=workbook_after:
                    workbook_edit=_updated_native_filter(workbook_before,old.read(sheet_path),
                                                        new.read(sheet_path),removed)
                    if workbook_edit is None or workbook_semantics(workbook_edit)!=workbook_after:
                        raise Stop('УЗК: Excel изменил настройки книги или именованные диапазоны.')
            vba='xl/vbaProject.bin'
            if vba in changed and vba_streams(old.read(vba))!=vba_streams(new.read(vba)):
                raise Stop('УЗК: Excel изменил содержимое VBA.')
            # Keep the original metadata and VBA. Patch the one verified
            # name alongside the worksheet when Excel shortens its filter.
            replacements={name:new.read(name) for name in changed & allowed}
            if workbook_edit is not None:
                replacements[workbook]=workbook_edit
            with ZipFile(verified,'w') as out:
                for item in old.infolist():
                    name=item.filename
                    out.writestr(item,replacements[name] if name in replacements else old.read(name))
        if cells(verified)!=before or sha(source)!=original:
            raise Stop('УЗК: проверка сохранённой копии после Excel не пройдена.')
        if target.exists():
            raise Stop('УЗК: кандидат XLSM уже существует.')
        verified.rename(target)
    return {'sha256':sha(target),'deleted_rows':deletions,'updated_rows':updates,
            'sheet_xml':sheet_path,'changed_parts':sorted(replacements)}
