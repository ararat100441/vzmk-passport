"""Microsoft Excel is the only writer of XLSM; macros/events/links are disabled."""
import hashlib
import io
from pathlib import Path
import shutil
import posixpath
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile

from .office import excel_session
from .storage import Stop, sha, digest

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def vba_streams(path):
    # Excel may reorder Compound File sectors on Save. Every actual stream,
    # including source and compiled VBA, must remain byte-for-byte identical.
    import olefile
    with ZipFile(path) as z:
        binary = z.read('xl/vbaProject.bin')
    with olefile.OleFileIO(io.BytesIO(binary)) as ole:
        return {'/'.join(n):hashlib.sha256(ole.openstream(n).read()).hexdigest() for n in ole.listdir()}


def sheet_part(path, title='УЗК'):
    with ZipFile(path) as z:
        tree = ET.fromstring(z.read('xl/workbook.xml'))
        matches = [s for s in tree.findall('s:sheets/s:sheet',NS) if s.get('name') == title]
        if len(matches) != 1:
            raise Stop('UZK_SHEET_MISSING_OR_DUPLICATE')
        identity = matches[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel = [r for r in rels if r.get('Id') == identity]
        if len(rel) != 1 or rel[0].get('TargetMode') == 'External':
            raise Stop('UZK_SHEET_RELATIONSHIP_INVALID')
        target = posixpath.normpath(posixpath.join('xl',rel[0].get('Target','')))
        if not target.startswith('xl/worksheets/') or target not in z.namelist():
            raise Stop('UZK_SHEET_TARGET_INVALID')
        return target


def package(path, ignored_cells=()):
    """All package parts are compared, with only explicit calculation metadata
    and target values excluded. Formula text and every style/structure remain.
    """
    result = {}
    try:
        with ZipFile(path) as z:
            if z.testzip() is not None:
                raise Stop('XLSM_CRC_ERROR')
            if 'xl/vbaProject.bin' not in z.namelist():
                raise Stop('VBA_PROJECT_MISSING')
            for name in sorted(z.namelist()):
                data = z.read(name)
                if name == 'xl/vbaProject.bin':
                    result[name] = digest(vba_streams(path))
                    continue
                if name == 'xl/calcChain.xml':
                    continue
                if name.endswith(('.xml', '.rels')):
                    tree = ET.fromstring(data)
                    for parent in tree.iter():
                        if name == 'xl/workbook.xml' and parent.tag.split('}')[-1] == 'absPath':
                            parent.attrib.pop('url', None)
                        if name == 'xl/workbook.xml' and parent.tag.split('}')[-1] == 'revisionPtr':
                            parent.attrib.pop('documentId', None)
                        for node in list(parent):
                            local = node.tag.split('}')[-1]
                            if ((name == 'docProps/core.xml' and local in ('modified', 'lastModifiedBy'))
                                    or (name == 'xl/workbook.xml' and local == 'calcPr')
                                    or (name == 'xl/workbook.xml' and local == 'functionGroups'
                                        and len(node) == 0 and node.get('builtInGroupCount') == '19')
                                    or (name in ('[Content_Types].xml', 'xl/_rels/workbook.xml.rels')
                                        and 'calcChain' in str(node.attrib))):
                                parent.remove(node)
                    if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                        for cell in tree.findall('.//s:c', NS):
                            if cell.find('s:f', NS) is not None:
                                cached = cell.find('s:v', NS)
                                if cached is not None:
                                    cell.remove(cached)
                                cell.attrib.pop('t', None)
                            if (name, cell.get('r')) in ignored_cells:
                                for child in list(cell):
                                    if child.tag.split('}')[-1] in ('v', 'is'):
                                        cell.remove(child)
                                cell.attrib.pop('t', None)
                    data = ET.canonicalize(ET.tostring(tree, encoding='unicode')).encode('utf8')
                result[name] = hashlib.sha256(data).hexdigest()
    except (BadZipFile, ET.ParseError, OSError) as exc:
        raise Stop('XLSM_PACKAGE_UNREADABLE') from exc
    return result


def inspect(path, expected_cells=None):
    before = sha(path)
    parts = package(path)
    with excel_session() as excel:
        book = excel.Workbooks.Open(Filename=str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True,
                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False, CorruptLoad=0)
        try:
            if not book.ReadOnly or int(book.FileFormat) != 52 or not book.HasVBProject:
                raise Stop('NATIVE_XLSM_FORMAT_OR_VBA_MISMATCH')
            for coord, expected in (expected_cells or {}).items():
                name, address = coord.split('!')
                if book.Worksheets(name).Range(address).Value2 != expected:
                    raise Stop('SAVED_DEDUCTION_MISMATCH: '+coord)
            formulas, sheets = {}, []
            for s in book.Worksheets:
                sheets.append({'name': s.Name, 'visible': int(s.Visible), 'rows': int(s.UsedRange.Rows.Count),
                               'columns': int(s.UsedRange.Columns.Count)})
                for row in s.UsedRange.Rows:
                    for c in row.Cells:
                        if c.HasFormula:
                            formulas[s.Name+'!'+c.Address] = {'formula': c.Formula, 'value': c.Value2}
                            if str(c.Text) in ('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A', '#NUM!', '#NULL!'):
                                raise Stop('NATIVE_FORMULA_ERROR: '+s.Name+'!'+c.Address)
            info = {'status': 'PASS', 'excel_version': str(excel.Version), 'format': 52,
                    'normal_open': True, 'read_only': True, 'macros_disabled': int(excel.AutomationSecurity) == 3,
                    'has_vba': bool(book.HasVBProject), 'vba_streams_sha256': parts['xl/vbaProject.bin'],
                    'vba_streams':vba_streams(path),
                    'sheets': sheets, 'formulas': formulas, 'sha256': before}
        finally:
            book.Close(SaveChanges=False)
    if sha(path) != before:
        raise Stop('NATIVE_READ_CHANGED_FILE')
    return info


def stage(registry, output, deductions, fault=None):
    output = Path(output)
    if output.exists():
        raise Stop('STAGE_FILE_EXISTS')
    shutil.copyfile(registry, output)
    part = sheet_part(registry)
    ignored = {(part, 'C'+str(d['row'])) for d in deductions}
    before = package(registry, ignored)
    with excel_session() as excel:
        book = excel.Workbooks.Open(Filename=str(output.resolve()), UpdateLinks=0, ReadOnly=False,
                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False, CorruptLoad=0)
        try:
            if book.ReadOnly or int(book.FileFormat) != 52 or not book.HasVBProject:
                raise Stop('STAGE_XLSM_NOT_WRITABLE')
            if int(book.Connections.Count) or int(book.Worksheets('УЗК').QueryTables.Count):
                raise Stop('EXTERNAL_DATA_FORBIDDEN')
            sheet = book.Worksheets('УЗК')
            if sheet.ProtectContents:
                raise Stop('UNSUPPORTED_UZK_SHEET_OR_PROTECTION')
            for d in deductions:
                c = sheet.Cells(d['row'], 3)
                if c.HasFormula or c.MergeCells or c.Value2 != d['raw_before']:
                    raise Stop('STAGE_TARGET_CHANGED_OR_NOT_LITERAL')
                c.Value2 = d['after']
            if fault:
                fault('before_save', output)
            excel.CalculateFull()
            if fault:
                fault('save_io', output)
            book.Save()
            if not book.Saved:
                raise Stop('EXCEL_SAVE_NOT_COMPLETE')
        finally:
            book.Close(SaveChanges=False)
    if fault:
        # Test seam for a torn native-save result, before any validation/publish.
        fault('during_save', output)
    after = package(output, ignored)
    differences = sorted(n for n in set(before) | set(after) if before.get(n) != after.get(n))
    if differences:
        raise Stop('XLSM_UNEXPECTED_PACKAGE_CHANGE: '+', '.join(differences))
    native = inspect(output, {'УЗК!C'+str(d['row']):d['after'] for d in deductions})
    return {'status': 'PASS', 'all_package_parts_preserved': True, 'allowed_changes': deductions,
            'vba_streams_sha256': after['xl/vbaProject.bin'], 'native': native, 'sha256': sha(output),
            'ignored_save_metadata':['core modified/lastModifiedBy','workbook calcPr/calcChain',
                'save directory absPath','coauthor documentId','empty built-in functionGroups=19',
                'formula caches (formula expressions remain exact)'],
            'vba_container_repacked':_vba_container_sha(registry)!=_vba_container_sha(output)}


def _vba_container_sha(path):
    with ZipFile(path) as z:
        return hashlib.sha256(z.read('xl/vbaProject.bin')).hexdigest()
