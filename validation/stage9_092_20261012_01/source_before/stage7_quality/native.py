"""Normal read-only Microsoft Word verification; EMF/PNG, never PDF."""
from pathlib import Path
import subprocess
from .provenance import Stage7Error, sha, write


def inspect(path, directory, *, kind, required):
    from passport_app.vik import word_session, open_word
    directory = Path(directory) / ('native_' + kind)
    directory.mkdir(exist_ok=False)
    before = sha(path)
    with word_session() as word:
        doc = open_word(word, Path(path), readonly=True)
        try:
            doc.Repaginate()
            doc.ActiveWindow.View.Type = 3
            pages = doc.ComputeStatistics(2)
            text = doc.Content.Text
            if pages != 2 or any(value not in text for value in required):
                raise Stage7Error('NATIVE_WORD_PAGES_OR_REQUIRED_TEXT: ' + kind)
            result = {'status': 'PASS', 'path': str(path), 'sha256': before,
                'open': 'NORMAL; OpenAndRepair=False; ReadOnly=True; Close(False)',
                'pages': pages, 'sections': doc.Sections.Count, 'tables': doc.Tables.Count,
                'readonly': bool(doc.ReadOnly), 'text': text, 'page_texts': [],
                'list_labels': [], 'fields': [], 'version': word.Version, 'build': word.Build,
                'hyperlinks': doc.Hyperlinks.Count, 'pdf_export': False,
                'page_width': doc.PageSetup.PageWidth, 'page_height': doc.PageSetup.PageHeight}
            for i in range(1, pages + 1):
                start = doc.GoTo(1, 1, i).Start
                end = doc.GoTo(1, 1, i + 1).Start if i < pages else doc.Content.End
                result['page_texts'].append(doc.Range(start, end).Text)
                (directory / f'page-{i}.emf').write_bytes(bytes(doc.ActiveWindow.Panes(1).Pages(i).EnhMetaFileBits))
            if kind == 'QUALITY':
                if doc.Sections.Count != 1 or doc.Tables.Count != 2:
                    raise Stage7Error('NATIVE_QUALITY_STRUCTURE_CHANGED')
                for i in range(1, doc.Paragraphs.Count + 1):
                    p = doc.Paragraphs(i)
                    label = p.Range.ListFormat.ListString
                    if label:
                        result['list_labels'].append({'label': label, 'text': p.Range.Text})
                labels = [v['label'].rstrip('.') for v in result['list_labels']]
                if [str(i) for i in range(1, 15)] != [x for x in labels if x.isdigit()]:
                    raise Stage7Error('NATIVE_LOGICAL_NUMBERING_CHANGED')
                for i in range(1, doc.Fields.Count + 1):
                    f = doc.Fields(i)
                    result['fields'].append({'code': f.Code.Text, 'result': f.Result.Text})
                if len(result['fields']) != 1 or 'MERGEFIELD' not in result['fields'][0]['code']:
                    raise Stage7Error('NATIVE_MERGEFIELD_CHANGED')
                if 'Журавлев А.Н.' not in result['page_texts'][1]:
                    raise Stage7Error('NATIVE_SIGNATURE_PAGE_CHANGED')
        finally:
            doc.Close(False)
    if before != sha(path):
        raise Stage7Error('NATIVE_OPEN_CHANGED_SOURCE')
    result['unchanged_after_close'] = True
    write(directory / 'WORD.json', result)
    command = ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
               str(Path(__file__).with_name('render_emf.ps1')), '-Directory', str(directory),
               '-PageWidth', str(result['page_width']), '-PageHeight', str(result['page_height'])]
    rendered = subprocess.run(command, capture_output=True, text=True, timeout=60)
    if rendered.returncode or not all((directory / f'page-{i}.png').is_file() for i in (1, 2)):
        raise Stage7Error('NATIVE_EMF_RENDER_FAILED: ' + rendered.stderr)
    return {'status': 'PASS', 'evidence': str(directory / 'WORD.json'),
            'evidence_sha256': sha(directory / 'WORD.json'), 'pages': pages,
            'images': [str(directory / f'page-{i}.png') for i in (1, 2)]}


def open_readonly(path):
    """Explicit user action opens a dedicated normal read-only Word window."""
    import win32com.client
    from .provenance import allowed_source
    word = win32com.client.DispatchEx('Word.Application')
    word.AutomationSecurity = 3
    word.Options.UpdateLinksAtOpen = False
    from passport_app.vik import open_word
    doc = open_word(word, allowed_source(path), readonly=True)
    word.Visible = True
    return word, doc
