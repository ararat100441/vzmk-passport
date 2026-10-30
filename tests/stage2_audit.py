"""Read-only source audit; Word opens only a separate validation copy."""
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from passport_app.excel import sha256


def main():
    import pythoncom
    import win32com.client
    base = Path(__file__).resolve().parents[1]
    run = base / 'validation' / ('stage2_' + datetime.now().strftime('%Y%m%d_%H%M%S'))
    run.mkdir()
    manifest = json.loads((base / 'REVIEW_MANIFEST.json').read_text(encoding='utf-8'))
    template = base.parent / 'Гараж казахстан/Паспорт/1_2026 - ВИК.docx'
    protected = {str((base / p).resolve()): h for p, h in manifest['code_sha256'].items()}
    protected.update(manifest['evidence_sha256'])
    protected.update(manifest['production_sha256'])
    protected[str(base / 'REVIEW_MANIFEST.json')] = sha256(base / 'REVIEW_MANIFEST.json')
    for folder in [base / 'e2e_user_stage1', template.parent]:
        for p in folder.rglob('*'):
            if p.is_file() and not p.name.startswith('~$') and '__pycache__' not in p.parts:
                protected[str(p)] = sha256(p)
    before = {p: sha256(p) for p in protected}
    assert before == protected, 'Unexpected Stage 1 checkpoint changes'
    (run / 'protected_before.json').write_text(json.dumps(before, ensure_ascii=False, indent=2), encoding='utf-8')
    shutil.copy2(template, run / 'template.docx')
    shutil.copy2(base / 'e2e_user_stage1/ВЗМК.xlsx', run / 'ВЗМК.xlsx')
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        doc = word.Documents.Open(str(run / 'template.docx'), ReadOnly=True, AddToRecentFiles=False, OpenAndRepair=False)
        doc.Repaginate()
        result = {'source': str(template), 'sha256': sha256(template), 'pages': doc.ComputeStatistics(2),
                  'tables': doc.Tables.Count, 'sections': doc.Sections.Count, 'paragraphs': []}
        for i in range(1, doc.Paragraphs.Count + 1):
            r = doc.Paragraphs(i).Range
            result['paragraphs'].append({'index': i, 'start': r.Start, 'end': r.End, 'page': r.Information(3), 'text': r.Text})
        # Word's native print-layout page metafiles, no PDF creation.
        doc.ActiveWindow.View.Type = 3
        for i in range(1, result['pages'] + 1):
            data = doc.ActiveWindow.Panes(1).Pages(i).EnhMetaFileBits
            (run / f'template-page-{i}.emf').write_bytes(bytes(data))
        (run / 'template_audit.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
        doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
    assert {p: sha256(p) for p in protected} == before
    print(run)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
