"""Native Excel read-only reopen of already saved validation files; never Save."""
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from stage3_evidence import product_hashes, timestamp, evidence_root, bind_result
PRODUCT_BEFORE, STARTED_AT = product_hashes(), timestamp()
from passport_app.excel import sha256, excel_session
from passport_app.catalog import HEADERS


def main():
    checkpoint = json.loads((ROOT / 'STAGE3_CHECKPOINT.json').read_text(encoding='utf-8'))
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    out = evidence_root(checkpoint)
    results = []
    with excel_session() as excel:
        for name, expected in [('creation', [HEADERS, ('К2', '20Ш1', None, None), ('Б4-1', 'Гнз120x60x6', None, None)]),
                               ('conflict_True', [HEADERS, ('К1', '35Б2', 'RAL 7004', 0)])]:
            path = out / name / 'Документы' / 'Гараж казахстан' / 'Справочник марок.xlsx'
            before = sha256(path)
            book = excel.Workbooks.Open(Filename=str(path), UpdateLinks=0, ReadOnly=True,
                                        IgnoreReadOnlyRecommended=True, AddToMru=False, CorruptLoad=0)
            try:
                assert book.ReadOnly and book.Worksheets.Count == 1 and book.FileFormat == 51
                sheet = book.Worksheets(1)
                actual = list(sheet.Range(f'A1:D{len(expected)}').Value2)
                assert actual == expected, repr(actual)
                assert sheet.AutoFilter.Range.Address == f'$A$1:$D${len(expected)}'
                results.append({'status': 'PASS', 'path': str(path), 'readonly': True,
                                'values': actual, 'sha256': before, 'native_normal_reopen': True})
            finally:
                book.Close(SaveChanges=False)
            assert sha256(path) == before
    assert sha256(checkpoint['production_path']) == checkpoint['production_sha256'], 'STOP: production changed'
    bound = bind_result({'status': 'PASS', 'checks': results, 'pass_count': len(results), 'fail_count': 0}, PRODUCT_BEFORE, STARTED_AT)
    (out / 'native_reopen_results.json').write_text(json.dumps(bound, ensure_ascii=False, indent=2), encoding='utf-8')
    print('Native Excel read-only reopen: 2/2 PASS')


if __name__ == '__main__':
    main()
