from pathlib import Path
ROOT=Path(__file__).resolve().parents[2]
for name in ('lifecycle.py','e2e.py','negative.py'):
    text=(ROOT/'tests/stage9_091'/name).read_text(encoding='utf8')
    text=text.replace('source=ROOT,','source=SOURCE,').replace("Path(os.environ['LOCALAPPDATA'])/'VZMKPassport/workspaces'","Path(SETUP['base'])/'VZMKPassport/workspaces'")
    text=text.replace("vom=ROOT.parent/'Гараж казахстан/Паспорт/9 от 06.10.2026/ВОМ отгрузочный счет 183  (5).xls'","vom=VOM")
    text=text.replace("smoke['version']=='0.9.1'","smoke['version']=='0.9.2'")
    text=text.replace("native['word_documents']==4 and native['excel_documents']==3","native['word_documents']==3 and native['excel_documents']==3 and native['physical_files']==6 and native['document_records']==7 and native['word_records']==4")
    if name=='lifecycle.py':
        start=text.index('users={};old_owned={}')
        end=text.index("save(RUN/'USER_OFFICE_BEFORE.json',users)")
        text=text[:start]+"users={int(pid):r for pid,r in SETUP['user_office'].items()}\n"+text[end:]
    (ROOT/'tests/stage9_092'/name).write_text(text,encoding='utf8')
print('new isolated suites prepared')
