from pathlib import Path
import sys
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from profiles_stage4.windows_generation import ProtectedFile
folder = Path(__file__).resolve().parent/'windows_probe'
folder.mkdir(exist_ok=True)
a, b, c = (folder/name for name in ['a.txt','b.txt','c.txt'])
a.write_bytes(b'old'); b.write_bytes(b'external')
with ProtectedFile(a) as handle:
    assert handle.read() == b'old'
    try:
        a.write_bytes(b'overwritten')
    except PermissionError:
        pass
    else:
        raise AssertionError('External write not blocked')
    try:
        handle.rename(b)
    except ValueError:
        pass
    else:
        raise AssertionError('Existing external destination overwritten')
    handle.rename(c)
    assert handle.read() == b'old'
assert c.read_bytes() == b'old' and b.read_bytes() == b'external'
print('WINDOWS HANDLE PROTECTION + OWN RENAME + NO_REPLACE: PASS')
