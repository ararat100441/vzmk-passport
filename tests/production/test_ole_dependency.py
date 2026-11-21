from pathlib import Path
import unittest


class UzkDependency(unittest.TestCase):
    def test_olefile_is_a_direct_runtime_requirement(self):
        requirements = Path(__file__).resolve().parents[2] / 'requirements.txt'
        import re
        entries = [re.split(r'[<>=!~\s;]', line.strip(), maxsplit=1)[0].casefold()
                   for line in requirements.read_text().splitlines()]
        self.assertIn('olefile', entries)
