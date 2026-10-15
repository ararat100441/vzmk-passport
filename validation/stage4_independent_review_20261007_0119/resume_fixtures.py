"""Resume transaction checks after correcting a review-only path test fixture."""
from pathlib import Path
source=Path(__file__).with_name('review_fixtures.py').read_text(encoding='utf-8')
prefix=source[:source.index('# Real offline renderer/build:')]
tail=source[source.index("for path in (ROOT/'Профили.xlsx'"):]
exec(compile(prefix,str(Path(__file__).with_name('review_fixtures.py')),'exec'))
results=read(OUT/'fixture_results.json')
exec(compile(tail,str(Path(__file__).with_name('review_fixtures.py')),'exec'))
