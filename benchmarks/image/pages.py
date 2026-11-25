"""Settings and helpers shared by the page-image benchmark's conversion and scoring.

The benchmark reuses the Word-to-PDF benchmark's corpus, fixtures, ink scorer, timing helpers and
README helpers from benchmarks/word, which this module puts on the import path.
"""
from hashlib import sha256
from pathlib import Path
import re
import sys

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
WORD = ROOT / 'benchmarks' / 'word'
if str(WORD) not in sys.path:
    # After this directory, so that this benchmark's own modules take precedence.
    sys.path.insert(1, str(WORD))

from fetch import BENCHMARK_FIXTURES  # noqa: E402

OUTPUT = ROOT / 'out' / 'benchmark-image'
# Page images are rendered at this resolution, then resampled to the ink scorer's 96 dpi.
DPI = 150
ENGINES = {
    'ours': '@kanera/docx2pdf',
    'libreoffice': 'LibreOffice + pdftoppm',
    'aspose': 'Aspose.Words (.NET)',
    'syncfusion': 'Syncfusion DocIO (.NET)',
    'rdocx': 'rdocx (Rust)',
}
# Engines that join the run only when their command-line flag is given.
OPTIONAL = {'aspose': 'aspose', 'syncfusion': 'syncfusion', 'rdocx': 'rdocx'}
# DocIO 28's RenderAsImages has no resolution setting and always renders at 96 dpi.
ENGINE_DPI = {'syncfusion': 96}
PAGE = re.compile(r'page-(\d+)\.png')


def engine_dpi(engine):
    return ENGINE_DPI.get(engine, DPI)


def page_files(directory):
    """The page-<n>.png files in directory, in page order."""
    if not directory.is_dir():
        return []
    pages = [(int(match[1]), path) for path in directory.iterdir() if (match := PAGE.fullmatch(path.name))]
    return [path for _, path in sorted(pages)]


def is_png(path):
    try:
        with path.open('rb') as file:
            return file.read(8) == b'\x89PNG\r\n\x1a\n'
    except OSError:
        return False


def select_fixtures(cases=None, limit=None):
    """The benchmark fixtures, or the named cases, or the first `limit` of the fixed timing order.

    The timing order sorts case IDs by SHA-256, so a small limit still mixes the suites.
    """
    fixtures = BENCHMARK_FIXTURES
    if cases:
        known = {fixture['case'] for fixture in fixtures}
        unknown = sorted(set(cases) - known)
        if unknown:
            raise SystemExit(f"Unknown cases: {', '.join(unknown)}")
        fixtures = [fixture for fixture in fixtures if fixture['case'] in cases]
    if limit:
        order = sorted(fixtures, key=lambda fixture: sha256(fixture['case'].encode()).hexdigest())
        chosen = {fixture['case'] for fixture in order[:limit]}
        fixtures = [fixture for fixture in fixtures if fixture['case'] in chosen]
    return fixtures
