from concurrent.futures import ThreadPoolExecutor
from hashlib import sha256
from pathlib import Path
from urllib.request import urlopen
import json

HERE = Path(__file__).resolve().parent
OUTPUT = HERE.parents[1] / 'out' / 'benchmark-word'
MANIFEST = json.loads((HERE / 'fixtures.json').read_text())
CORPUS = f"{MANIFEST['source']}/resolve/{MANIFEST['revision']}/word/"
UPSTREAM_DIR = HERE / 'upstream'
UPSTREAM = json.loads((UPSTREAM_DIR / 'fixtures.json').read_text(encoding='utf-8'))
BENCHMARK_FIXTURES = [
    dict(fixture, corpus='huggingface', suite='sample', reference_application='Microsoft Word for Mac')
    for fixture in MANIFEST['fixtures']
] + [
    dict(fixture, corpus='upstream', case='upstream-' + fixture['id'].replace(' ', '-'))
    for fixture in UPSTREAM['fixtures']
]


def verify(path, digest):
    data = path.read_bytes()
    if sha256(data).hexdigest() != digest:
        raise ValueError(f'SHA-256 mismatch for {path}')
    return data


def verify_reference(directory, fixture):
    if 'pdf' in fixture:
        verify(directory / fixture['pdf'], fixture['pdf_sha256'])
    else:
        for image in fixture['reference_images']:
            verify(directory / image['path'], image['sha256'])


def load_upstream():
    documents = UPSTREAM['fixtures']
    if len({row['id'] for row in documents}) != len(documents):
        raise ValueError('Duplicate IDs in the vendored fixture manifest')
    for row in documents:
        verify(UPSTREAM_DIR / row['docx'], row['docx_sha256'])
        verify_reference(UPSTREAM_DIR, row)
    return documents


def benchmark_reference(fixture):
    if fixture['corpus'] == 'upstream':
        return UPSTREAM_DIR, fixture
    return OUTPUT / 'cases' / fixture['case'], dict(fixture, pdf='word.pdf')


def fetch(fixture):
    case = OUTPUT / 'cases' / fixture['case']
    case.mkdir(parents=True, exist_ok=True)
    if fixture['corpus'] == 'upstream':
        data = verify(UPSTREAM_DIR / fixture['docx'], fixture['docx_sha256'])
        verify_reference(UPSTREAM_DIR, fixture)
        target = case / 'input.docx'
        if target.exists():
            verify(target, fixture['docx_sha256'])
        else:
            target.write_bytes(data)
        return fixture['case']

    for kind, filename in [('docx', 'input.docx'), ('pdf', 'word.pdf')]:
        target = case / filename
        digest = fixture[kind + '_sha256']
        if target.exists():
            verify(target, digest)
            continue
        with urlopen(CORPUS + fixture[kind], timeout=60) as response:
            data = response.read()
        if sha256(data).hexdigest() != digest:
            raise ValueError(f'SHA-256 mismatch for {fixture[kind]}')
        target.write_bytes(data)
    return fixture['case']


if __name__ == '__main__':
    with ThreadPoolExecutor(max_workers=6) as pool:
        for name in pool.map(fetch, BENCHMARK_FIXTURES):
            print(f'Verified {name}')
