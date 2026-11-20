"""Compare the current renderer with the pinned Hugging Face and vendored reference files."""

from argparse import ArgumentParser
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import ExitStack
import csv
from datetime import datetime, timezone
from functools import partial
from hashlib import file_digest, sha256
import json
from os import cpu_count, environ, getuid
from pathlib import Path
from queue import SimpleQueue
from shutil import copyfile, copyfileobj, copytree, which
from socket import socketpair
from statistics import mean
import subprocess
import sys
from tempfile import NamedTemporaryFile, TemporaryDirectory, TemporaryFile, gettempdir
from time import sleep
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import urlopen
from zipfile import BadZipFile, ZipFile

import numpy
import PIL
from PIL import Image

from fetch import CORPUS, HERE, MANIFEST, UPSTREAM, UPSTREAM_DIR, load_upstream
from ink import DPI, INK_THRESHOLD, INK_TOLERANCE_PX, page_scores, render, render_reference
from readme import update_accuracy_readme

ROOT = HERE.parents[1]
REVISION = MANIFEST['revision']
MANIFEST_SHA256 = '1682abfccf34d12f9da8dfff9f18087580e90065f0c3397d83321657a6458848'
CACHE = Path(gettempdir()) / f'kanera-docx-word-{getuid()}' / REVISION
REFERENCE_CACHE = ROOT / 'out' / 'word-reference-rasters'
CONVERSION_TIMEOUT = 180
EXPECTED_DOCUMENTS = 2925
# Word exported the corpus's tracked changes and comments in markup view (struck-through deletions,
# comment balloons on a shrunken page). docx2pdf renders the final text without comments, so only
# the `clean` pairs are scored. These clean pairs are in scripts no bundled font covers; users
# supply fonts for them.
UNBUNDLED_SCRIPTS = {
    '521bb53c59': 'Devanagari', '92f6ee8aac': 'Devanagari', 'e3d87c5d40': 'Devanagari',
    'a9477e435a': 'Khmer', '55fcbe9086': 'Thaana', 'de4654ccf5': 'Thai',
}
SCORED_DOCUMENTS = 2472
SOURCES = {
    'huggingface': {
        'source': MANIFEST['source'], 'revision': REVISION, 'documents': SCORED_DOCUMENTS,
    },
    'upstream': {
        'source': UPSTREAM['source'], 'revision': UPSTREAM['revision'],
        'classic_reference_revision': UPSTREAM['classic_references']['revision'],
        'documents': len(UPSTREAM['fixtures']),
    },
}
METRIC = {
    'name': 'document mean ink F1',
    'dpi': DPI,
    'ink_threshold': INK_THRESHOLD,
    'tolerance_px': INK_TOLERANCE_PX,
}


def checksum(path):
    with path.open('rb') as file:
        return file_digest(file, 'sha256').hexdigest()


def cache_file(cache, path, digest):
    target = cache / path
    if target.is_file() and checksum(target) == digest:
        return target

    target.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(3):
        try:
            url = CORPUS + quote(path)
            # A failed CDN redirect can be cached. Retry using a fresh resolve URL.
            if attempt:
                url += f'?download=true&retry={attempt}'
            with NamedTemporaryFile(dir=target.parent, delete=False) as partial_file:
                temporary = Path(partial_file.name)
                try:
                    with urlopen(url, timeout=60) as response:
                        copyfileobj(response, partial_file)
                    partial_file.close()
                    if checksum(temporary) != digest:
                        raise ValueError('SHA-256 mismatch')
                    temporary.replace(target)
                    return target
                finally:
                    temporary.unlink(missing_ok=True)
        except (URLError, TimeoutError, ConnectionError, ValueError) as error:
            if attempt == 2:
                raise RuntimeError(f'Could not download {path}: {error}') from error
            delay = 2 ** attempt
            if isinstance(error, HTTPError) and error.code == 429:
                retry_after = error.headers.get('Retry-After', '')
                delay = int(retry_after) + 1 if retry_after.isdecimal() else 300
            sleep(delay)


def load_corpus(cache):
    manifest = cache_file(cache, 'MANIFEST.sha256.json', MANIFEST_SHA256)
    digests = json.loads(manifest.read_text(encoding='utf-8'))['files']
    index = cache_file(cache, 'documents.csv', digests['documents.csv'])
    with index.open(encoding='utf-8', newline='') as file:
        rows = list(csv.DictReader(file))
    if len(rows) != EXPECTED_DOCUMENTS or len({row['id'] for row in rows}) != EXPECTED_DOCUMENTS:
        raise ValueError('Unexpected document index size or duplicate IDs')
    for row in rows:
        if digests[row['docx']] != row['sha256'] or row['pdf'] not in digests:
            raise ValueError(f"Inconsistent source manifest for {row['stem']}")
        row['reference_application'] = 'Microsoft Word for Mac'
        row['pdf_sha256'] = digests[row['pdf']]
    rows = [row for row in rows if row['state'] == 'clean' and row['id'] not in UNBUNDLED_SCRIPTS]
    if len(rows) != SCORED_DOCUMENTS:
        raise ValueError('Unexpected number of clean documents in the index')
    return rows, digests


def fetch_pair(row, cache, digests):
    for kind in ('docx', 'pdf'):
        path = row[kind]
        cache_file(cache, path, digests[path])


def run(command, **options):
    result = subprocess.run(command, capture_output=True, text=True, **options)
    if result.returncode:
        message = (result.stderr or result.stdout).strip()
        raise RuntimeError(message or f'{command[0]} exited with code {result.returncode}')
    return result


def build_renderer(workspace):
    # Keep the production build and its fonts isolated from changes in the checkout.
    for name in ('package.json', 'tsconfig.json', 'tsconfig.build.json'):
        copyfile(ROOT / name, workspace / name)
    for name in ('src', 'fonts'):
        copytree(ROOT / name, workspace / name)
    copyfile(HERE / 'accuracy-worker.mjs', workspace / 'accuracy-worker.mjs')
    (workspace / 'node_modules').symlink_to(ROOT / 'node_modules', target_is_directory=True)
    subprocess.run(['pnpm', '--config.verify-deps-before-run=false', 'run', 'build'],
                   cwd=workspace, stdout=sys.stderr, check=True)


def stop_worker(process):
    process.kill()
    process.wait()


class RendererWorker:
    def __init__(self, workspace):
        self.command = ['node', workspace / 'accuracy-worker.mjs']
        self.resources = None

    def start(self):
        with ExitStack() as resources:
            self.channel, child = socketpair()
            resources.enter_context(self.channel)
            self.diagnostics = resources.enter_context(TemporaryFile())
            with child:
                process = subprocess.Popen(self.command, stdin=child, stdout=child, stderr=self.diagnostics)
            resources.callback(stop_worker, process)
            self.channel.settimeout(CONVERSION_TIMEOUT)
            self.responses = resources.enter_context(self.channel.makefile('rb'))
            self.resources = resources.pop_all()

    def close(self):
        if self.resources is not None:
            self.resources.close()
            self.resources = None

    def convert(self, source, candidate):
        try:
            if self.resources is None:
                self.start()
            self.channel.sendall(json.dumps([str(source), str(candidate)]).encode() + b'\n')
            reply = self.responses.readline()
            if not reply:
                self.diagnostics.seek(0)
                message = self.diagnostics.read().decode(errors='replace').strip()
                raise RuntimeError(message or 'Node conversion worker exited without a response')
            response = json.loads(reply)
            if not isinstance(response, dict):
                raise RuntimeError('Invalid Node conversion worker response')
            if 'error' in response:
                raise RuntimeError(response['error'])
        except TimeoutError as error:
            self.close()
            raise subprocess.TimeoutExpired(self.command, CONVERSION_TIMEOUT) from error
        except (OSError, RuntimeError, ValueError) as error:
            self.close()
            raise RuntimeError(str(error)) from error


def reference_settings(environment):
    # Poppler can substitute local fonts for fonts missing from a PDF.
    renderer = Path(which('pdftoppm')).resolve()
    paths = {renderer}
    if which('fc-list'):
        paths.update(Path(name) for name in run(['fc-list', '-f', '%{file}\n']).stdout.splitlines() if name)
    if which('fc-conflist'):
        for line in run(['fc-conflist']).stdout.splitlines():
            if line.startswith('+ '):
                paths.add(Path(line[2:].split(':', 1)[0]))
    stats = []
    for path in sorted(paths):
        if path.is_file():
            stat = path.stat()
            stats.append((str(path), stat.st_size, stat.st_mtime_ns))
    font_env = [environ.get(key) for key in ('FONTCONFIG_FILE', 'FONTCONFIG_PATH', 'XDG_CONFIG_HOME')]
    # Version the on-disk cache alongside the rendering settings.
    settings = (1, DPI, environment['poppler'], environment['pillow'], stats, font_env)
    return sha256(json.dumps(settings).encode()).hexdigest()


def cached_reference(cache, settings, directory, row, work):
    source = (row['pdf_sha256'] if 'pdf' in row else
              [row['reference_dpi'], [page['sha256'] for page in row['reference_images']]])
    key = sha256(json.dumps([settings, source]).encode()).hexdigest()
    entry = cache / f'{key}.zip'
    try:
        with ZipFile(entry) as archive:
            pages = [work / f'reference-{number}.png' for number in range(1, len(archive.infolist()) + 1)]
            if pages and archive.namelist() == [page.name for page in pages]:
                for page in pages:
                    with archive.open(page.name) as packed, page.open('wb') as target:
                        copyfileobj(packed, target)
                return pages
    except (FileNotFoundError, BadZipFile):
        pass

    with TemporaryDirectory(dir=work) as raw_dir:
        raw = render_reference(directory, row, Path(raw_dir), raster_format='ppm')
        if not raw:
            raise RuntimeError(f"No reference pages for {row['id']}")
        pages = []
        with NamedTemporaryFile(dir=cache, delete=False) as temporary:
            staged = Path(temporary.name)
            try:
                with ZipFile(temporary, 'w') as archive:
                    for number, image in enumerate(raw, 1):
                        page = work / f'reference-{number}.png'
                        with Image.open(image) as source:
                            source.convert('L').save(page, compress_level=1)
                        archive.write(page, page.name)
                        pages.append(page)
                # Publish all pages together, even if two runs render the same reference.
                temporary.close()
                staged.replace(entry)
            finally:
                staged.unlink(missing_ok=True)
    return pages


def score_document(row, corpus, directory, workspace, raster_cache, settings, workers):
    with TemporaryDirectory(dir=workspace) as temporary:
        work = Path(temporary)
        reference = cached_reference(raster_cache, settings, directory, row, work)
        candidate = work / 'candidate.pdf'
        result = {
            'corpus': corpus, 'suite': row.get('suite', 'full'), 'id': row['id'],
            'document': row['docx'],
            'reference': row['pdf'] if 'pdf' in row else [image['path'] for image in row['reference_images']],
            'reference_application': row['reference_application'],
            'reference_format': 'pdf' if 'pdf' in row else 'png',
            'reference_pages': len(reference),
            'pages': 0, 'ink_overlap_pct': 0.0, 'error': None,
        }
        try:
            worker = workers.get()
            try:
                worker.convert(directory / row['docx'], candidate)
            finally:
                workers.put(worker)
            pages = render(candidate, work / 'candidate', raster_format='ppm')
        except (RuntimeError, subprocess.SubprocessError) as error:
            result['error'] = str(error)
            return result
        if not pages:
            result['error'] = 'No pages rasterized from candidate PDF'
            return result
        result.update(pages=len(pages), ink_overlap_pct=100 * mean(page_scores(reference, pages)))
        return result


def process_all(action, documents, jobs, label):
    with ThreadPoolExecutor(max_workers=jobs) as pool:
        futures = {pool.submit(action, row): row['id'] for row in documents}
        for i, future in enumerate(as_completed(futures), 1):
            try:
                yield future.result()
            except Exception as error:
                for pending in futures:
                    pending.cancel()
                raise RuntimeError(f'{label} failed for {futures[future]}: {error}') from error
            if i % 100 == 0 or i == len(futures):
                print(f'{label} {i}/{len(futures)}', file=sys.stderr, flush=True)


def summarize(results):
    return {
        'documents': len(results),
        'reference_pages': sum(row['reference_pages'] for row in results),
        'conversion_failures': sum(row['error'] is not None for row in results),
        'matching_page_counts': sum(row['pages'] == row['reference_pages'] for row in results),
        'accuracy_pct': mean(row['ink_overlap_pct'] for row in results),
    }


def main():
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--corpus', choices=('all', *SOURCES), default='all',
                        help='Reference corpus to run (default: %(default)s)')
    parser.add_argument('--cache-dir', type=Path, default=CACHE,
                        help='Directory containing the Hugging Face word/ files (default: %(default)s)')
    parser.add_argument('--reference-cache-dir', type=Path, default=REFERENCE_CACHE,
                        help='Directory for rendered reference pages (default: %(default)s)')
    parser.add_argument('--out', type=Path, default=ROOT / 'out/word-accuracy.json',
                        help='JSON report path (default: %(default)s)')
    parser.add_argument('--baseline', type=Path,
                        help='Report from a previous run for the percentage-point comparison')
    parser.add_argument('--jobs', type=int, default=min(8, cpu_count() or 1),
                        help='Concurrent conversions (default: %(default)s)')
    args = parser.parse_args()
    if args.jobs < 1:
        parser.error('--jobs must be positive')

    sources = SOURCES if args.corpus == 'all' else {args.corpus: SOURCES[args.corpus]}
    baseline = None
    if args.baseline:
        baseline = json.loads(args.baseline.read_text(encoding='utf-8'))
        baseline_sources = baseline.get('sources')
        if baseline_sources is None:
            baseline_sources = {'huggingface': {
                key: baseline.get(key) for key in SOURCES['huggingface']
            }}
        if baseline_sources != sources or baseline.get('metric') != METRIC:
            parser.error('--baseline must use the same corpora, revisions, and ink metric')

    poppler = run(['pdftoppm', '-v'])
    environment = {
        'node': run(['node', '--version']).stdout.strip(),
        'poppler': (poppler.stderr or poppler.stdout).splitlines()[0],
        'numpy': numpy.__version__,
        'pillow': PIL.__version__,
    }
    raster_cache = args.reference_cache_dir.resolve()
    raster_cache.mkdir(parents=True, mode=0o700, exist_ok=True)
    settings = reference_settings(environment)
    corpora = {}
    if 'upstream' in sources:
        print('Verifying the vendored corpus', file=sys.stderr, flush=True)
        corpora['upstream'] = (UPSTREAM_DIR, load_upstream())
    if 'huggingface' in sources:
        cache = args.cache_dir.resolve()
        cache.mkdir(parents=True, mode=0o700, exist_ok=True)
        print(f'Verifying the Hugging Face corpus in {cache}', file=sys.stderr, flush=True)
        documents, digests = load_corpus(cache)
        fetch = partial(fetch_pair, cache=cache, digests=digests)
        for _ in process_all(fetch, documents, 16, 'Verified Hugging Face'):
            pass
        corpora['huggingface'] = (cache, documents)

    with TemporaryDirectory(prefix='kanera-docx-accuracy-') as temporary, ExitStack() as resources:
        workspace = Path(temporary)
        print('Building the current renderer', file=sys.stderr, flush=True)
        build_renderer(workspace)
        workers = SimpleQueue()
        for _ in range(args.jobs):
            worker = RendererWorker(workspace)
            resources.callback(worker.close)
            workers.put(worker)
        results = []
        by_corpus = {}
        for corpus, (directory, documents) in corpora.items():
            score = partial(score_document, corpus=corpus, directory=directory, workspace=workspace,
                            raster_cache=raster_cache, settings=settings, workers=workers)
            scored = sorted(process_all(score, documents, args.jobs, f'Scored {corpus}'),
                            key=lambda row: row['id'])
            results.extend(scored)
            by_corpus[corpus] = summarize(scored)
    completed_at = datetime.now(timezone.utc)
    report = {
        'completed_at': completed_at.isoformat(),
        'sources': sources,
        'metric': METRIC,
        'environment': environment,
        **summarize(results),
        'by_corpus': by_corpus,
        'results': results,
    }
    if baseline is not None:
        report['change_pp'] = report['accuracy_pct'] - baseline['accuracy_pct']
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
    if args.corpus == 'all':
        update_accuracy_readme(HERE.parents[1], report, completed_at.date().isoformat())
    print(f"Accuracy {report['accuracy_pct']:.2f}% across {report['documents']} documents "
          f"and {report['reference_pages']} reference pages ({report['conversion_failures']} conversion failures).")
    if baseline is not None:
        print(f"Change from baseline {report['change_pp']:+.2f} percentage points.")
    print(f'Report written to {args.out.resolve()}')
    return int(report['conversion_failures'] > 0)


if __name__ == '__main__':
    try:
        sys.exit(main())
    except (OSError, RuntimeError, ValueError, subprocess.SubprocessError) as error:
        sys.exit(str(error))
