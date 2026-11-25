"""Render the Word benchmark's DOCX fixtures to PNG page images, cold and warm, one engine at a time.

The timing protocol, sampling, timeouts and GNU time measurement are the Word-to-PDF benchmark's
(benchmarks/word/convert.py), whose helpers this script imports. See README.md.
"""
from argparse import ArgumentParser
from datetime import datetime, timezone
from hashlib import file_digest, sha256
import json
import math
import os
from pathlib import Path
import shutil
import statistics
import subprocess
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

from pages import DPI, HERE, OUTPUT, ROOT, WORD, engine_dpi, is_png, page_files, select_fixtures
from fetch import OUTPUT as WORD_OUTPUT, benchmark_lock, child_environment, verify
from convert import (FULL_TIMING_MS, GNU_TIME, SAMPLE_SIZE, TIMEOUT_S, has_reference, host_cpu, file_digests,
                     link_or_copy, measure, read_batch, reference_digests, sample_cases, version, write_log)
from libreoffice_warm import batch_times

# Raise when the measurement changes, so that cached results are measured again.
CACHE_FORMAT = 1
# Kanera is always measured; the other engines' results and pages are reused (see cache_key).
UNCACHED = frozenset({'ours'})
# A warm batch renders every timed document in one process. Page images take longer than PDFs, so
# the batch gets more time than the Word benchmark's five minutes.
WARM_BATCH_TIMEOUT_S = 30 * 60
RUN_DIRECTORIES = ('', 'kanera-only')
RUNTIMES = {'aspose': 'dotnet', 'syncfusion': 'dotnet', 'libreoffice': 'poppler'}
ASPOSE = OUTPUT / 'aspose-bin' / 'AsposeImageBenchmark.dll'
SYNCFUSION = OUTPUT / 'syncfusion-bin' / 'SyncfusionImageBenchmark.dll'
# LibreOffice exports only the first page with --convert-to png, so the all-pages route is a PDF
# export followed by pdftoppm, timed together in one shell.
LIBREOFFICE = ('soffice -env:UserInstallation="$1" --headless --convert-to pdf:writer_pdf_Export --outdir "$2" "$3"'
               ' && pdftoppm -r "$4" -png "$2/input.pdf" "$5/page"')


def converted(row):
    return 'error' not in row and not row.get('skipped')


def source(name):
    return WORD_OUTPUT / 'cases' / name / 'input.docx'


def normalize_pages(directory):
    """Rename the engine's <prefix>-<n>.png files to page-<n>.png and return them in page order."""
    for path in directory.glob('*.png'):
        stem, _, number = path.stem.rpartition('-')
        if stem and number.isdigit() and path.name != f'page-{int(number)}.png':
            path.rename(directory / f'page-{int(number)}.png')
    return page_files(directory)


def check_pages(directory, label):
    pages = normalize_pages(directory)
    if not pages:
        raise RuntimeError(f'{label} wrote no page images')
    if not all(map(is_png, pages)):
        raise RuntimeError(f'{label} wrote an invalid PNG')
    return pages


def cold_command(engine, input_path, target, work, workers):
    """The command that renders one document's pages in a fresh process into target."""
    if engine == 'ours':
        return ['node', ROOT / 'dist' / 'cli.js', input_path, '--png', target, '--dpi', str(DPI), '--quiet']
    if engine == 'libreoffice':
        # Each LibreOffice invocation uses a separate profile and output directory.
        work.mkdir(parents=True)
        return ['sh', '-c', LIBREOFFICE, 'sh', (work / 'profile').as_uri(), work, input_path, str(DPI), target]
    if engine == 'rdocx':
        return [*workers[engine], input_path, target, str(DPI)]
    return [*workers[engine], input_path, target]


def batch_command(engine, jobs_file, workers):
    if engine == 'ours':
        return ['node', HERE / 'kanera.mjs', '--batch', jobs_file, str(DPI)]
    if engine == 'rdocx':
        return [*workers[engine], '--batch', jobs_file, str(DPI)]
    return [*workers[engine], '--batch', jobs_file]


def convert_document(engine, fixture, work, output, workers, *, timed=True, env=None):
    """Render one document in a fresh process. Untimed rows stay out of the speed and memory figures."""
    name = fixture['case']
    input_path = source(name)
    verify(input_path, fixture['docx_sha256'])
    target = output / 'cases' / name / engine
    shutil.rmtree(target, ignore_errors=True)
    target.mkdir(parents=True)
    scratch = work / engine / name
    command = cold_command(engine, input_path, target, scratch, workers)
    flags = {} if timed else {'untimed': True}
    start = perf_counter()
    try:
        metrics, result = measure(command, work / 'rss.txt', TIMEOUT_S, env)
        pages = check_pages(target, f'{name} ({engine})')
    except (subprocess.CalledProcessError, RuntimeError) as error:
        stderr = getattr(error, 'stderr', None)
        message = (stderr or '').strip()[-2000:] or str(error)
        shutil.rmtree(target, ignore_errors=True)
        summary = next((line.strip() for line in message.splitlines() if line.strip()), str(error))
        print(f'{name} {engine} FAILED: {summary[:200]}', flush=True)
        return {'case': name, 'engine': engine, 'cold_ms': (perf_counter() - start) * 1000,
                'peak_rss_kib': None, **flags, 'error': message}
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    details = {'pages': len(pages), 'png_bytes': sum(page.stat().st_size for page in pages)}
    if result.stderr.strip():
        details['stderr'] = result.stderr.strip()[-2000:]
    print(f"{name} {engine} {'cold' if timed else 'untimed'} {metrics['cold_ms']:.1f} ms, "
          f"{metrics['peak_rss_kib'] / 1024:.1f} MiB, {len(pages)} pages", flush=True)
    return {'case': name, 'engine': engine, **metrics, **flags, **details}


def convert_engine(engine, fixtures, sample, work, output, workers, env=None):
    """Cold conversions of one engine, one process at a time: the timing sample first, then the others."""
    rows = {}
    first = [fixture for fixture in fixtures if fixture['case'] in sample]
    for fixture in first:
        rows[fixture['case']] = convert_document(engine, fixture, work, output, workers, env=env)
    rest = [fixture for fixture in fixtures if fixture['case'] not in sample]
    estimate = statistics.fmean(rows[fixture['case']]['cold_ms'] for fixture in first) * len(fixtures)
    timed = estimate <= FULL_TIMING_MS
    if rest and not timed:
        # A slow engine is timed on the sample only and renders the other documents untimed for scoring.
        for fixture in rest:
            if not has_reference(fixture):
                rows[fixture['case']] = {'case': fixture['case'], 'engine': engine, 'skipped': True}
                shutil.rmtree(output / 'cases' / fixture['case'] / engine, ignore_errors=True)
        rest = [fixture for fixture in rest if has_reference(fixture)]
        print(f'{engine}: about {estimate / 60_000:.1f} min for all {len(fixtures)} documents, so timed on '
              f'the {len(first)}-document sample; converting the other {len(rest)} untimed', flush=True)
    for fixture in rest:
        rows[fixture['case']] = convert_document(engine, fixture, work, output, workers, timed=timed, env=env)
    return [rows[fixture['case']] for fixture in fixtures]


def libreoffice_warm(cases, work, by_case, env=None, max_restarts=5):
    """Warm LibreOffice + pdftoppm: one LibreOffice batch writes the PDFs (timed between successive
    PDF completions, as in the Word benchmark), then pdftoppm rasterizes each PDF in its own process,
    as in the cold route. A document's warm time is the sum of the two."""
    remaining = cases
    for attempt in range(max_restarts + 1):
        attempt_dir = work / f'attempt-{attempt}'
        attempt_dir.mkdir()
        jobs = [(str(source(name)), str(attempt_dir / f'{i}.pdf')) for i, name in enumerate([remaining[0], *remaining])]
        times = batch_times(jobs, attempt_dir, env=env)
        if any(not math.isfinite(t) or t <= 0 for t in times):
            raise ValueError('Invalid warm timings for libreoffice')
        timed = remaining[:max(0, len(times) - 1)]
        for index, name in enumerate(timed, 1):
            pages = attempt_dir / f'pages-{index}'
            pages.mkdir()
            start = perf_counter()
            result = subprocess.run(['pdftoppm', '-r', str(DPI), '-png', jobs[index][1], pages / 'page'],
                                    capture_output=True, text=True, timeout=TIMEOUT_S, env=env)
            elapsed = (perf_counter() - start) * 1000
            try:
                if result.returncode:
                    raise RuntimeError(result.stderr.strip() or f'pdftoppm exited with status {result.returncode}')
                check_pages(pages, f'{name} (libreoffice warm)')
                by_case[name]['warm_ms'] = times[index] + elapsed
            except RuntimeError as error:
                by_case[name]['warm_error'] = str(error)[-2000:]
            shutil.rmtree(pages)
        if len(timed) == len(remaining):
            return
        # The batch stopped on the warm-up or on the next document after those it timed.
        stopped = remaining[len(timed)]
        by_case[stopped]['warm_error'] = 'warm batch stopped on this document'
        remaining = remaining[len(timed) + 1:]
        if not remaining:
            return
    raise RuntimeError(f'libreoffice warm batch stopped more than {max_restarts} times')


def measure_warm(engine, rows, workspace, workers, env=None):
    by_case = {row['case']: row for row in rows}
    work = workspace / f'warm-{engine}'
    work.mkdir()
    try:
        # Only time documents the engine converted, timed, in the cold run.
        cases = [row['case'] for row in rows if converted(row) and not row.get('untimed')]
        if not cases:
            return
        print(f'{engine}: warm-up and {len(cases)} warm conversions', flush=True)
        if engine == 'libreoffice':
            libreoffice_warm(cases, work, by_case, env=env)
            return
        # Render the first document once without recording it, then time each document in order.
        jobs = [(str(source(name)), str(work / str(i))) for i, name in enumerate([cases[0], *cases])]
        jobs_file = work / 'jobs.json'
        jobs_file.write_text(json.dumps(jobs))
        times = read_batch(batch_command(engine, jobs_file, workers), WARM_BATCH_TIMEOUT_S, env)
        if len(times) != len(jobs) or any(not math.isfinite(t) or t <= 0 for t in times):
            raise ValueError(f'Invalid warm timings for {engine}')
        for _, target in jobs:
            try:
                check_pages(Path(target), f'{engine} warm')
            except RuntimeError as error:
                raise ValueError(f'Invalid warm conversion output: {error}') from None
        for name, ms in zip(cases, times[1:], strict=True):
            by_case[name]['warm_ms'] = ms
    finally:
        # Warm pages are only checked, not kept.
        shutil.rmtree(work, ignore_errors=True)


def renderer_digest():
    digest = sha256()
    for path in (HERE / 'kanera.mjs', *sorted((ROOT / 'dist').rglob('*.js'))):
        digest.update(str(path.relative_to(ROOT)).encode())
        digest.update(b'\0')
        with path.open('rb') as file:
            digest.update(file_digest(file, 'sha256').digest())
    return digest.hexdigest()


def adapter_files(engine):
    """The runner sources and lock files that one engine renders through."""
    if engine == 'libreoffice':
        return [WORD / 'libreoffice_warm.py']
    if engine in ('aspose', 'syncfusion'):
        return sorted(path for path in (HERE / engine).iterdir()
                      if path.suffix in ('.cs', '.csproj') or path.name == 'packages.lock.json')
    if engine == 'rdocx':
        return [HERE / 'rdocx_runner.py', WORD / 'rust' / 'requirements.txt']
    raise ValueError(f'No adapter files listed for {engine}')


def cache_key(engine, versions, fixtures, cpu):
    """A digest of everything one engine's cached results depend on."""
    runtime = {name: versions.get(name) for name in (engine, RUNTIMES.get(engine), 'time') if name}
    if engine == 'rdocx':
        runtime['python'] = sys.version
    key = {
        'format': CACHE_FORMAT,
        'engine': engine,
        'version': runtime,
        'adapter': file_digests(adapter_files(engine)),
        'inputs': [[fixture['case'], fixture['docx_sha256'], reference_digests(fixture)] for fixture in fixtures],
        'cpu': cpu,
        'settings': {'dpi': engine_dpi(engine), 'sample_size': SAMPLE_SIZE, 'full_timing_ms': FULL_TIMING_MS,
                     'timeout_s': TIMEOUT_S, 'warm_timeout_s': WARM_BATCH_TIMEOUT_S},
    }
    return sha256(json.dumps(key, sort_keys=True).encode()).hexdigest()


def has_pages(directory, row):
    pages = page_files(directory)
    return len(pages) == row['pages'] and all(map(is_png, pages))


def restore_pages(source_dir, target):
    shutil.rmtree(target, ignore_errors=True)
    target.mkdir(parents=True)
    for page in page_files(source_dir):
        # Conversions replace page directories instead of rewriting files, so a hard link is safe.
        link_or_copy(page, target / page.name)


def load_cache(directory, key, engine, fixtures, output):
    """Restore one engine's cached rows and pages, or return None when the entry is missing or incomplete."""
    try:
        saved = json.loads((directory / 'results.json').read_text())
        rows = saved['rows']
        if (saved['key'] != key or [row['case'] for row in rows] != [fixture['case'] for fixture in fixtures]
                or any(row['engine'] != engine for row in rows)):
            return None
    except (OSError, KeyError, TypeError, ValueError):
        return None
    if not all(has_pages(directory / 'pages' / row['case'], row) for row in rows if converted(row)):
        return None
    for row in rows:
        target = output / 'cases' / row['case'] / engine
        shutil.rmtree(target, ignore_errors=True)
        if converted(row):
            restore_pages(directory / 'pages' / row['case'], target)
    return saved


def save_cache(directory, key, rows, output):
    """Keep one engine's rows and pages, replacing that engine's older cache entries."""
    staging = directory.with_name(directory.name + '.partial')
    shutil.rmtree(staging, ignore_errors=True)
    (staging / 'pages').mkdir(parents=True)
    for row in rows:
        if converted(row):
            restore_pages(output / 'cases' / row['case'] / row['engine'], staging / 'pages' / row['case'])
    measured_at = datetime.now(timezone.utc).isoformat()
    (staging / 'results.json').write_text(
        json.dumps({'key': key, 'measured_at': measured_at, 'rows': rows}, indent=2) + '\n')
    engine = directory.name.rsplit('-', 1)[0]
    for old in directory.parent.iterdir():
        if old != staging and old.is_dir() and old.name.rsplit('-', 1)[0] == engine:
            shutil.rmtree(old)
    staging.rename(directory)


def load_checkpoint(checkpoint, output, settings, engines):
    saved = json.loads(checkpoint.read_text())
    if not isinstance(saved, dict) or 'engines' not in saved:
        raise ValueError('Checkpoint was written by an earlier version of convert_pages.py')
    for key, value in settings.items():
        if saved.get(key) != value:
            raise ValueError(f'Checkpoint {key} differs from the current run')
    progress = saved['engines']
    if not set(progress) <= set(engines):
        raise ValueError('Checkpoint has engines that this run does not include')
    for case, digest in settings['inputs'].items():
        verify(source(case), digest)
    for engine, entry in progress.items():
        if [row['case'] for row in entry['rows']] != list(settings['inputs']):
            raise ValueError(f'Checkpoint does not cover the fixtures for {engine}')
        for row in entry['rows']:
            target = output / 'cases' / row['case'] / engine
            if row['engine'] != engine or converted(row) and not has_pages(target, row):
                raise ValueError(f'Invalid cold output {target}')
    return progress


def run_conversions(args, versions, workers, fixtures, root):
    output = root / ('kanera-only' if args.kanera_only else '')
    output.mkdir(parents=True, exist_ok=True)
    log = output / 'conversion.json'
    checkpoint = output / 'conversion-cold.json'
    settings = {
        'format': CACHE_FORMAT,
        'versions': versions,
        'inputs': {fixture['case']: fixture['docx_sha256'] for fixture in fixtures},
        'renderer_build': renderer_digest(),
        'dpi': DPI,
    }
    engines = ('ours',) if args.kanera_only else ('ours', 'libreoffice', *workers)
    # The checkpoint records each engine's cold rows and whether its warm pass finished, so a
    # stopped run resumes with the engine it stopped on.
    progress = {}
    if not args.fresh and not log.exists() and checkpoint.exists():
        try:
            progress = load_checkpoint(checkpoint, output, settings, engines)
        except (KeyError, OSError, TypeError, ValueError) as error:
            raise SystemExit(f'Cannot resume from {checkpoint}: {error}. '
                             'Retry with the original settings or use --fresh.') from None
        if args.no_cache:
            progress = {engine: entry for engine, entry in progress.items() if not entry.get('cached')}
        print(f"Resuming from {checkpoint} with {', '.join(progress) or 'no engines'} converted", flush=True)
    else:
        # Invalidate old logs before replacing any cold pages.
        checkpoint.unlink(missing_ok=True)
        log.unlink(missing_ok=True)

    cpu = host_cpu()
    sample = sample_cases(fixtures)
    cache_root = root / 'cache'
    with TemporaryDirectory(prefix='convert-', dir=output) as temporary:
        work = Path(temporary)
        # Converters keep their temporary files in this run's work directory, not in /tmp.
        env = child_environment(work / 'tmp')
        for engine in engines:
            entry = progress.get(engine)
            cache = key = None
            if engine not in UNCACHED:
                key = cache_key(engine, versions, fixtures, cpu)
                cache = cache_root / f'{engine}-{key[:16]}'
            if entry is None and cache and not args.no_cache:
                saved = load_cache(cache, key, engine, fixtures, output)
                if saved is not None:
                    progress[engine] = {'rows': saved['rows'], 'warm': True, 'cached': saved['measured_at']}
                    write_log(checkpoint, work, dict(settings, engines=progress))
                    print(f"{engine}: reusing {cache} from {saved['measured_at']} (--no-cache measures it again)",
                          flush=True)
                    continue
            if entry is None:
                rows = convert_engine(engine, fixtures, sample, work, output, workers, env)
                entry = progress[engine] = {'rows': rows, 'warm': False}
                write_log(checkpoint, work, dict(settings, engines=progress))
            if not entry['warm']:
                try:
                    measure_warm(engine, entry['rows'], work, workers, env)
                except Exception:
                    print(f'Cold results saved in {checkpoint}. Retry the same command without --fresh to resume.',
                          file=sys.stderr, flush=True)
                    raise
                # Saved as soon as the engine finishes, so a stopped run keeps it.
                if cache:
                    save_cache(cache, key, entry['rows'], output)
                entry['warm'] = True
                write_log(checkpoint, work, dict(settings, engines=progress))

        write_log(log, work, {
            'completed_at': datetime.now(timezone.utc).isoformat(),
            'versions': versions,
            'dpi': {engine: engine_dpi(engine) for engine in engines},
            'cases': [fixture['case'] for fixture in fixtures],
            'results': [row for engine in engines for row in progress[engine]['rows']],
            'cached': {engine: progress[engine]['cached'] for engine in engines if progress[engine].get('cached')},
        })


def poppler_version():
    result = subprocess.run(['pdftoppm', '-v'], check=True, capture_output=True, text=True)
    return (result.stderr or result.stdout).splitlines()[0]


def main():
    parser = ArgumentParser(description='Render the pinned DOCX fixtures to page images, cold and warm.')
    standalone = parser.add_mutually_exclusive_group()
    standalone.add_argument('--kanera-only', action='store_true', help='Run Kanera alone.')
    standalone.add_argument('--all', action='store_true', help='Run all comparison engines.')
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose.Words .NET runner.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO .NET runner.')
    parser.add_argument('--rdocx', action='store_true', help='Include rdocx (Rust, through its Python bindings).')
    parser.add_argument('--fresh', action='store_true', help='Start cold conversions even if a warm pass is pending.')
    parser.add_argument('--no-cache', action='store_true',
                        help='Measure every engine again instead of reusing cached results.')
    parser.add_argument('--limit', type=int, help='Use only the first N fixtures of the fixed timing order (smoke tests).')
    parser.add_argument('--cases', help='Use only these comma-separated case IDs (smoke tests).')
    parser.add_argument('--output', type=Path, help=f'Output directory (default: {OUTPUT.relative_to(ROOT)}).')
    args = parser.parse_args()
    if args.all:
        args.aspose = args.syncfusion = args.rdocx = True
    if args.kanera_only and (args.aspose or args.syncfusion or args.rdocx):
        parser.error('--kanera-only cannot be combined with comparison engine flags')
    if args.limit is not None and args.limit < 1:
        parser.error('--limit must be positive')
    fixtures = select_fixtures(args.cases.split(',') if args.cases else None, args.limit)
    for fixture in fixtures:
        if not source(fixture['case']).is_file():
            parser.error(f'Run benchmarks/word/fetch.py first: {source(fixture["case"])} is missing.')
    if not (ROOT / 'dist' / 'cli.js').is_file():
        parser.error('Run pnpm build first.')
    versions = {
        'time': version([GNU_TIME, '--version']).splitlines()[0],
        'node': version(['node', '--version']),
        'ours': json.loads((ROOT / 'package.json').read_text())['version'],
    }
    if not args.kanera_only:
        versions['libreoffice'] = version(['soffice', '--version'])
        versions['poppler'] = poppler_version()
    workers = {}
    if args.aspose:
        license_path = os.environ.get('ASPOSE_WORDS_LICENSE')
        if not license_path or not Path(license_path).is_file():
            parser.error('Set ASPOSE_WORDS_LICENSE to a local Aspose.Words for .NET license.')
        if not ASPOSE.is_file():
            parser.error(f'Build the Aspose image runner into {ASPOSE.parent} (see README.md).')
        workers['aspose'] = ['dotnet', ASPOSE]
    if args.syncfusion:
        if not os.environ.get('SYNCFUSION_LICENSE_KEY'):
            parser.error('Set SYNCFUSION_LICENSE_KEY to a Syncfusion license key for the pinned version.')
        if not SYNCFUSION.is_file():
            parser.error(f'Build the Syncfusion image runner into {SYNCFUSION.parent} (see README.md).')
        workers['syncfusion'] = ['dotnet', SYNCFUSION]
    for worker in workers.values():
        versions.update(json.loads(version([*worker, '--version'])))
    if args.rdocx:
        runner = [sys.executable, '-I', HERE / 'rdocx_runner.py']
        try:
            versions.update(json.loads(version([*runner, '--version'])))
        except subprocess.CalledProcessError:
            parser.error('Install benchmarks/word/rust/requirements.txt into the benchmark venv.')
        workers['rdocx'] = runner

    root = args.output.resolve() if args.output else OUTPUT
    with benchmark_lock(root, stale=[(root / directory, 'convert-') for directory in RUN_DIRECTORIES]):
        run_conversions(args, versions, workers, fixtures, root)


if __name__ == '__main__':
    main()
