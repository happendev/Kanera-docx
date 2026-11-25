from argparse import ArgumentParser
from datetime import datetime, timezone
from hashlib import file_digest, sha256
import json
import math
import os
from pathlib import Path
import platform
import shutil
import statistics
import subprocess
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

from fetch import BENCHMARK_FIXTURES, HERE, OUTPUT, benchmark_lock, child_environment, verify
from libreoffice_warm import batch_times

GNU_TIME = '/usr/bin/time'
OFFICE2PDF = OUTPUT / 'office2pdf-bin' / 'office2pdf'
MINIPDF = OUTPUT / 'minipdf-bin' / 'MiniPdfBenchmark.dll'
# A failed or timed-out conversion is recorded as a zero-score case instead of stopping the run.
TIMEOUT_S = 300
WARM_TIMEOUT_S = 300
# Each engine converts a fixed sample first, cold and timed. When the sample's mean cold time puts
# the whole corpus past FULL_TIMING_MS, the engine is timed (cold, peak RSS, warm) on the sample
# only and converts the other scored documents untimed.
SAMPLE_SIZE = 50
FULL_TIMING_MS = 5 * 60_000
# Each comparison engine's results and PDFs are kept here and reused by later runs with the same
# cache key (see cache_key). Our converter is always measured.
CACHE = OUTPUT / 'cache'
UNCACHED = frozenset({'ours'})
# Raise when the measurement changes, so that cached results are measured again.
CACHE_FORMAT = 1
# The output directories of the full comparison and the standalone modes. Each run works in a
# convert-* directory inside one of them.
RUN_DIRECTORIES = ('', 'kanera-only', 'minipdf-only')
# Engines that fetch resources such as fonts on first use get one untimed pass over the documents
# before any timed conversion. None of the current engines do: EigenPal, dxpdf, rdocx and
# office2pdf read installed or packaged fonts, and the .NET runners and LibreOffice download nothing.
PRIME = frozenset()
RUNTIMES = {'eigenpal': 'node', 'aspose': 'dotnet', 'syncfusion': 'dotnet', 'minipdf': 'dotnet'}


def version(command):
    return subprocess.check_output(command, text=True).strip()


def measure(command, rss_file, timeout=None, env=None):
    if timeout:
        # coreutils timeout waits for the converter, so GNU time still sees its peak RSS.
        command = ['timeout', '-k', '5', str(timeout), *command]
    start = perf_counter()
    result = subprocess.run(
        [GNU_TIME, '-f', '%M', '-o', rss_file, *command],
        capture_output=True,
        text=True,
        check=True,
        env=env,
    )
    ms = (perf_counter() - start) * 1000
    rss_kib = int(rss_file.read_text())
    if rss_kib <= 0:
        raise ValueError(f'No peak RSS recorded for {command[0]}')
    return {'cold_ms': ms, 'peak_rss_kib': rss_kib}, result


def office2pdf_command(workspace, output, sources):
    return [OFFICE2PDF, '-j', '1', '--outdir', output, *sources]


def converted(row):
    return 'error' not in row and not row.get('skipped')


def has_reference(fixture):
    # Scoring compares a document with its reference PDF or page images.
    return 'pdf_sha256' in fixture or bool(fixture.get('reference_images'))


def sample_cases(fixtures):
    """The fixed timing sample: the SAMPLE_SIZE case IDs with the lowest SHA-256 hex digests."""
    names = sorted((fixture['case'] for fixture in fixtures), key=lambda name: sha256(name.encode()).hexdigest())
    return set(names[:SAMPLE_SIZE])


def cold_command(engine, input_path, target, work, workers):
    """The command that converts one document in a fresh process, and the path of the PDF it writes."""
    if engine in ('ours', 'eigenpal'):
        return ['node', HERE / 'convert.mjs', engine, input_path, target], target
    if engine == 'libreoffice':
        # Each LibreOffice invocation uses a separate profile and output directory.
        work.mkdir(parents=True)
        return [
            'soffice', '-env:UserInstallation=' + (work / 'profile').as_uri(),
            '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', work, input_path,
        ], work / 'input.pdf'
    if engine == 'office2pdf':
        return [*workers[engine], '-j', '1', '-o', target, input_path], target
    return [*workers[engine], input_path, target], target


def convert_document(engine, fixture, work, output, workers, *, timed=True, env=None):
    """Convert one document in a fresh process. Untimed rows stay out of the speed and memory figures."""
    name = fixture['case']
    input_path = OUTPUT / 'cases' / name / 'input.docx'
    verify(input_path, fixture['docx_sha256'])
    directory = output / 'cases' / name
    directory.mkdir(parents=True, exist_ok=True)
    target = directory / f'{engine}.pdf'
    target.unlink(missing_ok=True)
    command, written = cold_command(engine, input_path, target, work / engine / name, workers)
    flags = {} if timed else {'untimed': True}
    start = perf_counter()
    try:
        metrics, result = measure(command, work / 'rss.txt', TIMEOUT_S, env)
        if written != target and written.is_file():
            written.replace(target)
        if not target.is_file():
            raise RuntimeError(f'{name} ({engine}) did not produce a PDF')
        with target.open('rb') as pdf:
            if pdf.read(5) != b'%PDF-':
                raise RuntimeError(f'{name} ({engine}) wrote an invalid PDF')
    except (subprocess.CalledProcessError, RuntimeError) as error:
        stderr = getattr(error, 'stderr', None)
        message = (stderr or '').strip()[-2000:] or str(error)
        target.unlink(missing_ok=True)
        summary = next((line.strip() for line in message.splitlines() if line.strip()), str(error))
        print(f'{name} {engine} FAILED: {summary[:200]}', flush=True)
        return {'case': name, 'engine': engine, 'cold_ms': (perf_counter() - start) * 1000,
                'peak_rss_kib': None, **flags, 'error': message}
    details = json.loads(result.stdout) if engine in ('ours', 'eigenpal') else {}
    if result.stderr.strip():
        details['stderr'] = result.stderr.strip()
    print(f"{name} {engine} {'cold' if timed else 'untimed'} {metrics['cold_ms']:.1f} ms, "
          f"{metrics['peak_rss_kib'] / 1024:.1f} MiB", flush=True)
    return {'case': name, 'engine': engine, **metrics, **flags, **details}


def prime(engine, fixtures, work, workers, env=None):
    """One untimed conversion of each document, so that no timing includes a first-use download."""
    directory = work / f'prime-{engine}'
    directory.mkdir()
    print(f'{engine}: priming with {len(fixtures)} untimed conversions', flush=True)
    for fixture in fixtures:
        name = fixture['case']
        command, _ = cold_command(engine, OUTPUT / 'cases' / name / 'input.docx', directory / f'{name}.pdf',
                                  directory / name, workers)
        # A failure here shows up again in the timed conversion.
        subprocess.run(['timeout', '-k', '5', str(TIMEOUT_S), *command], capture_output=True, env=env)
    shutil.rmtree(directory, ignore_errors=True)


def convert_engine(engine, fixtures, sample, work, output, workers, env=None):
    """Cold conversions of one engine, one process at a time: the sample first, then the others."""
    if engine in PRIME:
        prime(engine, fixtures, work, workers, env)
    rows = {}
    first = [fixture for fixture in fixtures if fixture['case'] in sample]
    for fixture in first:
        rows[fixture['case']] = convert_document(engine, fixture, work, output, workers, env=env)
    rest = [fixture for fixture in fixtures if fixture['case'] not in sample]
    estimate = statistics.fmean(rows[fixture['case']]['cold_ms'] for fixture in first) * len(fixtures)
    timed = estimate <= FULL_TIMING_MS
    if rest and not timed:
        # A slow engine is timed on the sample only. It converts the other documents untimed for
        # scoring and skips any document that nothing scores.
        for fixture in rest:
            if not has_reference(fixture):
                rows[fixture['case']] = {'case': fixture['case'], 'engine': engine, 'skipped': True}
                (output / 'cases' / fixture['case'] / f'{engine}.pdf').unlink(missing_ok=True)
        rest = [fixture for fixture in rest if has_reference(fixture)]
        print(f'{engine}: about {estimate / 60_000:.1f} min for all {len(fixtures)} documents, so timed on '
              f'the {len(first)}-document sample; converting the other {len(rest)} untimed', flush=True)
    for fixture in rest:
        rows[fixture['case']] = convert_document(engine, fixture, work, output, workers, timed=timed, env=env)
    return [rows[fixture['case']] for fixture in fixtures]


def read_batch(command, timeout=WARM_TIMEOUT_S, env=None):
    result = subprocess.run(command, capture_output=True, text=True, timeout=timeout, env=env)
    if result.returncode:
        raise RuntimeError(f'Warm benchmark worker failed: {result.stderr.strip()}')
    return json.loads(result.stdout)


def resumable_warm(engine, cases, work, by_case, max_restarts=5, env=None):
    # A CLI batch can exit partway (LibreOffice occasionally does). Record what it timed, leave the
    # document it stopped on without a warm time, and continue with a fresh batch after it.
    command = (office2pdf_command, 'office2pdf') if engine == 'office2pdf' else ()
    remaining = cases
    for attempt in range(max_restarts + 1):
        attempt_dir = work / f'attempt-{attempt}'
        attempt_dir.mkdir()
        jobs = [(str(OUTPUT / 'cases' / name / 'input.docx'), str(attempt_dir / f'{i}.pdf'))
                for i, name in enumerate([remaining[0], *remaining])]
        times = batch_times(jobs, attempt_dir, *command, env=env)
        if any(not math.isfinite(t) or t <= 0 for t in times):
            raise ValueError(f'Invalid warm timings for {engine}')
        for _, target in jobs[:len(times)]:
            with Path(target).open('rb') as pdf:
                if pdf.read(5) != b'%PDF-':
                    raise ValueError(f'Invalid warm conversion output for {engine}')
        timed = remaining[:max(0, len(times) - 1)]
        for name, ms in zip(timed, times[1:]):
            by_case[name]['warm_ms'] = ms
        if len(timed) == len(remaining):
            return
        # The batch stopped on the warm-up or on the next document after those it timed.
        stopped = remaining[len(timed)]
        by_case[stopped]['warm_error'] = 'warm batch stopped on this document'
        remaining = remaining[len(timed) + 1:]
        if not remaining:
            return
    raise RuntimeError(f'{engine} warm batch stopped more than {max_restarts} times')


def measure_warm(engine, rows, workspace, workers, env=None):
    by_case = {row['case']: row for row in rows}
    work = workspace / f'warm-{engine}'
    work.mkdir()
    # Only time documents the engine converted, timed, in the cold run.
    cases = [row['case'] for row in rows if converted(row) and not row.get('untimed')]
    if not cases:
        return
    # Run the first document once without recording it, then time each document in order.
    jobs = [(str(OUTPUT / 'cases' / name / 'input.docx'), str(work / f'{i}.pdf'))
            for i, name in enumerate([cases[0], *cases])]
    print(f'{engine}: warm-up and {len(cases)} warm conversions', flush=True)
    if engine in ('libreoffice', 'office2pdf'):
        resumable_warm(engine, cases, work, by_case, env=env)
        return
    jobs_file = work / 'jobs.json'
    jobs_file.write_text(json.dumps(jobs))
    worker = workers.get(engine) or ['node', HERE / 'convert.mjs', engine]
    times = read_batch([*worker, '--batch', jobs_file], env=env)
    if (len(times) != len(jobs) or times[0] is None or
            any(engine != 'minipdf' if t is None else not math.isfinite(t) or t <= 0 for t in times)):
        raise ValueError(f'Invalid warm timings for {engine}')
    for (_, target), ms in zip(jobs, times, strict=True):
        if ms is None:
            continue
        with Path(target).open('rb') as pdf:
            if pdf.read(5) != b'%PDF-':
                raise ValueError(f'Invalid warm conversion output for {engine}')
    for name, ms in zip(cases, times[1:], strict=True):
        if ms is None:
            by_case[name]['warm_error'] = 'Warm conversion failed'
        else:
            by_case[name]['warm_ms'] = ms


def renderer_digest():
    root = HERE.parents[1]
    digest = sha256()
    for path in (HERE / 'convert.mjs', *sorted((root / 'dist').rglob('*.js'))):
        digest.update(str(path.relative_to(root)).encode())
        digest.update(b'\0')
        with path.open('rb') as file:
            digest.update(file_digest(file, 'sha256').digest())
    return digest.hexdigest()


def host_cpu():
    try:
        lines = Path('/proc/cpuinfo').read_text().splitlines()
    except OSError:
        return platform.processor()
    return next((line.split(':', 1)[1].strip() for line in lines if line.startswith('model name')),
                platform.processor())


def adapter_files(engine):
    """The worker scripts, runner sources and lock files that one engine runs through."""
    if engine == 'eigenpal':
        return [HERE / 'convert.mjs', HERE / 'package.json', HERE / 'package-lock.json']
    if engine == 'libreoffice':
        return [HERE / 'libreoffice_warm.py']
    if engine in ('aspose', 'syncfusion', 'minipdf'):
        return sorted(path for path in (HERE / engine).iterdir()
                      if path.suffix in ('.cs', '.csproj') or path.name == 'packages.lock.json')
    if engine in ('dxpdf', 'rdocx'):
        return [HERE / 'rust' / 'runner.py', HERE / 'rust' / 'requirements.txt']
    if engine == 'office2pdf':
        return [HERE / 'rust' / 'install.py', HERE / 'libreoffice_warm.py']
    raise ValueError(f'No adapter files listed for {engine}')


def file_digests(paths):
    root = HERE.parents[1]
    return {path.relative_to(root).as_posix(): sha256(path.read_bytes()).hexdigest() for path in paths}


def reference_digests(fixture):
    if 'pdf_sha256' in fixture:
        return [fixture['pdf_sha256']]
    return [image['sha256'] for image in fixture.get('reference_images', ())]


def cache_key(engine, versions, fixtures, cpu):
    """A digest of everything one engine's cached results depend on."""
    runtime = {name: versions.get(name) for name in (engine, RUNTIMES.get(engine), 'time') if name}
    if engine in ('dxpdf', 'rdocx'):
        runtime['python'] = sys.version
    key = {
        'format': CACHE_FORMAT,
        'engine': engine,
        'version': runtime,
        'adapter': file_digests(adapter_files(engine)),
        'inputs': [[fixture['case'], fixture['docx_sha256'], reference_digests(fixture)] for fixture in fixtures],
        'cpu': cpu,
        'settings': {'sample_size': SAMPLE_SIZE, 'full_timing_ms': FULL_TIMING_MS, 'timeout_s': TIMEOUT_S,
                     'warm_timeout_s': WARM_TIMEOUT_S, 'primed': engine in PRIME},
    }
    return sha256(json.dumps(key, sort_keys=True).encode()).hexdigest()


def link_or_copy(source, target):
    target.unlink(missing_ok=True)
    try:
        # Conversions replace PDFs instead of rewriting them, so a hard link is safe.
        os.link(source, target)
    except OSError:
        shutil.copyfile(source, target)


def is_pdf(path):
    try:
        with path.open('rb') as file:
            return file.read(5) == b'%PDF-'
    except OSError:
        return False


def load_cache(directory, key, engine, fixtures, output):
    """Restore one engine's cached rows and PDFs, or return None when the entry is missing or incomplete."""
    try:
        saved = json.loads((directory / 'results.json').read_text())
        rows = saved['rows']
        if (saved['key'] != key or [row['case'] for row in rows] != [fixture['case'] for fixture in fixtures]
                or any(row['engine'] != engine for row in rows)):
            return None
    except (OSError, KeyError, TypeError, ValueError):
        return None
    if not all(is_pdf(directory / 'pdfs' / f"{row['case']}.pdf") for row in rows if converted(row)):
        return None
    for row in rows:
        target = output / 'cases' / row['case'] / f'{engine}.pdf'
        target.parent.mkdir(parents=True, exist_ok=True)
        target.unlink(missing_ok=True)
        if converted(row):
            link_or_copy(directory / 'pdfs' / f"{row['case']}.pdf", target)
    return saved


def save_cache(directory, key, rows, output):
    """Keep one engine's rows and PDFs, replacing that engine's older cache entries."""
    staging = directory.with_name(directory.name + '.partial')
    shutil.rmtree(staging, ignore_errors=True)
    (staging / 'pdfs').mkdir(parents=True)
    for row in rows:
        if converted(row):
            link_or_copy(output / 'cases' / row['case'] / f"{row['engine']}.pdf",
                         staging / 'pdfs' / f"{row['case']}.pdf")
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
        raise ValueError('Checkpoint was written by an earlier version of convert.py')
    for key, value in settings.items():
        if saved.get(key) != value:
            raise ValueError(f'Checkpoint {key} differs from the current run')
    progress = saved['engines']
    if not set(progress) <= set(engines):
        raise ValueError('Checkpoint has engines that this run does not include')
    for case, digest in settings['inputs'].items():
        verify(OUTPUT / 'cases' / case / 'input.docx', digest)
    for engine, entry in progress.items():
        if [row['case'] for row in entry['rows']] != list(settings['inputs']):
            raise ValueError(f'Checkpoint does not cover the fixtures for {engine}')
        for row in entry['rows']:
            pdf = output / 'cases' / row['case'] / f'{engine}.pdf'
            if row['engine'] != engine or converted(row) and not is_pdf(pdf):
                raise ValueError(f'Invalid cold output {pdf}')
    return progress


def write_log(path, work, data):
    temporary = work / path.name
    temporary.write_text(json.dumps(data, indent=2) + '\n')
    temporary.replace(path)


def run_conversions(args, versions, workers):
    directory = 'minipdf-only' if args.minipdf_only else 'kanera-only' if args.kanera_only else ''
    output = OUTPUT / directory
    output.mkdir(parents=True, exist_ok=True)
    log = output / 'conversion.json'
    checkpoint = output / 'conversion-cold.json'
    if args.minipdf_only:
        with MINIPDF.open('rb') as file:
            build = file_digest(file, 'sha256').hexdigest()
    else:
        build = renderer_digest()
    settings = {
        'format': CACHE_FORMAT,
        'versions': versions,
        'inputs': {fixture['case']: fixture['docx_sha256'] for fixture in BENCHMARK_FIXTURES},
        'renderer_build': build,
    }
    base_engines = () if args.minipdf_only else ('ours',) if args.kanera_only else ('ours', 'eigenpal', 'libreoffice')
    engines = (*base_engines, *workers)
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
        # Invalidate old logs before replacing any cold PDFs.
        checkpoint.unlink(missing_ok=True)
        log.unlink(missing_ok=True)

    cpu = host_cpu()
    sample = sample_cases(BENCHMARK_FIXTURES)
    with TemporaryDirectory(prefix='convert-', dir=output) as temporary:
        work = Path(temporary)
        # Converters keep their temporary files in this run's work directory, not in /tmp.
        env = child_environment(work / 'tmp')
        for engine in engines:
            entry = progress.get(engine)
            cache = key = None
            if engine not in UNCACHED:
                key = cache_key(engine, versions, BENCHMARK_FIXTURES, cpu)
                cache = CACHE / f'{engine}-{key[:16]}'
            if entry is None and cache and not args.no_cache:
                saved = load_cache(cache, key, engine, BENCHMARK_FIXTURES, output)
                if saved is not None:
                    progress[engine] = {'rows': saved['rows'], 'warm': True, 'cached': saved['measured_at']}
                    write_log(checkpoint, work, dict(settings, engines=progress))
                    print(f"{engine}: reusing {cache.relative_to(OUTPUT.parents[1])} from {saved['measured_at']} "
                          '(--no-cache measures it again)', flush=True)
                    continue
            if entry is None:
                rows = convert_engine(engine, BENCHMARK_FIXTURES, sample, work, output, workers, env)
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
            'results': [row for engine in engines for row in progress[engine]['rows']],
            'cached': {engine: progress[engine]['cached'] for engine in engines if progress[engine].get('cached')},
        })


def main():
    parser = ArgumentParser(description='Run cold and warm conversions on the pinned DOCX fixtures.')
    standalone = parser.add_mutually_exclusive_group()
    standalone.add_argument('--kanera-only', action='store_true', help='Run Kanera alone on all benchmark fixtures.')
    standalone.add_argument('--minipdf-only', action='store_true', help='Run MiniPdf alone on all benchmark fixtures.')
    standalone.add_argument('--all', action='store_true', help='Run all comparison engines.')
    parser.add_argument('--minipdf', action='store_true', help='Include MiniPdf in the full comparison.')
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose .NET runner.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO .NET runner.')
    parser.add_argument('--rust', action='store_true', help='Include the Rust engines dxpdf, rdocx and office2pdf.')
    parser.add_argument('--fresh', action='store_true', help='Start cold conversions even if a warm pass is pending.')
    parser.add_argument('--no-cache', action='store_true',
                        help='Measure every engine again instead of reusing cached results.')
    args = parser.parse_args()
    if args.all:
        args.aspose = args.syncfusion = args.minipdf = args.rust = True
    if (args.kanera_only or args.minipdf_only) and (args.aspose or args.syncfusion or args.rust or args.minipdf):
        parser.error('Standalone modes cannot be combined with full-comparison engine flags')
    versions = {'time': version([GNU_TIME, '--version']).splitlines()[0]}
    if not args.minipdf_only:
        versions['node'] = version(['node', '--version'])
        versions['ours'] = json.loads((HERE.parents[1] / 'package.json').read_text())['version']
    if not (args.kanera_only or args.minipdf_only):
        versions['eigenpal'] = json.loads((HERE / 'node_modules/@docx-editor.dev/docx-to-pdf/package.json').read_text())['version']
        versions['libreoffice'] = version(['soffice', '--version'])
    dotnet_workers = {}
    if args.minipdf or args.minipdf_only:
        if not MINIPDF.is_file():
            parser.error('Build the MiniPdf runner under out/benchmark-word/minipdf-bin.')
        dotnet_workers['minipdf'] = ['dotnet', MINIPDF]
    if args.aspose:
        license_path = os.environ.get('ASPOSE_WORDS_LICENSE')
        if not license_path or not Path(license_path).is_file():
            parser.error('Set ASPOSE_WORDS_LICENSE to a local Aspose.Words for .NET license.')
        dotnet_workers['aspose'] = ['dotnet', OUTPUT / 'aspose-bin/Benchmark.dll']
    if args.syncfusion:
        if not os.environ.get('SYNCFUSION_LICENSE_KEY'):
            parser.error('Set SYNCFUSION_LICENSE_KEY to a Syncfusion license key for the pinned version.')
        dotnet_workers['syncfusion'] = ['dotnet', OUTPUT / 'syncfusion-bin/SyncfusionBenchmark.dll']
    for worker in dotnet_workers.values():
        versions.update(json.loads(version([*worker, '--version'])))
    workers = dict(dotnet_workers)
    if args.rust:
        if not OFFICE2PDF.is_file():
            parser.error('Run benchmarks/word/rust/install.py to install office2pdf.')
        runner = [sys.executable, '-I', HERE / 'rust' / 'runner.py']
        try:
            versions.update(json.loads(version([*runner, '--version'])))
        except subprocess.CalledProcessError:
            parser.error('Install benchmarks/word/rust/requirements.txt into the benchmark venv.')
        versions['office2pdf'] = version([OFFICE2PDF, '--version']).removeprefix('office2pdf ')
        workers['dxpdf'] = [*runner, 'dxpdf']
        workers['rdocx'] = [*runner, 'rdocx']
        workers['office2pdf'] = [OFFICE2PDF]

    with benchmark_lock(OUTPUT, stale=[(OUTPUT / directory, 'convert-') for directory in RUN_DIRECTORIES]):
        run_conversions(args, versions, workers)


if __name__ == '__main__':
    main()
