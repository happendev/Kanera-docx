from argparse import ArgumentParser
from datetime import datetime, timezone
from hashlib import file_digest, sha256
import json
import math
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

from fetch import BENCHMARK_FIXTURES, HERE, OUTPUT, benchmark_lock, verify
from libreoffice_warm import batch_times

GNU_TIME = '/usr/bin/time'
OFFICE2PDF = OUTPUT / 'office2pdf-bin' / 'office2pdf'
MINIPDF = OUTPUT / 'minipdf-bin' / 'MiniPdfBenchmark.dll'
# A failed or timed-out conversion is recorded as a zero-score case instead of stopping the run.
TIMEOUT_S = 300


def version(command):
    return subprocess.check_output(command, text=True).strip()


def measure(command, rss_file, timeout=None):
    if timeout:
        # coreutils timeout waits for the converter, so GNU time still sees its peak RSS.
        command = ['timeout', '-k', '5', str(timeout), *command]
    start = perf_counter()
    result = subprocess.run(
        [GNU_TIME, '-f', '%M', '-o', rss_file, *command],
        capture_output=True,
        text=True,
        check=True,
    )
    ms = (perf_counter() - start) * 1000
    rss_kib = int(rss_file.read_text())
    if rss_kib <= 0:
        raise ValueError(f'No peak RSS recorded for {command[0]}')
    return {'cold_ms': ms, 'peak_rss_kib': rss_kib}, result


def office2pdf_command(workspace, output, sources):
    return [OFFICE2PDF, '-j', '1', '--outdir', output, *sources]


def convert_case(fixture, work, node_worker, workers, output, *, kanera_only=False, minipdf_only=False):
    name = fixture['case']
    input_path = OUTPUT / 'cases' / name / 'input.docx'
    verify(input_path, fixture['docx_sha256'])
    directory = output / 'cases' / name
    directory.mkdir(parents=True, exist_ok=True)

    commands = {}
    if not minipdf_only:
        commands['ours'] = ['node', node_worker, 'ours', input_path, directory / 'ours.pdf']
    if not (kanera_only or minipdf_only):
        # Each LibreOffice invocation uses a separate profile and output directory.
        lo = work / name
        lo.mkdir()
        commands['eigenpal'] = ['node', node_worker, 'eigenpal', input_path, directory / 'eigenpal.pdf']
        commands['libreoffice'] = [
            'soffice', '-env:UserInstallation=' + (lo / 'profile').as_uri(),
            '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', lo, input_path,
        ]
    for engine, worker in workers.items():
        if engine == 'office2pdf':
            commands[engine] = [*worker, '-j', '1', '-o', directory / f'{engine}.pdf', input_path]
        else:
            commands[engine] = [*worker, input_path, directory / f'{engine}.pdf']

    for engine, command in commands.items():
        target = directory / f'{engine}.pdf'
        target.unlink(missing_ok=True)
        start = perf_counter()
        try:
            metrics, result = measure(command, work / 'rss.txt', TIMEOUT_S)
            if engine == 'libreoffice':
                (lo / 'input.pdf').replace(target)
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
            yield {'case': name, 'engine': engine, 'cold_ms': (perf_counter() - start) * 1000,
                   'peak_rss_kib': None, 'error': message}
            continue
        details = json.loads(result.stdout) if engine in ('ours', 'eigenpal') else {}
        if result.stderr.strip():
            details['stderr'] = result.stderr.strip()
        print(f"{name} {engine} cold {metrics['cold_ms']:.1f} ms, {metrics['peak_rss_kib'] / 1024:.1f} MiB", flush=True)
        yield {'case': name, 'engine': engine, **metrics, **details}


def read_batch(command):
    result = subprocess.run(command, capture_output=True, text=True, timeout=300)
    if result.returncode:
        raise RuntimeError(f'Warm benchmark worker failed: {result.stderr.strip()}')
    return json.loads(result.stdout)


def resumable_warm(engine, cases, work, by_case, max_restarts=5):
    # A CLI batch can exit partway (LibreOffice occasionally does). Record what it timed, leave the
    # document it stopped on without a warm time, and continue with a fresh batch after it.
    command = (office2pdf_command, 'office2pdf') if engine == 'office2pdf' else ()
    remaining = cases
    for attempt in range(max_restarts + 1):
        attempt_dir = work / f'attempt-{attempt}'
        attempt_dir.mkdir()
        jobs = [(str(OUTPUT / 'cases' / name / 'input.docx'), str(attempt_dir / f'{i}.pdf'))
                for i, name in enumerate([remaining[0], *remaining])]
        times = batch_times(jobs, attempt_dir, *command)
        if any(not math.isfinite(t) or t <= 0 for t in times):
            raise ValueError(f'Invalid warm timings for {engine}')
        for _, target in jobs[:len(times)]:
            with Path(target).open('rb') as pdf:
                if pdf.read(5) != b'%PDF-':
                    raise ValueError(f'Invalid warm conversion output for {engine}')
        timed = remaining[:max(0, len(times) - 1)]
        for name, ms in zip(timed, times[1:]):
            by_case[engine, name]['warm_ms'] = ms
        if len(timed) == len(remaining):
            return
        # The batch stopped on the warm-up or on the next document after those it timed.
        stopped = remaining[len(timed)]
        by_case[engine, stopped]['warm_error'] = 'warm batch stopped on this document'
        remaining = remaining[len(timed) + 1:]
        if not remaining:
            return
    raise RuntimeError(f'{engine} warm batch stopped more than {max_restarts} times')


def measure_warm(results, workspace, workers):
    names = [fixture['case'] for fixture in BENCHMARK_FIXTURES]
    by_case = {(row['engine'], row['case']): row for row in results}
    for engine in dict.fromkeys(row['engine'] for row in results):
        work = workspace / f'warm-{engine}'
        work.mkdir()
        # Only time documents the engine converted in the cold run.
        cases = [name for name in names if 'error' not in by_case[engine, name]]
        if not cases:
            continue
        # Run the first document once without recording it, then time each document in order.
        jobs = [(str(OUTPUT / 'cases' / name / 'input.docx'), str(work / f'{i}.pdf'))
                for i, name in enumerate([cases[0], *cases])]
        print(f'{engine}: warm-up and {len(cases)} warm conversions', flush=True)
        if engine in ('libreoffice', 'office2pdf'):
            resumable_warm(engine, cases, work, by_case)
            continue
        else:
            jobs_file = work / 'jobs.json'
            jobs_file.write_text(json.dumps(jobs))
            worker = workers.get(engine) or ['node', HERE / 'convert.mjs', engine]
            times = read_batch([*worker, '--batch', jobs_file])
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
                by_case[engine, name]['warm_error'] = 'Warm conversion failed'
            else:
                by_case[engine, name]['warm_ms'] = ms


def renderer_digest():
    root = HERE.parents[1]
    digest = sha256()
    for path in (HERE / 'convert.mjs', *sorted((root / 'dist').rglob('*.js'))):
        digest.update(str(path.relative_to(root)).encode())
        digest.update(b'\0')
        with path.open('rb') as file:
            digest.update(file_digest(file, 'sha256').digest())
    return digest.hexdigest()


def load_cold_checkpoint(checkpoint, output, settings, engines):
    saved = json.loads(checkpoint.read_text())
    if not isinstance(saved, dict) or saved.get('versions') != settings['versions']:
        raise ValueError('Checkpoint versions differ from the current run')
    # Earlier checkpoints only recorded versions and results.
    for key in ('inputs', 'renderer_build'):
        if key in saved and saved[key] != settings[key]:
            raise ValueError(f'Checkpoint {key} differs from the current run')
    results = saved['results']
    expected = {(case, engine) for case in settings['inputs'] for engine in engines}
    if len(results) != len(expected) or {(row['case'], row['engine']) for row in results} != expected:
        raise ValueError('Checkpoint does not cover the requested engines and fixtures')
    for case, digest in settings['inputs'].items():
        verify(OUTPUT / 'cases' / case / 'input.docx', digest)
    for row in results:
        if 'error' not in row:
            pdf = output / 'cases' / row['case'] / f"{row['engine']}.pdf"
            with pdf.open('rb') as file:
                if file.read(5) != b'%PDF-':
                    raise ValueError(f'Invalid cold output {pdf}')
    return results


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
        'versions': versions,
        'inputs': {fixture['case']: fixture['docx_sha256'] for fixture in BENCHMARK_FIXTURES},
        'renderer_build': build,
    }
    with TemporaryDirectory(prefix='convert-', dir=output) as temporary:
        work = Path(temporary)
        if not args.fresh and not log.exists() and checkpoint.exists():
            base_engines = () if args.minipdf_only else ('ours',) if args.kanera_only else ('ours', 'eigenpal', 'libreoffice')
            engines = (*base_engines, *workers)
            try:
                results = load_cold_checkpoint(checkpoint, output, settings, engines)
            except (KeyError, OSError, TypeError, ValueError) as error:
                raise SystemExit(f'Cannot resume from {checkpoint}: {error}. '
                                 'Retry with the original settings or use --fresh.') from None
            print(f'Resuming warm conversions from {checkpoint} ({len(results)} cold results)', flush=True)
        else:
            # Invalidate old logs before replacing any cold PDFs.
            checkpoint.unlink(missing_ok=True)
            log.unlink(missing_ok=True)
            results = []
            for fixture in BENCHMARK_FIXTURES:
                results.extend(convert_case(fixture, work, HERE / 'convert.mjs', workers, output,
                                            kanera_only=args.kanera_only, minipdf_only=args.minipdf_only))
            write_log(checkpoint, work, dict(settings, results=results))

        try:
            measure_warm(results, work, workers)
        except Exception:
            print(f'Cold results saved in {checkpoint}. Retry the same command without --fresh to resume.',
                  file=sys.stderr, flush=True)
            raise
        write_log(log, work, {
            'completed_at': datetime.now(timezone.utc).isoformat(),
            'versions': versions,
            'results': results,
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

    with benchmark_lock(OUTPUT):
        run_conversions(args, versions, workers)


if __name__ == '__main__':
    main()
