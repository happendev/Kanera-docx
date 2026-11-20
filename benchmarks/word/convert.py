from argparse import ArgumentParser
from datetime import datetime, timezone
import json
import math
import os
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

from fetch import BENCHMARK_FIXTURES, HERE, OUTPUT, verify
from libreoffice_warm import batch_times

GNU_TIME = '/usr/bin/time'
OFFICE2PDF = OUTPUT / 'office2pdf-bin' / 'office2pdf'
# Engines whose failures are recorded as zero-score cases instead of stopping the run.
TOLERANT = ('dxpdf', 'rdocx', 'office2pdf')
TOLERANT_TIMEOUT_S = 300


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


def convert_case(fixture, work, node_worker, workers, *, kanera_only=False):
    name = fixture['case']
    input_path = OUTPUT / 'cases' / name / 'input.docx'
    verify(input_path, fixture['docx_sha256'])
    output = OUTPUT / 'kanera-only' if kanera_only else OUTPUT
    directory = output / 'cases' / name
    directory.mkdir(parents=True, exist_ok=True)

    commands = {'ours': ['node', node_worker, 'ours', input_path, directory / 'ours.pdf']}
    if not kanera_only:
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
        tolerant = engine in TOLERANT
        start = perf_counter()
        try:
            metrics, result = measure(command, work / 'rss.txt', TOLERANT_TIMEOUT_S if tolerant else None)
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
            if not tolerant:
                raise RuntimeError(f'{name} ({engine}) failed: {message}') from error
            target.unlink(missing_ok=True)
            print(f'{name} {engine} FAILED: {message.splitlines()[-1] if message else error}', flush=True)
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
        if engine == 'libreoffice':
            times = batch_times(jobs, work)
        elif engine == 'office2pdf':
            times = batch_times(jobs, work, office2pdf_command, 'office2pdf')
        else:
            jobs_file = work / 'jobs.json'
            jobs_file.write_text(json.dumps(jobs))
            worker = workers.get(engine) or ['node', HERE / 'convert.mjs', engine]
            times = read_batch([*worker, '--batch', jobs_file])
        if len(times) != len(jobs) or any(not math.isfinite(t) or t <= 0 for t in times):
            raise ValueError(f'Invalid warm timings for {engine}')
        for _, target in jobs:
            with Path(target).open('rb') as pdf:
                if pdf.read(5) != b'%PDF-':
                    raise ValueError(f'Invalid warm conversion output for {engine}')
        for name, ms in zip(cases, times[1:], strict=True):
            by_case[engine, name]['warm_ms'] = ms


def main():
    parser = ArgumentParser(description='Run cold and warm conversions on the pinned DOCX fixtures.')
    parser.add_argument('--kanera-only', action='store_true', help='Run Kanera alone on all benchmark fixtures.')
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose .NET runner.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO .NET runner.')
    parser.add_argument('--rust', action='store_true', help='Include the Rust engines dxpdf, rdocx and office2pdf.')
    args = parser.parse_args()
    if args.kanera_only and (args.aspose or args.syncfusion or args.rust):
        parser.error('--kanera-only cannot be combined with --aspose, --syncfusion or --rust')
    versions = {
        'node': version(['node', '--version']),
        'ours': json.loads((HERE.parents[1] / 'package.json').read_text())['version'],
        'time': version([GNU_TIME, '--version']).splitlines()[0],
    }
    if not args.kanera_only:
        versions['eigenpal'] = json.loads((HERE / 'node_modules/@docx-editor.dev/docx-to-pdf/package.json').read_text())['version']
        versions['libreoffice'] = version(['soffice', '--version'])
    dotnet_workers = {}
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

    output = OUTPUT / 'kanera-only' if args.kanera_only else OUTPUT
    output.mkdir(parents=True, exist_ok=True)
    log = output / 'conversion.json'
    log.unlink(missing_ok=True)
    results = []
    with TemporaryDirectory(prefix='convert-', dir=output) as temporary:
        work = Path(temporary)
        for fixture in BENCHMARK_FIXTURES:
            results.extend(convert_case(fixture, work, HERE / 'convert.mjs', workers,
                                        kanera_only=args.kanera_only))
        measure_warm(results, work, workers)
    completed_at = datetime.now(timezone.utc).isoformat()
    log.write_text(json.dumps({
        'completed_at': completed_at,
        'versions': versions,
        'results': results,
    }, indent=2) + '\n')


if __name__ == '__main__':
    main()
