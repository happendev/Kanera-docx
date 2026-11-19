from argparse import ArgumentParser
import json
import math
import os
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory
from time import perf_counter

from fetch import BENCHMARK_FIXTURES, HERE, OUTPUT, verify
from libreoffice_warm import batch_times

GNU_TIME = '/usr/bin/time'


def version(command):
    return subprocess.check_output(command, text=True).strip()


def measure(command, rss_file):
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


def convert_case(fixture, work, node_worker, dotnet_workers):
    name = fixture['case']
    directory = OUTPUT / 'cases' / name
    input_path = directory / 'input.docx'
    verify(input_path, fixture['docx_sha256'])

    # Give every LibreOffice invocation its own directory and profile so it cannot
    # forward the work to a running desktop process or reuse the previous job.
    lo = work / name
    lo.mkdir()
    commands = {
        'ours': ['node', node_worker, 'ours', input_path, directory / 'ours.pdf'],
        'eigenpal': ['node', node_worker, 'eigenpal', input_path, directory / 'eigenpal.pdf'],
        'libreoffice': [
            'soffice', '-env:UserInstallation=' + (lo / 'profile').as_uri(),
            '--headless', '--convert-to', 'pdf:writer_pdf_Export', '--outdir', lo, input_path,
        ],
    }
    for engine, worker in dotnet_workers.items():
        commands[engine] = [*worker, input_path, directory / f'{engine}.pdf']

    for engine, command in commands.items():
        target = directory / f'{engine}.pdf'
        target.unlink(missing_ok=True)
        try:
            metrics, result = measure(command, work / 'rss.txt')
        except subprocess.CalledProcessError as error:
            raise RuntimeError(f'{name} ({engine}) failed: {error.stderr.strip()}') from error
        if engine == 'libreoffice':
            (lo / 'input.pdf').replace(target)
        if not target.is_file():
            raise RuntimeError(f'{name} ({engine}) did not produce a PDF')
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


def measure_warm(results, workspace, dotnet_workers):
    names = [fixture['case'] for fixture in BENCHMARK_FIXTURES]
    by_case = {(row['engine'], row['case']): row for row in results}
    for engine in dict.fromkeys(row['engine'] for row in results):
        work = workspace / f'warm-{engine}'
        work.mkdir()
        # Run the first document once without recording it, then time each document in order.
        jobs = [(str(OUTPUT / 'cases' / name / 'input.docx'), str(work / f'{i}.pdf'))
                for i, name in enumerate([names[0], *names])]
        print(f'{engine}: warm-up and {len(names)} warm conversions', flush=True)
        if engine == 'libreoffice':
            times = batch_times(jobs, work)
        else:
            jobs_file = work / 'jobs.json'
            jobs_file.write_text(json.dumps(jobs))
            worker = dotnet_workers.get(engine) or ['node', HERE / 'convert.mjs', engine]
            times = read_batch([*worker, '--batch', jobs_file])
        if len(times) != len(jobs) or any(not math.isfinite(t) or t <= 0 for t in times):
            raise ValueError(f'Invalid warm timings for {engine}')
        for _, target in jobs:
            with Path(target).open('rb') as pdf:
                if pdf.read(5) != b'%PDF-':
                    raise ValueError(f'Invalid warm conversion output for {engine}')
        for name, ms in zip(names, times[1:], strict=True):
            by_case[engine, name]['warm_ms'] = ms


def main():
    parser = ArgumentParser(description='Run cold and warm conversions on the pinned DOCX fixtures.')
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose .NET runner.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO .NET runner.')
    args = parser.parse_args()
    versions = {
        'node': version(['node', '--version']),
        'ours': json.loads((HERE.parents[1] / 'package.json').read_text())['version'],
        'eigenpal': json.loads((HERE / 'node_modules/@docx-editor.dev/docx-to-pdf/package.json').read_text())['version'],
        'libreoffice': version(['soffice', '--version']),
        'time': version([GNU_TIME, '--version']).splitlines()[0],
    }
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

    log = OUTPUT / 'conversion.json'
    log.unlink(missing_ok=True)
    results = []
    with TemporaryDirectory(prefix='convert-', dir=OUTPUT) as temporary:
        work = Path(temporary)
        for fixture in BENCHMARK_FIXTURES:
            results.extend(convert_case(fixture, work, HERE / 'convert.mjs', dotnet_workers))
        measure_warm(results, work, dotnet_workers)
    log.write_text(json.dumps({'versions': versions, 'results': results}, indent=2) + '\n')


if __name__ == '__main__':
    main()
