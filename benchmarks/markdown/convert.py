"""Convert the Word benchmark's fixtures to Markdown with each engine, timing cold and warm runs.

Run fetch.py from benchmarks/word first. The Markdown of the cold run is what score.py scores.
Aspose.Words and Syncfusion DocIO run when ASPOSE_WORDS_LICENSE and SYNCFUSION_LICENSE_KEY are set.
"""
from argparse import ArgumentParser
from datetime import datetime, timezone
from importlib.metadata import version as package_version
import json
import math
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
from time import perf_counter

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT / 'benchmarks' / 'word'))
from fetch import BENCHMARK_FIXTURES, OUTPUT as WORD_OUTPUT, verify  # noqa: E402
from libreoffice_warm import batch_times  # noqa: E402

OUTPUT = ROOT / 'out' / 'benchmark-markdown'
MARKDOWN = OUTPUT / 'md'
CASES = WORD_OUTPUT / 'cases'
DOTNET = HERE / 'dotnet' / 'bin' / 'Release' / 'net10.0' / 'MarkdownBenchmark.dll'
GNU_TIME = '/usr/bin/time'
TIMEOUT_S = 300
ENGINES = ['kanera', 'aspose', 'syncfusion', 'pandoc', 'docling', 'markitdown', 'mammoth', 'libreoffice']
LICENSED = {'aspose': 'ASPOSE_WORDS_LICENSE', 'syncfusion': 'SYNCFUSION_LICENSE_KEY'}


def pandoc():
    import pypandoc
    return pypandoc.get_pandoc_path()


def cold_command(engine, source, target, work):
    if engine == 'kanera':
        return ['node', ROOT / 'dist' / 'cli.js', source, '--markdown', target]
    if engine == 'mammoth':
        return ['node', HERE / 'worker.mjs', engine, source, target]
    if engine == 'pandoc':
        return [pandoc(), '-f', 'docx', '-t', 'gfm', '-o', target, source]
    if engine in ('docling', 'markitdown'):
        return [sys.executable, HERE / 'worker.py', engine, source, target]
    if engine in ('aspose', 'syncfusion'):
        return ['dotnet', DOTNET, engine, source, target]
    # Each LibreOffice invocation uses a separate profile and output directory.
    return ['soffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless',
            '--convert-to', 'md', '--outdir', work, source]


def warm_worker(engine):
    if engine in ('kanera', 'mammoth'):
        return ['node', HERE / 'worker.mjs', engine]
    if engine in ('docling', 'markitdown', 'pandoc'):
        return [sys.executable, HERE / 'worker.py', engine]
    return ['dotnet', DOTNET, engine]


def libreoffice_md(workspace, output, sources):
    return ['soffice', '-env:UserInstallation=' + (workspace / 'profile').as_uri(), '--headless',
            '--convert-to', 'md', '--outdir', output, *sources]


def measure(command, rss_file):
    # coreutils timeout waits for the converter, so GNU time still sees its peak RSS.
    command = [GNU_TIME, '-f', '%M', '-o', rss_file, 'timeout', '-k', '5', str(TIMEOUT_S), *map(str, command)]
    start = perf_counter()
    result = subprocess.run(command, capture_output=True, text=True)
    ms = (perf_counter() - start) * 1000
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout).strip().splitlines()[-1:] or f'exit {result.returncode}')
    return ms, int(rss_file.read_text().split()[-1])


def cold(engine, names, work):
    rows = {}
    target_dir = MARKDOWN / engine
    shutil.rmtree(target_dir, ignore_errors=True)
    target_dir.mkdir(parents=True)
    for name in names:
        source = CASES / name / 'input.docx'
        target = target_dir / f'{name}.md'
        case_work = work / name
        case_work.mkdir()
        try:
            ms, rss = measure(cold_command(engine, source, target, case_work), work / 'rss.txt')
            if engine == 'libreoffice':
                (case_work / 'input.md').replace(target)
                rss = None  # GNU time sees the Flatpak launcher, not LibreOffice.
            if not target.exists():
                raise RuntimeError('no output')
            rows[name] = {'case': name, 'engine': engine, 'cold_ms': ms, 'peak_rss_kib': rss}
        except Exception as error:
            target.unlink(missing_ok=True)
            rows[name] = {'case': name, 'engine': engine, 'error': str(error)[:300]}
        shutil.rmtree(case_work, ignore_errors=True)
    return rows


def warm(engine, rows, work):
    # Only documents the engine converted cold; the first is converted once untimed as a warm-up.
    cases = [name for name, row in rows.items() if 'error' not in row]
    if not cases:
        return
    remaining = cases
    attempt = 0
    while remaining:
        attempt_dir = work / f'warm-{attempt}'
        attempt_dir.mkdir()
        jobs = [(str(CASES / name / 'input.docx'), str(attempt_dir / f'{i}.md'))
                for i, name in enumerate([remaining[0], *remaining])]
        if engine == 'libreoffice':
            # One CLI batch; each time is the interval between successive outputs.
            times = batch_times(jobs, attempt_dir, libreoffice_md, extension='md')
        else:
            jobs_file = attempt_dir / 'jobs.json'
            jobs_file.write_text(json.dumps(jobs))
            result = subprocess.run([*map(str, warm_worker(engine)), '--batch', str(jobs_file)],
                                    capture_output=True, text=True, timeout=TIMEOUT_S * len(jobs))
            if result.returncode:
                raise RuntimeError(f'{engine} warm worker failed: {result.stderr.strip()[-500:]}')
            times = json.loads(result.stdout.strip().splitlines()[-1])
        timed = remaining[:max(0, len(times) - 1)]
        for name, ms in zip(timed, times[1:]):
            if ms is None or not math.isfinite(ms):
                rows[name]['warm_error'] = 'warm conversion failed'
            else:
                rows[name]['warm_ms'] = ms
        if len(timed) == len(remaining):
            return
        # A LibreOffice batch can stop partway: skip the document it stopped on and resume.
        rows[remaining[len(timed)]]['warm_error'] = 'warm batch stopped on this document'
        remaining = remaining[len(timed) + 1:]
        attempt += 1
        if attempt > 5:
            raise RuntimeError(f'{engine} warm batch stopped more than five times')


def versions(engines):
    found = {'node': subprocess.check_output(['node', '--version'], text=True).strip(),
             'python': platform.python_version()}
    package = json.loads((ROOT / 'package.json').read_text())
    for engine in engines:
        if engine == 'kanera':
            found[engine] = f"{package['name']} {package['version']}"
        elif engine == 'mammoth':
            mods = HERE / 'node_modules'
            found[engine] = ' + '.join(f"{m} {json.loads((mods / m / 'package.json').read_text())['version']}"
                                       for m in ('mammoth', 'turndown', 'turndown-plugin-gfm'))
        elif engine == 'pandoc':
            found[engine] = subprocess.check_output([pandoc(), '--version'], text=True).splitlines()[0]
        elif engine in ('docling', 'markitdown'):
            found[engine] = f'{engine} {package_version(engine)}'
        elif engine in ('aspose', 'syncfusion'):
            info = json.loads(subprocess.check_output(['dotnet', str(DOTNET), '--version'], text=True))
            found[engine] = f"{info[engine]} ({info['dotnet']})"
        elif engine == 'libreoffice':
            found[engine] = subprocess.check_output(['soffice', '--version'], text=True).strip()
    return found


def main():
    parser = ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--engines', default=','.join(ENGINES),
                        help='comma-separated engines to run; results for other engines are kept')
    parser.add_argument('--no-warm', action='store_true', help='skip the warm runs')
    args = parser.parse_args()
    engines = [e for e in args.engines.split(',') if e]
    unknown = set(engines) - set(ENGINES)
    if unknown:
        parser.error(f'unknown engines: {", ".join(sorted(unknown))}')
    for engine, variable in LICENSED.items():
        if engine in engines and not os.environ.get(variable):
            print(f'Skipping {engine}: set {variable} to run it.', flush=True)
            engines.remove(engine)

    names = []
    for fixture in BENCHMARK_FIXTURES:
        verify(CASES / fixture['case'] / 'input.docx', fixture['docx_sha256'])
        names.append(fixture['case'])

    OUTPUT.mkdir(parents=True, exist_ok=True)
    log = OUTPUT / 'conversion.json'
    data = json.loads(log.read_text()) if log.exists() else {'versions': {}, 'results': []}
    data['results'] = [row for row in data['results'] if row['engine'] not in engines]
    data['versions'].update(versions(engines))
    for engine in engines:
        # Work under the output directory: a sandboxed LibreOffice (Flatpak) has a /tmp of its own.
        with TemporaryDirectory(prefix=f'work-{engine}-', dir=OUTPUT) as tmp:
            start = perf_counter()
            rows = cold(engine, names, Path(tmp))
            failed = sum('error' in row for row in rows.values())
            print(f'{engine}: {len(names) - failed}/{len(names)} cold conversions in {perf_counter() - start:.0f} s', flush=True)
            if not args.no_warm:
                warm(engine, rows, Path(tmp))
        data['results'] += rows.values()
        data['completed_at'] = datetime.now(timezone.utc).isoformat()
        log.write_text(json.dumps(data, indent=1, ensure_ascii=False))


if __name__ == '__main__':
    main()
