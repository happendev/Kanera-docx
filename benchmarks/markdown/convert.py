"""Convert the Word benchmark's fixtures to Markdown or plain text with each engine, timing cold and warm runs.

Run fetch.py from benchmarks/word first. The output of the cold run is what score.py scores. --format text
runs the plain-text benchmark (benchmarks/text/convert.py is the same with that option).
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
from fetch import BENCHMARK_FIXTURES, OUTPUT as WORD_OUTPUT, benchmark_lock, child_environment, verify  # noqa: E402
from libreoffice_warm import batch_times  # noqa: E402

CASES = WORD_OUTPUT / 'cases'
DOTNET = HERE / 'dotnet' / 'bin' / 'Release' / 'net10.0' / 'MarkdownBenchmark.dll'
GNU_TIME = '/usr/bin/time'
TIMEOUT_S = 300
ENGINES = ['kanera', 'aspose', 'syncfusion', 'pandoc', 'docling', 'markitdown', 'mammoth', 'libreoffice']
# MarkItDown writes only Markdown. LibreOffice 26.8's plain-text export, with any text filter, writes
# without end on many documents, so the text benchmark leaves it out too.
TEXT_ENGINES = [e for e in ENGINES if e not in ('markitdown', 'libreoffice')]
LICENSED = {'aspose': 'ASPOSE_WORDS_LICENSE', 'syncfusion': 'SYNCFUSION_LICENSE_KEY'}
OUTPUTS = {'markdown': ROOT / 'out' / 'benchmark-markdown', 'text': ROOT / 'out' / 'benchmark-text'}
# One lock for both formats, so a Markdown run and a text run never time each other.
LOCK = OUTPUTS['markdown']


class Format:
    """What the run converts to, set once in main(): Markdown, or plain text with --format text."""
    name = 'markdown'
    extension = 'md'
    flag = '--markdown'
    output = OUTPUTS['markdown']
    results = output / 'md'

    @classmethod
    def use(cls, name):
        cls.name = name
        cls.extension = 'md' if name == 'markdown' else 'txt'
        cls.flag = f'--{name}'
        cls.output = OUTPUTS[name]
        cls.results = cls.output / cls.extension

    @classmethod
    def markdown(cls):
        return cls.name == 'markdown'


def pandoc():
    import pypandoc
    return pypandoc.get_pandoc_path()


def cold_command(engine, source, target, work):
    # The workers and the .NET runner write Markdown or text by the target's extension.
    if engine == 'kanera':
        return ['node', ROOT / 'dist' / 'cli.js', source, Format.flag, target]
    if engine == 'mammoth':
        return ['node', HERE / 'worker.mjs', engine, source, target]
    if engine == 'pandoc':
        return [pandoc(), '-f', 'docx', '-t', 'gfm' if Format.markdown() else 'plain', '-o', target, source]
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


def measure(command, rss_file, env):
    # coreutils timeout waits for the converter, so GNU time still sees its peak RSS.
    command = [GNU_TIME, '-f', '%M', '-o', rss_file, 'timeout', '-k', '5', str(TIMEOUT_S), *map(str, command)]
    start = perf_counter()
    result = subprocess.run(command, capture_output=True, text=True, env=env)
    ms = (perf_counter() - start) * 1000
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout).strip().splitlines()[-1:] or f'exit {result.returncode}')
    return ms, int(rss_file.read_text().split()[-1])


def cold(engine, names, work, env):
    rows = {}
    target_dir = Format.results / engine
    shutil.rmtree(target_dir, ignore_errors=True)
    target_dir.mkdir(parents=True)
    for name in names:
        source = CASES / name / 'input.docx'
        target = target_dir / f'{name}.{Format.extension}'
        case_work = work / name
        case_work.mkdir()
        try:
            ms, rss = measure(cold_command(engine, source, target, case_work), work / 'rss.txt', env)
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


def warm(engine, rows, work, env):
    # Only documents the engine converted cold; the first is converted once untimed as a warm-up.
    cases = [name for name, row in rows.items() if 'error' not in row]
    if not cases:
        return
    remaining = cases
    attempt = 0
    while remaining:
        attempt_dir = work / f'warm-{attempt}'
        attempt_dir.mkdir()
        jobs = [(str(CASES / name / 'input.docx'), str(attempt_dir / f'{i}.{Format.extension}'))
                for i, name in enumerate([remaining[0], *remaining])]
        if engine == 'libreoffice':
            # One CLI batch; each time is the interval between successive outputs.
            times = batch_times(jobs, attempt_dir, libreoffice_md, extension='md', env=env)
        else:
            jobs_file = attempt_dir / 'jobs.json'
            jobs_file.write_text(json.dumps(jobs))
            result = subprocess.run([*map(str, warm_worker(engine)), '--batch', str(jobs_file)],
                                    capture_output=True, text=True, timeout=TIMEOUT_S * len(jobs), env=env)
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
            # Plain text is mammoth's own raw text; Markdown goes through turndown.
            used = ('mammoth', 'turndown', 'turndown-plugin-gfm') if Format.markdown() else ('mammoth',)
            found[engine] = ' + '.join(f"{m} {json.loads((mods / m / 'package.json').read_text())['version']}"
                                       for m in used)
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


def main(argv=None):
    parser = ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--format', choices=('markdown', 'text'), default='markdown',
                        help='what to convert to (default: %(default)s)')
    parser.add_argument('--engines', help='comma-separated engines to run (default: every engine for the format); '
                                          'results for other engines are kept')
    parser.add_argument('--no-warm', action='store_true', help='skip the warm runs')
    parser.add_argument('--limit', type=int, help='convert only the first N documents, for a quick check; '
                                                  'score with the same --limit')
    args = parser.parse_args(argv)
    Format.use(args.format)
    known = ENGINES if Format.markdown() else TEXT_ENGINES
    engines = [e for e in (args.engines or ','.join(known)).split(',') if e]
    unknown = set(engines) - set(known)
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
    if args.limit is not None:
        names = names[:args.limit]

    # One run at a time, Markdown or text: it holds the lock, and work-* directories found in either
    # output directory were left by a stopped run.
    with benchmark_lock(LOCK, stale=[(output, f'work-{engine}-') for output in OUTPUTS.values() for engine in ENGINES]):
        run(engines, names, args.no_warm)


def run(engines, names, no_warm):
    Format.output.mkdir(parents=True, exist_ok=True)
    log = Format.output / 'conversion.json'
    data = json.loads(log.read_text()) if log.exists() else {'versions': {}, 'results': []}
    data['results'] = [row for row in data['results'] if row['engine'] not in engines]
    data['versions'].update(versions(engines))
    for engine in engines:
        # Work under the output directory: a sandboxed LibreOffice (Flatpak) has a /tmp of its own.
        with TemporaryDirectory(prefix=f'work-{engine}-', dir=Format.output) as tmp:
            # Converters keep their temporary files in the work directory, not in /tmp.
            env = child_environment(Path(tmp) / 'tmp')
            start = perf_counter()
            rows = cold(engine, names, Path(tmp), env)
            failed = sum('error' in row for row in rows.values())
            print(f'{engine}: {len(names) - failed}/{len(names)} cold conversions in {perf_counter() - start:.0f} s', flush=True)
            if not no_warm:
                warm(engine, rows, Path(tmp), env)
        data['results'] += rows.values()
        data['completed_at'] = datetime.now(timezone.utc).isoformat()
        log.write_text(json.dumps(data, indent=1, ensure_ascii=False))


if __name__ == '__main__':
    main()
