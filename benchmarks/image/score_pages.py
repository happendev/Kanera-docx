"""Score rendered page images against Word's pages with the Word benchmark's ink scorer.

Each PNG is resampled from its rendering resolution to the scorer's 96 dpi, as the classic suite's
150 dpi Word page images are, then scored with benchmarks/word/ink.py. See README.md.
"""
from argparse import ArgumentParser, BooleanOptionalAction
from concurrent.futures import ProcessPoolExecutor, as_completed
import csv
from datetime import datetime, timezone
from html import escape
import json
from os import cpu_count
from pathlib import Path
from platform import platform, processor
import shutil
import statistics
import subprocess
from tempfile import TemporaryDirectory

from PIL import Image

from pages import DPI, ENGINES, HERE, OPTIONAL, OUTPUT, ROOT, WORD, engine_dpi, page_files
from fetch import BENCHMARK_FIXTURES, MANIFEST, UPSTREAM, benchmark_lock, benchmark_reference, child_environment, verify_reference
from ink import DPI as SCORE_DPI, INK_THRESHOLD, INK_TOLERANCE_PX, page_scores, render_reference
from score_cache import FORMAT, ScoreCache, code_digests, digest, file_sha256
from readme import markdown_table, sampled_note
import score as word_score
from image_readme import COLUMNS, OTHER_DPI, other_dpi_note, update_readmes

# The gallery images label each converter with this benchmark's names.
word_score.ENGINES.update(ENGINES)
SCORE_CODE = (HERE / 'score_pages.py', HERE / 'pages.py', WORD / 'ink.py', WORD / 'score_cache.py')


def cache_context(poppler, env):
    """The Word benchmark's reference-page context, and a score context for page images."""
    context = word_score.cache_context(poppler, env)
    reference = context['reference']
    return {
        'reference': reference,
        'score': {'format': FORMAT, 'kind': 'page-image score', 'tools': context['score']['tools'],
                  'settings': {'dpi': SCORE_DPI, 'ink_threshold': INK_THRESHOLD, 'tolerance_px': INK_TOLERANCE_PX},
                  'code': code_digests(SCORE_CODE, ROOT)},
    }


def resample(page, dpi, target):
    """Write page as an opaque RGB PNG at the scorer's resolution."""
    with Image.open(page) as image:
        image.load()
        if image.mode in ('RGBA', 'LA', 'PA') or 'transparency' in image.info:
            # Transparent areas are paper, not ink.
            rgba = image.convert('RGBA')
            image = Image.alpha_composite(Image.new('RGBA', rgba.size, 'white'), rgba)
        image = image.convert('RGB')
        if dpi != SCORE_DPI:
            scale = SCORE_DPI / dpi
            size = (max(1, round(image.width * scale)), max(1, round(image.height * scale)))
            image = image.resize(size, Image.Resampling.LANCZOS)
        image.save(target, compress_level=1)
    return target


def resample_pages(pages, dpi, directory):
    """The resampled pages, or the reason they cannot be read."""
    directory.mkdir(parents=True, exist_ok=True)
    try:
        return [resample(page, dpi, directory / page.name) for page in pages], None
    except (OSError, SyntaxError, ValueError, Image.DecompressionBombError) as error:
        return [], f'unreadable PNG: {error}'


def score_case(fixture, engines, metrics, dpis, visuals, *, output, cache, context, work, env):
    name = fixture['case']
    reference_dir, reference_fixture = benchmark_reference(fixture)
    verify_reference(reference_dir, reference_fixture)
    source = word_score.reference_source(reference_fixture)
    reference = cache.reference_pages(
        digest({**context['reference'], 'reference': source}),
        lambda target: render_reference(reference_dir, reference_fixture, target, env=env))
    if not reference:
        raise ValueError(f'No reference pages for {name}')

    # A sampled engine skips documents that nothing scores; they are left out of everything.
    engines = [engine for engine in engines if not metrics[name, engine].get('skipped')]
    scratch = work / 'cases' / name
    rasters = {'reference': reference}
    results = {}
    try:
        for engine in engines:
            rasters[engine] = []
            if 'error' in metrics[name, engine]:
                # A failed conversion has no pages, so every reference page scores zero.
                continue
            pages = page_files(output / 'cases' / name / engine)
            key = digest({**context['score'], 'dpi': dpis[engine], 'reference': source,
                          'candidate': [file_sha256(page) for page in pages]})
            result = cache.get(key)
            if result is None or visuals is not None:
                resampled, unreadable = resample_pages(pages, dpis[engine], scratch / engine)
                rasters[engine] = resampled
                if result is None:
                    result = {'unreadable': unreadable, 'pages': len(pages)}
                    if unreadable is None:
                        result['page_scores'] = page_scores(reference, resampled)
                        cache.put(key, result)
            if result['unreadable']:
                # Page images the scorer cannot read count as a failed conversion and score zero.
                print(f"Unreadable pages for {name} ({engine}): {result['unreadable']}; scoring them as failed",
                      flush=True)
                rasters[engine] = []
            results[engine] = result

        rows = []
        for engine in engines:
            row = metrics[name, engine]
            result = results.get(engine, {})
            invalid = result.get('unreadable')
            failed = 'error' in row or bool(invalid)
            scores = page_scores(reference, []) if failed else result['page_scores']
            rows.append({
                'case': name,
                'corpus': fixture['corpus'],
                'suite': fixture['suite'],
                'reference_application': fixture['reference_application'],
                'reference_format': 'pdf' if 'pdf' in reference_fixture else 'png',
                'engine': engine,
                'dpi': dpis[engine],
                'failed': failed,
                'invalid_png': invalid,
                'untimed': bool(row.get('untimed')),
                'cold_ms': round(row['cold_ms'], 2),
                'warm_ms': None if failed else row.get('warm_ms'),
                'peak_rss_kib': row['peak_rss_kib'],
                'reference_pages': len(reference),
                'pages': 0 if failed else result['pages'],
                'ink_f1': round(statistics.mean(scores) * 100, 2),
                'page_ink_f1': [round(score * 100, 2) for score in scores],
                'png_bytes': 0 if failed else row['png_bytes'],
            })

        if visuals is not None:
            # The Word benchmark's side-by-side images, captioned with each converter's results.
            captions = [dict(row, unreadable_pdf=row['invalid_png']) for row in rows]
            for page in range(max(map(len, rasters.values()))):
                word_score.comparison(fixture, rasters, captions, page, visuals)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    return rows


def summary_table(rows, engines, count):
    """The Word benchmark's summary columns, with the total PNG size in place of the PDF size."""
    yield COLUMNS
    for engine in engines:
        scores = [row for row in rows if row['engine'] == engine]
        total = len(scores)
        converted = [row for row in scores if not row['failed']]
        timed = [row for row in converted if not row['untimed']]
        mark = ' †' if total < count or len(timed) < len(converted) else ''
        timing = ('–',) * 6
        if timed:
            cold = statistics.median(row['cold_ms'] for row in timed)
            warm_times = [row['warm_ms'] for row in timed if row['warm_ms'] is not None]
            warm_columns = ('–',) * 3
            if warm_times:
                warm = statistics.median(warm_times)
                p95 = statistics.quantiles(warm_times, n=100, method='inclusive')[94] if len(warm_times) > 1 else warm
                warm_columns = (f'{warm:.1f} ms', f'{p95:.1f} ms', f'{1000 / statistics.mean(warm_times):.1f}')
            rss = [row['peak_rss_kib'] / 1024 for row in timed]
            timing = tuple(value if value == '–' else value + mark for value in (
                f'{cold:.1f} ms', *warm_columns, f'{statistics.median(rss):.1f} MiB', f'{max(rss):.1f} MiB'))
        matches = sum(row['pages'] == row['reference_pages'] for row in scores)
        ink = statistics.mean(row['ink_f1'] for row in scores) if scores else 0
        size = f"{sum(row['png_bytes'] for row in scores) / 1024**2:.2f} MiB"
        if engine_dpi(engine) != DPI:
            size += f' {OTHER_DPI}'
        yield (ENGINES[engine], f'{ink:.1f}%', f'{matches}/{total}', f'{total - len(converted)}/{total}',
               *timing, size)


def report(table, versions, cases, pages, *, include_gallery, run_date):
    lines = [
        '# DOCX-to-page-image benchmark',
        '',
        f'{len(cases)} DOCX files rendered to PNG at {DPI} dpi and compared with {pages} reference pages at '
        f'{SCORE_DPI} dpi.',
        f'Run date {run_date} (UTC).',
        '',
        f"[Hugging Face corpus]({MANIFEST['source']}/tree/{MANIFEST['revision']}) · "
        f"[Vendored fixtures]({UPSTREAM['source']}/commits/{UPSTREAM['revision']})"
        + (' · [Screenshot gallery](gallery.html)' if include_gallery else ''),
        '',
        markdown_table(table[0], table[1:]),
        '',
        'Cold time includes process startup, reading the DOCX and writing every PNG. Warm time reuses one '
        'process per engine after an unmeasured conversion. Peak RSS, accuracy and PNG sizes come from the cold '
        'runs. Times and memory cover successful conversions. A failed conversion scores zero ink overlap.',
    ]
    if any(cell.endswith('†') for row in table[1:] for cell in row):
        lines.extend(['', sampled_note()])
    if any(cell.endswith(OTHER_DPI) for row in table[1:] for cell in row):
        lines.extend(['', other_dpi_note()])
    lines.extend(['', '## Environment', '', f'- Host ({platform()}, {processor() or "processor unspecified"})'])
    lines.extend(f'- {key} ({value})' for key, value in versions.items())
    lines.extend(['', '[Raw scores](results.csv)', ''])
    return '\n'.join(lines)


def gallery(table, cases, output):
    yield ('<!doctype html>\n<html lang="en"><meta charset="utf-8">'
           '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
           '<title>DOCX-to-page-image benchmark</title>\n<style>body { font: 16px system-ui, sans-serif; '
           'margin: 1.5rem; } th, td { padding: .4rem .65rem; text-align: left; } img { display: block; '
           'height: auto; width: 100%; max-width: 2560px; } figure { margin: 0 0 1.5rem; }</style>\n'
           '<h1>DOCX-to-page-image benchmark</h1><p><a href="report.md">Report</a> · '
           '<a href="results.csv">Raw scores</a></p>\n<table><thead><tr>')
    yield ''.join(f'<th scope="col">{escape(cell)}</th>' for cell in table[0]) + '</tr></thead><tbody>'
    for row in table[1:]:
        yield '<tr>' + ''.join(f'<td>{escape(cell)}</td>' for cell in row) + '</tr>'
    yield '</tbody></table>'
    for name in cases:
        yield f'<section id="{escape(name)}"><h2>{escape(name)}</h2>'
        for page, image in enumerate(sorted((output / 'visuals').glob(f'{name}-*.png')), 1):
            path = escape(image.relative_to(output).as_posix())
            yield (f'<figure><figcaption>Page {page}</figcaption><a href="{path}"><img src="{path}" '
                   f'loading="lazy" alt="Side-by-side comparison of {escape(name)}, page {page}"></a></figure>')
        yield '</section>'
    yield '</html>'


def score_outputs(args, root, work):
    output = root / ('kanera-only' if args.kanera_only else '')
    engines = ('ours',) if args.kanera_only else tuple(
        engine for engine in ENGINES if engine not in OPTIONAL or getattr(args, OPTIONAL[engine]))
    data = json.loads((output / 'conversion.json').read_text())
    run_date = datetime.fromisoformat(data['completed_at']).astimezone(timezone.utc).date().isoformat()
    versions = data['versions']
    cases = data['cases']
    dpis = data['dpi']
    metrics = {(row['case'], row['engine']): row for row in data['results']}
    expected = {(case, engine) for case in cases for engine in engines}
    if metrics.keys() != expected or len(metrics) != len(data['results']):
        raise ValueError('Conversion logs do not cover the requested engines; pass the same engine flags as convert_pages.py')
    fixtures = {fixture['case']: fixture for fixture in BENCHMARK_FIXTURES}
    env = child_environment(work / 'tmp')
    poppler = subprocess.run(['pdftoppm', '-v'], check=True, capture_output=True, text=True, env=env)
    versions['poppler'] = (poppler.stderr or poppler.stdout).splitlines()[0]
    cache = ScoreCache(root / 'score-cache', work, read=not args.no_cache)
    context = cache_context(versions['poppler'], env)

    include_gallery = not args.kanera_only if args.gallery is None else args.gallery
    visuals = output / 'visuals'
    for old in visuals.glob('*.png'):
        old.unlink()
    if include_gallery:
        visuals.mkdir(exist_ok=True)
    else:
        (output / 'gallery.html').unlink(missing_ok=True)
        visuals = None
    scored = {}
    print(f'Scoring {len(cases)} documents with {args.jobs} workers', flush=True)
    with ProcessPoolExecutor(max_workers=args.jobs) as pool:
        futures = {}
        for name in cases:
            case_metrics = {(name, engine): metrics[name, engine] for engine in engines}
            future = pool.submit(score_case, fixtures[name], engines, case_metrics, dpis, visuals, output=output,
                                 cache=cache, context=context, work=work, env=env)
            futures[future] = name
        try:
            for future in as_completed(futures):
                scored[futures[future]] = future.result()
                print(f'Scored {futures[future]}', flush=True)
        except BaseException:
            pool.shutdown(cancel_futures=True)
            raise
    rows = [row for name in cases for row in scored[name]]
    with (output / 'results.csv').open('w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=rows[0])
        writer.writeheader()
        writer.writerows(rows)
    table = list(summary_table(rows, engines, len(cases)))
    pages = sum({row['case']: row['reference_pages'] for row in rows}.values())
    summary = report(table, versions, cases, pages, include_gallery=include_gallery, run_date=run_date)
    (output / 'report.md').write_text(summary, encoding='utf-8')
    if include_gallery:
        (output / 'gallery.html').write_text('\n'.join(gallery(table, cases, output)), encoding='utf-8')
    readme_root = args.readme_root.resolve() if args.readme_root else ROOT
    if len(cases) == len(fixtures) or args.readme_root:
        update_readmes(readme_root, rows, table, engines, versions, complete=engines == tuple(ENGINES))
        print(f'Updated the README tables under {readme_root}')
    else:
        print(f'Left the READMEs unchanged: this run covers {len(cases)} of {len(fixtures)} documents.')
    print(summary.split('\n## Environment', 1)[0])


def main():
    parser = ArgumentParser(description='Score rendered page images against the pinned reference pages.')
    standalone = parser.add_mutually_exclusive_group()
    standalone.add_argument('--kanera-only', action='store_true', help='Score the Kanera-only run.')
    standalone.add_argument('--all', action='store_true', help='Score all comparison engines and refresh all README tables.')
    parser.add_argument('--aspose', action='store_true', help='Include the Aspose.Words run.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the Syncfusion DocIO run.')
    parser.add_argument('--rdocx', action='store_true', help='Include the rdocx run.')
    parser.add_argument('--gallery', action=BooleanOptionalAction,
                        help='Generate the screenshot gallery (default: on for the comparison).')
    parser.add_argument('--jobs', type=int, default=min(8, cpu_count() or 1),
                        help='Concurrent scoring processes (default: %(default)s).')
    parser.add_argument('--no-cache', action='store_true', help='Score every page again instead of reusing cached scores.')
    parser.add_argument('--output', type=Path, help=f'The conversion output directory (default: {OUTPUT.relative_to(ROOT)}).')
    parser.add_argument('--readme-root', type=Path,
                        help='Write the README tables under this directory instead of the repository, even for a '
                             'partial run (for testing the README writer on a copy).')
    args = parser.parse_args()
    if args.all:
        args.aspose = args.syncfusion = args.rdocx = True
    if args.kanera_only and (args.aspose or args.syncfusion or args.rdocx):
        parser.error('--kanera-only cannot be combined with comparison engine flags')
    if args.jobs < 1:
        parser.error('--jobs must be positive')
    root = args.output.resolve() if args.output else OUTPUT
    with benchmark_lock(root, stale=[(root, 'scoring-')]), TemporaryDirectory(prefix='scoring-', dir=root) as work:
        score_outputs(args, root, Path(work))


if __name__ == '__main__':
    main()
