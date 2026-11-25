from argparse import ArgumentParser, BooleanOptionalAction
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, timezone
from html import escape
from os import cpu_count
from pathlib import Path
from platform import platform, processor
import csv
import json
import statistics
import subprocess
from tempfile import TemporaryDirectory
import unicodedata

import numpy as np
import PIL
import pymupdf
from PIL import Image, ImageDraw, ImageFont

from fetch import (BENCHMARK_FIXTURES, HERE, MANIFEST, UPSTREAM, OUTPUT, benchmark_lock,
                   benchmark_reference, child_environment, verify_reference)
from pdf_text import extract as extract_text
from ink import DPI, INK_THRESHOLD, INK_TOLERANCE_PX, page_scores, render, render_reference
from score_cache import FORMAT, ScoreCache, code_digests, digest, file_sha256, rendering_files
from readme import (markdown_table, sampled_note, update_measured_readme, update_mini_pdf_breakdown, update_readmes,
                    update_suite_scores)

ENGINES = {
    'reference': 'Reference',
    'ours': '@kanera/docx2pdf',
    'eigenpal': 'EigenPal',
    'libreoffice': 'LibreOffice',
    'aspose': 'Aspose.Words (.NET)',
    'syncfusion': 'Syncfusion DocIO (.NET)',
    'minipdf': 'MiniPdf (.NET)',
    'dxpdf': 'dxpdf (Rust)',
    'rdocx': 'rdocx (Rust)',
    'office2pdf': 'office2pdf (Rust)',
}
# Engines that join the run only when their command-line flag is given.
OPTIONAL = {
    'aspose': 'aspose',
    'syncfusion': 'syncfusion',
    'minipdf': 'minipdf',
    'dxpdf': 'rust',
    'rdocx': 'rust',
    'office2pdf': 'rust',
}
# Licensed engines join the run only when asked for; unlicensed output carries a watermark.
LICENSED = {
    'aspose': 'evaluation only. created with aspose.words',
    'syncfusion': 'trial version of syncfusion',
}
# Scores and reference pages are reused while their inputs, tools and the code below are unchanged
# (see score_cache.py). The cache is shared by every mode and by accuracy.py.
SCORE_CACHE = OUTPUT / 'score-cache'
SCORE_CODE = (HERE / 'score.py', HERE / 'ink.py', HERE / 'pdf_text.py', HERE / 'score_cache.py')
REFERENCE_CODE = (HERE / 'ink.py', HERE / 'score_cache.py')


def words(doc):
    pages = []
    for page in doc:
        result = []
        for x0, y0, _, _, text, *_ in page.get_text('words'):
            token = unicodedata.normalize('NFKC', text).strip()
            if token:
                result.append((token, x0, y0))
        pages.append(result)
    return pages


def page_metrics(baseline, output):
    # Unique tokens give an unambiguous position match within a page.
    matched = 0
    offsets = []
    for word_page, pdf_page in zip(baseline, output):
        word_counts = Counter(token for token, *_ in word_page)
        pdf_counts = Counter(token for token, *_ in pdf_page)
        matched += sum((word_counts & pdf_counts).values())
        unique = {token: (x, y) for token, x, y in pdf_page if pdf_counts[token] == 1}
        for token, x, y in word_page:
            if word_counts[token] == 1 and token in unique:
                pdf_x, pdf_y = unique[token]
                offsets.append(float(np.hypot(pdf_x - x, pdf_y - y)))
    total_words = sum(map(len, baseline))
    recall = matched / total_words if total_words else None
    median_offset = statistics.median(offsets) if offsets else None
    return recall, median_offset, len(offsets)


LABEL_FONT = ImageFont.load_default(size=16)
SMALL_FONT = ImageFont.load_default(size=13)


def comparison(fixture, rasters, rows, page, out):
    case = fixture['case']
    width, height = 500, 690
    spacing = 16
    images = {}
    for engine, pages in rasters.items():
        if page < len(pages):
            with Image.open(pages[page]) as image:
                images[engine] = image.convert('RGB')
    scale = min(1, width / max(p.width for p in images.values()), height / max(p.height for p in images.values()))
    canvas = Image.new('RGB', (len(rasters) * (width + spacing) + spacing, height + 111), '#edf0f4')
    draw = ImageDraw.Draw(canvas)
    draw.text((spacing, 12), f'{case.replace("-", " ").title()}  /  Page {page + 1}', font=LABEL_FONT, fill='#162437')
    by_engine = {row['engine']: row for row in rows}
    for i, engine in enumerate(rasters):
        x, y = spacing + i * (width + spacing), 87
        label = fixture['reference_application'] if engine == 'reference' else ENGINES[engine]
        draw.text((x, 41), label, font=LABEL_FONT, fill='#263849')
        if engine in by_engine:
            row = by_engine[engine]
            if row['unreadable_pdf']:
                caption = f"Unreadable PDF ({row['unreadable_pdf']}), scored as failed"
            elif row['failed']:
                caption = f"Conversion failed after {row['cold_ms']:.1f} ms"
            elif row['untimed']:
                caption = f"Untimed (sampled engine) · {row['ink_f1']:.1f}% document ink overlap"
            else:
                warm = 'no warm time' if row['warm_ms'] is None else f"{row['warm_ms']:.1f} ms warm"
                caption = (f"{row['cold_ms']:.1f} ms cold · {warm} · "
                           f"{row['peak_rss_kib'] / 1024:.1f} MiB · {row['ink_f1']:.1f}% document ink overlap")
            draw.text((x, 64), caption, font=SMALL_FONT, fill='#263849')
        if engine not in images:
            failed = engine in by_engine and by_engine[engine]['failed']
            draw.text((x, y + 20), 'Conversion failed' if failed else 'No corresponding page',
                      font=SMALL_FONT, fill='#263849')
            continue
        image = images[engine]
        thumb = image.resize((round(image.width * scale), round(image.height * scale)), Image.Resampling.LANCZOS)
        canvas.paste(thumb, (x, y))
        draw.rectangle((x, y, x + thumb.width - 1, y + thumb.height - 1), outline='#9ca6b2')
    canvas.save(out / f'{case}-{page + 1:02d}.png')


def cache_context(poppler, env):
    """What every cached score and reference page depends on besides its own inputs."""
    root = HERE.parents[1]
    stats, font_env = rendering_files(env)
    tools = {'poppler': poppler, 'fonts': digest([stats, font_env]), 'pillow': PIL.__version__}
    settings = {'dpi': DPI, 'ink_threshold': INK_THRESHOLD, 'tolerance_px': INK_TOLERANCE_PX}
    return {
        'reference': {'format': FORMAT, 'kind': 'reference', 'tools': tools, 'dpi': DPI,
                      'code': code_digests(REFERENCE_CODE, root)},
        'score': {'format': FORMAT, 'kind': 'score',
                  'tools': dict(tools, pymupdf=pymupdf.version, numpy=np.__version__),
                  'settings': settings, 'code': code_digests(SCORE_CODE, root)},
    }


def reference_source(fixture):
    """The digests a reference's pages are rendered from."""
    if 'pdf' in fixture:
        return {'pdf': fixture['pdf_sha256']}
    return {'images': [image['sha256'] for image in fixture['reference_images']], 'dpi': fixture['reference_dpi']}


def score_pdf(pdf, reference, reference_words, env):
    """Score one converted PDF against the reference pages. Returns the result and the rendered pages."""
    text = extract_text(pdf, env)
    if text is None:
        return {'unreadable': 'text extraction crashed'}, []
    pdf_words, pdf_text = text
    # Recorded for every licensed engine, so the result does not depend on which engine wrote the PDF.
    watermarks = sorted(engine for engine, mark in LICENSED.items() if mark in pdf_text)
    try:
        pages = render(pdf, env=env)
    except subprocess.CalledProcessError:
        return {'unreadable': 'rasterization failed', 'watermarks': watermarks}, []
    if len(pdf_words) != len(pages):
        return {'unreadable': 'page counts disagree between text extraction and rasterization',
                'watermarks': watermarks}, []
    text_metrics = page_metrics(reference_words, pdf_words) if reference_words is not None else (None, None, None)
    return {'unreadable': None, 'watermarks': watermarks, 'pages': len(pages),
            'page_scores': page_scores(reference, pages), 'text_metrics': list(text_metrics)}, pages


def score_case(fixture, engines, metrics, visuals, *, output=None, cache=None, context=None, env=None):
    name = fixture['case']
    directory = (output or OUTPUT) / 'cases' / name
    reference_dir, reference_fixture = benchmark_reference(fixture)
    verify_reference(reference_dir, reference_fixture)
    source = reference_source(reference_fixture)
    if cache is None:
        reference = render_reference(reference_dir, reference_fixture, directory / 'raster', env=env)
    else:
        reference = cache.reference_pages(
            digest({**context['reference'], 'reference': source}),
            lambda target: render_reference(reference_dir, reference_fixture, target, env=env))
    if not reference:
        raise ValueError(f'No reference pages for {name}')
    reference_words = None
    if 'pdf' in reference_fixture:
        with pymupdf.open(reference_dir / reference_fixture['pdf']) as doc:
            reference_words = words(doc)
        if len(reference_words) != len(reference):
            raise ValueError(f'Inconsistent reference page count for {name}')

    # A sampled engine skips documents that nothing scores; they are left out of everything.
    engines = [engine for engine in engines[1:] if not metrics[name, engine].get('skipped')]
    rasters = {'reference': reference}
    results = {}
    invalid = {}
    for engine in engines:
        pdf = directory / f'{engine}.pdf'
        rasters[engine] = []
        if 'error' in metrics[name, engine]:
            # A failed conversion has no pages, so every reference page scores zero.
            continue
        key = result = None
        if cache is not None:
            key = digest({**context['score'], 'candidate': file_sha256(pdf), 'reference': source})
            result = cache.get(key)
        if result is None:
            result, rasters[engine] = score_pdf(pdf, reference, reference_words, env)
            # An unreadable PDF can be a passing crash or timeout, so it is not cached.
            if key is not None and result['unreadable'] is None:
                cache.put(key, result)
        elif visuals is not None:
            # The side-by-side images need the pages even when the score is cached.
            rasters[engine] = render(pdf, env=env)
        if engine in result.get('watermarks', ()):
            raise ValueError(f'Evaluation watermark detected in {name}')
        if result['unreadable']:
            # A PDF the readers cannot process counts as a failed conversion and scores zero.
            print(f"Unreadable PDF for {name} ({engine}): {result['unreadable']}; scoring it as failed", flush=True)
            invalid[engine] = result['unreadable']
            rasters[engine] = []
            continue
        results[engine] = result

    rows = []
    for engine in engines:
        failed = 'error' in metrics[name, engine] or engine in invalid
        if engine in results:
            result = results[engine]
            page_count, scores = result['pages'], result['page_scores']
            recall, offset, matched = result['text_metrics']
        else:
            page_count, scores = 0, page_scores(reference, [])
            recall, offset, matched = (page_metrics(reference_words, [])
                                       if reference_words is not None else (None, None, None))
        rows.append({
            'case': name,
            'corpus': fixture['corpus'],
            'suite': fixture['suite'],
            'reference_application': fixture['reference_application'],
            'reference_format': 'pdf' if reference_words is not None else 'png',
            'engine': engine,
            'failed': failed,
            'unreadable_pdf': invalid.get(engine),
            # Untimed conversions of a sampled engine count for accuracy and failures only.
            'untimed': bool(metrics[name, engine].get('untimed')),
            'cold_ms': round(metrics[name, engine]['cold_ms'], 2),
            # A document a warm batch stopped on has no warm time (see resumable_warm in convert.py).
            'warm_ms': None if failed else metrics[name, engine].get('warm_ms'),
            'peak_rss_kib': metrics[name, engine]['peak_rss_kib'],
            'reference_pages': len(reference),
            'pages': page_count,
            'ink_f1': round(statistics.mean(scores) * 100, 2),
            'same_page_words': round(recall * 100, 2) if recall is not None else None,
            'median_offset_pt': round(offset, 2) if offset is not None else None,
            'unique_word_matches': matched,
            'page_ink_f1': [round(score * 100, 2) for score in scores],
            'pdf_bytes': 0 if failed else (directory / f'{engine}.pdf').stat().st_size,
        })

    if visuals is not None:
        for page in range(max(map(len, rasters.values()))):
            comparison(fixture, rasters, rows, page, visuals)
    return rows


HEADERS = ('Converter', 'Median cold time', 'Median warm time', 'Warm p95 latency', 'Warm docs/sec',
           'Median cold peak RSS', 'Max cold peak RSS', 'Failed conversions', 'Matching page count',
           'Mean ink overlap', 'Total PDF size')


def summary_table(rows, engines, count):
    yield HEADERS
    for engine in engines[1:]:
        # Skipped documents have no rows, so failures and page counts cover the converted documents.
        scores = [row for row in rows if row['engine'] == engine]
        total = len(scores)
        converted = [row for row in scores if not row['failed']]
        failed = total - len(converted)
        # Speed and memory cover successful timed conversions; failures score zero for accuracy.
        timed = [row for row in converted if not row['untimed']]
        # † marks an engine timed on the sample only.
        mark = ' †' if total < count or len(timed) < len(converted) else ''
        timing = ('–',) * 6
        if timed:
            cold = statistics.median(row['cold_ms'] for row in timed)
            warm_times = [row['warm_ms'] for row in timed if row['warm_ms'] is not None]
            warm_columns = ('–',) * 3
            if warm_times:
                warm = statistics.median(warm_times)
                warm_p95 = statistics.quantiles(warm_times, n=100, method='inclusive')[94] if len(warm_times) > 1 else warm
                throughput = 1000 / statistics.mean(warm_times)
                warm_columns = (f'{warm:.1f} ms', f'{warm_p95:.1f} ms', f'{throughput:.1f}')
            rss = [row['peak_rss_kib'] / 1024 for row in timed]
            timing = tuple(value if value == '–' else value + mark for value in (
                f'{cold:.1f} ms', *warm_columns, f'{statistics.median(rss):.1f} MiB', f'{max(rss):.1f} MiB'))
        matches = sum(row['pages'] == row['reference_pages'] for row in scores)
        ink = statistics.mean(row['ink_f1'] for row in scores) if scores else 0
        pdf_bytes = sum(row['pdf_bytes'] for row in scores)
        yield (ENGINES[engine], *timing, f'{failed}/{total}', f'{matches}/{total}', f'{ink:.1f}%',
               f'{pdf_bytes / 1024**2:.2f} MiB')


def case_images(output=None, include_gallery=True):
    output = output or OUTPUT
    for fixture in BENCHMARK_FIXTURES:
        name = fixture['case']
        images = sorted((output / 'visuals').glob(f'{name}-*.png')) if include_gallery else []
        yield name, [image.relative_to(output).as_posix() for image in images]


def report(table, versions, count, pages, *, output=None, include_gallery=True, run_date=None):
    lines = [
        '# DOCX-to-PDF benchmark',
        '',
        f'{count} DOCX files compared with {pages} reference pages.',
    ]
    if run_date:
        lines.append(f'Run date {run_date} (UTC).')
    lines.extend([
        'Reference applications and formats are recorded per case in the CSV. Text metrics apply to PDF references.',
        '',
        f"[Hugging Face corpus]({MANIFEST['source']}/tree/{MANIFEST['revision']}) · "
        f"[Vendored fixtures]({UPSTREAM['source']}/commits/{UPSTREAM['revision']})"
        + (' · [Screenshot gallery](gallery.html)' if include_gallery else ''),
        '',
        markdown_table(table[0], table[1:]),
    ])
    lines.extend(['', 'Cold time includes process startup and file I/O. Warm time includes file I/O and reuses one process per engine after an unmeasured conversion. Warm p95 is the 95th percentile across timed documents. Warm docs/sec divides successful warm conversions by their total timed duration. Peak RSS, accuracy, and PDF sizes come from the cold runs. Times and memory cover successful conversions. A failed conversion scores zero ink overlap.'])
    if any(cell.endswith('†') for row in table[1:] for cell in row):
        lines.extend(['', sampled_note()])
    lines.extend(['', '## Cases', ''])
    for name, images in case_images(output, include_gallery):
        links = ' · '.join(f'[{i}]({image})' for i, image in enumerate(images, 1))
        lines.append(f'- {name} ({links})' if images else f'- {name}')
    lines.extend(['', '## Environment', '', f'- Host ({platform()}, {processor() or "processor unspecified"})'])
    lines.extend(f'- {key} ({value})' for key, value in versions.items())
    lines.extend(['', '[Raw scores](results.csv)', ''])
    return '\n'.join(lines)


def gallery(table, output=None):
    yield '''<!doctype html>
<html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>DOCX-to-PDF benchmark</title>
<style>
body { font: 16px system-ui, sans-serif; margin: 1.5rem; }
th, td { padding: .4rem .65rem; text-align: left; }
img { display: block; height: auto; width: 100%; max-width: 2560px; }
figure { margin: 0 0 1.5rem; }
figcaption { margin-bottom: .3rem; }
</style>
<h1>DOCX-to-PDF benchmark</h1><p><a href="report.md">Report</a> · <a href="results.csv">Raw scores</a></p>
<table><thead>'''
    yield '<tr>' + ''.join(f'<th scope="col">{escape(cell)}</th>' for cell in table[0]) + '</tr></thead><tbody>'
    for row in table[1:]:
        yield '<tr>' + ''.join(f'<td>{escape(cell)}</td>' for cell in row) + '</tr>'
    yield '</tbody></table>'
    if any(cell.endswith('†') for row in table[1:] for cell in row):
        yield f'<p>{escape(sampled_note())}</p>'
    yield '<h2>Page comparisons</h2><nav aria-label="Benchmark cases"><ul>'
    for name, _ in case_images(output):
        yield f'<li><a href="#{escape(name)}">{escape(name)}</a></li>'
    yield '</ul></nav>'
    for name, images in case_images(output):
        yield f'<section id="{escape(name)}"><h3>{escape(name)}</h3>'
        for page, image in enumerate(images, 1):
            path = escape(image)
            yield (f'<figure><figcaption>Page {page}</figcaption><a href="{path}">'
                   f'<img src="{path}" loading="lazy" '
                   f'alt="Side-by-side comparison of {escape(name)}, page {page}"></a></figure>')
        yield '</section>'
    yield '</html>'


def score_outputs(args, work):
    standalone = 'ours' if args.kanera_only else 'minipdf' if args.minipdf_only else None
    directory = 'kanera-only' if args.kanera_only else 'minipdf-only' if args.minipdf_only else ''
    output = OUTPUT / directory
    include_gallery = standalone is None if args.gallery is None else args.gallery
    engines = ('reference', standalone) if standalone else tuple(
        engine for engine in ENGINES if engine not in OPTIONAL or getattr(args, OPTIONAL[engine])
    )
    log = output / 'conversion.json'
    data = json.loads(log.read_text())
    completed_at = datetime.fromisoformat(data['completed_at']) if 'completed_at' in data else (
        datetime.fromtimestamp(log.stat().st_mtime, timezone.utc)
    )
    run_date = completed_at.astimezone(timezone.utc).date().isoformat()
    versions = data['versions']
    metrics = {(row['case'], row['engine']): row for row in data['results']}
    # Child processes keep their temporary files in this run's work directory.
    env = child_environment(work / 'tmp')
    poppler = subprocess.run(['pdftoppm', '-v'], check=True, capture_output=True, text=True, env=env)
    versions['poppler'] = (poppler.stderr or poppler.stdout).splitlines()[0]
    cache = ScoreCache(SCORE_CACHE, work, read=not args.no_cache)
    context = cache_context(versions['poppler'], env)
    expected = {(fixture['case'], engine) for fixture in BENCHMARK_FIXTURES for engine in engines[1:]}
    if metrics.keys() != expected or len(metrics) != len(data['results']):
        raise ValueError('Conversion logs do not cover the requested engines and fixtures')

    visuals = output / 'visuals'
    for old in visuals.glob('*.png'):
        old.unlink()
    if include_gallery:
        visuals.mkdir(exist_ok=True)
    else:
        (output / 'gallery.html').unlink(missing_ok=True)
        visuals = None
    scored = {}
    print(f'Scoring {len(BENCHMARK_FIXTURES)} documents with {args.jobs} workers', flush=True)
    with ProcessPoolExecutor(max_workers=args.jobs) as pool:
        futures = {}
        for fixture in BENCHMARK_FIXTURES:
            name = fixture['case']
            case_metrics = {(name, engine): metrics[name, engine] for engine in engines[1:]}
            future = pool.submit(score_case, fixture, engines, case_metrics, visuals, output=output,
                                 cache=cache, context=context, env=env)
            futures[future] = name
        try:
            for future in as_completed(futures):
                name = futures[future]
                scored[name] = future.result()
                print(f'Scored {name}', flush=True)
        except BaseException:
            pool.shutdown(cancel_futures=True)
            raise
    rows = [row for fixture in BENCHMARK_FIXTURES for row in scored[fixture['case']]]
    with (output / 'results.csv').open('w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=rows[0])
        writer.writeheader()
        writer.writerows(rows)
    count = len(BENCHMARK_FIXTURES)
    table = list(summary_table(rows, engines, count))
    pages = sum({row['case']: row['reference_pages'] for row in rows}.values())
    summary = report(table, versions, count, pages, output=output, include_gallery=include_gallery, run_date=run_date)
    (output / 'report.md').write_text(summary, encoding='utf-8')
    if include_gallery:
        (output / 'gallery.html').write_text('\n'.join(gallery(table, output)), encoding='utf-8')
    if engines == tuple(ENGINES):
        update_readmes(HERE.parents[1], rows, table, engines[1:], versions)
    else:
        update_measured_readme(HERE.parents[1], table, versions, count, engines[1:])
        if standalone == 'minipdf':
            recorded = OUTPUT / 'results.csv'
            kanera_scores = {}
            if recorded.is_file():
                with recorded.open(newline='', encoding='utf-8') as file:
                    kanera_scores = {row['case']: float(row['ink_f1']) for row in csv.DictReader(file)
                                     if row['engine'] == 'ours'}
            update_mini_pdf_breakdown(HERE.parents[1], rows, ENGINES[standalone], kanera_scores)
        else:
            update_suite_scores(HERE.parents[1], rows, {ENGINES[engine]: engine for engine in engines[1:]})
    print(summary.split('\n## Cases', 1)[0])


def main():
    parser = ArgumentParser(description='Score locally converted PDFs against the pinned reference pages.')
    standalone = parser.add_mutually_exclusive_group()
    standalone.add_argument('--kanera-only', action='store_true', help='Score the Kanera-only run.')
    standalone.add_argument('--minipdf-only', action='store_true', help='Score the MiniPdf-only run.')
    standalone.add_argument('--all', action='store_true', help='Score all comparison engines and refresh all README tables.')
    parser.add_argument('--gallery', action=BooleanOptionalAction,
                        help='Generate the screenshot gallery (default: on for the full comparison).')
    parser.add_argument('--jobs', type=int, default=min(8, cpu_count() or 1),
                        help='Concurrent scoring processes (default: %(default)s).')
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose run.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO run.')
    parser.add_argument('--minipdf', action='store_true', help='Include MiniPdf in the full comparison.')
    parser.add_argument('--rust', action='store_true', help='Include the dxpdf, rdocx and office2pdf runs.')
    parser.add_argument('--no-cache', action='store_true',
                        help='Score every PDF and render every reference again instead of reusing cached results.')
    args = parser.parse_args()
    if args.all:
        args.aspose = args.syncfusion = args.minipdf = args.rust = True
    if (args.kanera_only or args.minipdf_only) and (args.aspose or args.syncfusion or args.rust or args.minipdf):
        parser.error('Standalone modes cannot be combined with full-comparison engine flags')
    if args.jobs < 1:
        parser.error('--jobs must be positive')
    with (benchmark_lock(OUTPUT, stale=[(OUTPUT, 'scoring-')]),
          TemporaryDirectory(prefix='scoring-', dir=OUTPUT) as work):
        score_outputs(args, Path(work))


if __name__ == '__main__':
    main()
