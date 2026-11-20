from argparse import ArgumentParser, BooleanOptionalAction
from collections import Counter
from datetime import datetime, timezone
from html import escape
from platform import platform, processor
import csv
import json
import statistics
import subprocess
import unicodedata

import numpy as np
import pymupdf
from PIL import Image, ImageDraw, ImageFont

from fetch import BENCHMARK_FIXTURES, HERE, MANIFEST, UPSTREAM, OUTPUT, benchmark_reference, verify_reference
from ink import page_scores, render, render_reference
from readme import markdown_table, update_readmes

ENGINES = {
    'reference': 'Reference',
    'ours': '@kanera/docx2pdf',
    'eigenpal': 'EigenPal',
    'libreoffice': 'LibreOffice',
    'aspose': 'Aspose.Words (.NET)',
    'syncfusion': 'Syncfusion DocIO (.NET)',
    'dxpdf': 'dxpdf (Rust)',
    'rdocx': 'rdocx (Rust)',
    'office2pdf': 'office2pdf (Rust)',
}
# Engines that join the run only when their command-line flag is given.
OPTIONAL = {
    'aspose': 'aspose',
    'syncfusion': 'syncfusion',
    'dxpdf': 'rust',
    'rdocx': 'rust',
    'office2pdf': 'rust',
}
# Licensed engines join the run only when asked for; unlicensed output carries a watermark.
LICENSED = {
    'aspose': 'evaluation only. created with aspose.words',
    'syncfusion': 'trial version of syncfusion',
}


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
            if row['failed']:
                caption = f"Conversion failed after {row['cold_ms']:.1f} ms"
            else:
                caption = (f"{row['cold_ms']:.1f} ms cold · {row['warm_ms']:.1f} ms warm · "
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


def score_case(fixture, engines, metrics, visuals, *, output=None):
    name = fixture['case']
    directory = (output or OUTPUT) / 'cases' / name
    reference_dir, reference_fixture = benchmark_reference(fixture)
    verify_reference(reference_dir, reference_fixture)
    reference = render_reference(reference_dir, reference_fixture, directory / 'raster')
    if not reference:
        raise ValueError(f'No reference pages for {name}')
    reference_words = None
    if 'pdf' in reference_fixture:
        with pymupdf.open(reference_dir / reference_fixture['pdf']) as doc:
            reference_words = words(doc)
        if len(reference_words) != len(reference):
            raise ValueError(f'Inconsistent reference page count for {name}')

    rasters = {'reference': reference}
    extracted = {}
    for engine in engines[1:]:
        pdf = directory / f'{engine}.pdf'
        if 'error' in metrics[name, engine]:
            # A failed conversion has no pages, so every reference page scores zero.
            rasters[engine] = []
            extracted[engine] = []
            continue
        with pymupdf.open(pdf) as doc:
            if engine in LICENSED and any(LICENSED[engine] in page.get_text().lower() for page in doc):
                raise ValueError(f'Evaluation watermark detected in {name}')
            if reference_words is not None:
                extracted[engine] = words(doc)
            page_count = len(doc)
        rasters[engine] = render(pdf)
        if page_count != len(rasters[engine]):
            raise ValueError(f'Inconsistent page count for {name} ({engine})')

    rows = []
    for engine in engines[1:]:
        pages = rasters[engine]
        failed = 'error' in metrics[name, engine]
        scores = page_scores(reference, pages)
        recall, offset, matched = (page_metrics(reference_words, extracted[engine])
                                   if reference_words is not None else (None, None, None))
        rows.append({
            'case': name,
            'corpus': fixture['corpus'],
            'suite': fixture['suite'],
            'reference_application': fixture['reference_application'],
            'reference_format': 'pdf' if reference_words is not None else 'png',
            'engine': engine,
            'failed': failed,
            'cold_ms': round(metrics[name, engine]['cold_ms'], 2),
            'warm_ms': None if failed else metrics[name, engine]['warm_ms'],
            'peak_rss_kib': metrics[name, engine]['peak_rss_kib'],
            'reference_pages': len(reference),
            'pages': len(pages),
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
        scores = [row for row in rows if row['engine'] == engine]
        converted = [row for row in scores if not row['failed']]
        failed = len(scores) - len(converted)
        if not converted:
            yield (ENGINES[engine], *('–',) * 6, f'{failed}/{count}', f'0/{count}', '0.0%', '0.00 MiB')
            continue
        # Speed and memory cover successful conversions; failures score zero for accuracy.
        cold = statistics.median(row['cold_ms'] for row in converted)
        warm_times = [row['warm_ms'] for row in converted]
        warm = statistics.median(warm_times)
        warm_p95 = statistics.quantiles(warm_times, n=100, method='inclusive')[94] if len(warm_times) > 1 else warm
        throughput = 1000 / statistics.mean(warm_times)
        rss = [row['peak_rss_kib'] / 1024 for row in converted]
        matches = sum(row['pages'] == row['reference_pages'] for row in scores)
        ink = statistics.mean(row['ink_f1'] for row in scores)
        pdf_bytes = sum(row['pdf_bytes'] for row in scores)
        yield (ENGINES[engine], f'{cold:.1f} ms', f'{warm:.1f} ms', f'{warm_p95:.1f} ms',
               f'{throughput:.1f}', f'{statistics.median(rss):.1f} MiB', f'{max(rss):.1f} MiB',
               f'{failed}/{count}', f'{matches}/{count}', f'{ink:.1f}%', f'{pdf_bytes / 1024**2:.2f} MiB')


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
    lines.extend(['', 'Cold time includes process startup and file I/O. Warm time includes file I/O and reuses one process per engine after an unmeasured conversion. Warm p95 is the 95th percentile across timed documents. Warm docs/sec divides successful warm conversions by their total timed duration. Peak RSS, accuracy, and PDF sizes come from the cold runs. Times and memory cover successful conversions. A failed conversion scores zero ink overlap.',
                  '', '## Cases', ''])
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
    yield '</tbody></table><h2>Page comparisons</h2><nav aria-label="Benchmark cases"><ul>'
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


def main():
    parser = ArgumentParser(description='Score locally converted PDFs against the pinned reference pages.')
    parser.add_argument('--kanera-only', action='store_true', help='Score the Kanera-only run.')
    parser.add_argument('--gallery', action=BooleanOptionalAction,
                        help='Generate the screenshot gallery (default: on for the full comparison).')
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose run.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO run.')
    parser.add_argument('--rust', action='store_true', help='Include the dxpdf, rdocx and office2pdf runs.')
    args = parser.parse_args()
    if args.kanera_only and (args.aspose or args.syncfusion or args.rust):
        parser.error('--kanera-only cannot be combined with --aspose, --syncfusion or --rust')
    output = OUTPUT / 'kanera-only' if args.kanera_only else OUTPUT
    include_gallery = not args.kanera_only if args.gallery is None else args.gallery
    engines = ('reference', 'ours') if args.kanera_only else tuple(
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
    poppler = subprocess.run(['pdftoppm', '-v'], check=True, capture_output=True, text=True)
    versions['poppler'] = (poppler.stderr or poppler.stdout).splitlines()[0]
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
    rows = []
    for fixture in BENCHMARK_FIXTURES:
        print(f"Scoring {fixture['case']}", flush=True)
        rows.extend(score_case(fixture, engines, metrics, visuals, output=output))
    with (output / 'results.csv').open('w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=rows[0])
        writer.writeheader()
        writer.writerows(rows)
    count = len(BENCHMARK_FIXTURES)
    table = list(summary_table(rows, engines, count))
    pages = sum(row['reference_pages'] for row in rows if row['engine'] == 'ours')
    summary = report(table, versions, count, pages, output=output, include_gallery=include_gallery, run_date=run_date)
    (output / 'report.md').write_text(summary, encoding='utf-8')
    if include_gallery:
        (output / 'gallery.html').write_text('\n'.join(gallery(table, output)), encoding='utf-8')
    if not args.kanera_only:
        update_readmes(HERE.parents[1], rows, table, engines[1:], versions, run_date)
    print(summary.split('\n## Cases', 1)[0])


if __name__ == '__main__':
    main()
