from argparse import ArgumentParser
from collections import Counter
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

from fetch import BENCHMARK_FIXTURES, MANIFEST, UPSTREAM, OUTPUT, benchmark_reference, verify_reference
from ink import page_scores, render, render_reference

ENGINES = {
    'reference': 'Reference',
    'ours': '@happendev/docx2pdf',
    'eigenpal': 'EigenPal',
    'libreoffice': 'LibreOffice',
    'aspose': 'Aspose.Words for .NET',
    'syncfusion': 'Syncfusion DocIO',
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
            draw.text((x, 64), f"{row['cold_ms']:.1f} ms cold · {row['warm_ms']:.1f} ms warm · "
                      f"{row['peak_rss_kib'] / 1024:.1f} MiB · {row['ink_f1']:.1f}% document ink overlap",
                      font=SMALL_FONT, fill='#263849')
        if engine not in images:
            draw.text((x, y + 20), 'No corresponding page', font=SMALL_FONT, fill='#263849')
            continue
        image = images[engine]
        thumb = image.resize((round(image.width * scale), round(image.height * scale)), Image.Resampling.LANCZOS)
        canvas.paste(thumb, (x, y))
        draw.rectangle((x, y, x + thumb.width - 1, y + thumb.height - 1), outline='#9ca6b2')
    canvas.save(out / f'{case}-{page + 1:02d}.png')


def score_case(fixture, engines, metrics, visuals):
    name = fixture['case']
    directory = OUTPUT / 'cases' / name
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
            'cold_ms': round(metrics[name, engine]['cold_ms'], 2),
            'warm_ms': round(metrics[name, engine]['warm_ms'], 2),
            'peak_rss_kib': metrics[name, engine]['peak_rss_kib'],
            'reference_pages': len(reference),
            'pages': len(pages),
            'ink_f1': round(statistics.mean(scores) * 100, 2),
            'same_page_words': round(recall * 100, 2) if recall is not None else None,
            'median_offset_pt': round(offset, 2) if offset is not None else None,
            'unique_word_matches': matched,
            'page_ink_f1': [round(score * 100, 2) for score in scores],
            'pdf_bytes': (directory / f'{engine}.pdf').stat().st_size,
        })

    for page in range(max(map(len, rasters.values()))):
        comparison(fixture, rasters, rows, page, visuals)
    return rows


HEADERS = ('Converter', 'Median cold time', 'Median warm time', 'Median cold peak RSS',
           'Max cold peak RSS', 'Matching page count', 'Mean ink overlap', 'Total PDF size')


def summary_table(rows, engines, count):
    yield HEADERS
    for engine in engines[1:]:
        scores = [row for row in rows if row['engine'] == engine]
        cold = statistics.median(row['cold_ms'] for row in scores)
        warm = statistics.median(row['warm_ms'] for row in scores)
        rss = [row['peak_rss_kib'] / 1024 for row in scores]
        matches = sum(row['pages'] == row['reference_pages'] for row in scores)
        ink = statistics.mean(row['ink_f1'] for row in scores)
        pdf_bytes = sum(row['pdf_bytes'] for row in scores)
        yield (ENGINES[engine], f'{cold:.1f} ms', f'{warm:.1f} ms', f'{statistics.median(rss):.1f} MiB',
               f'{max(rss):.1f} MiB', f'{matches}/{count}', f'{ink:.1f}%', f'{pdf_bytes / 1024**2:.2f} MiB')


def case_images():
    for fixture in BENCHMARK_FIXTURES:
        name = fixture['case']
        images = sorted((OUTPUT / 'visuals').glob(f'{name}-*.png'))
        yield name, [image.relative_to(OUTPUT).as_posix() for image in images]


def report(table, versions, count, pages):
    lines = [
        '# DOCX-to-PDF benchmark',
        '',
        f'{count} DOCX files compared with {pages} reference pages.',
        'Reference applications and formats are recorded per case in the CSV. Text metrics apply to PDF references.',
        '',
        f"[Hugging Face corpus]({MANIFEST['source']}/tree/{MANIFEST['revision']}) · "
        f"[Vendored fixtures]({UPSTREAM['source']}/commits/{UPSTREAM['revision']}) · "
        '[Screenshot gallery](gallery.html)',
        '',
        '| ' + ' | '.join(table[0]) + ' |',
        '| --- |' + ' ---: |' * (len(table[0]) - 1),
    ]
    lines.extend('| ' + ' | '.join(row) + ' |' for row in table[1:])
    lines.extend(['', 'Cold time includes process startup and file I/O. Warm time includes file I/O and reuses one process per engine after an unmeasured conversion. Peak RSS, accuracy, and PDF sizes come from the cold runs.',
                  '', '## Cases', ''])
    for name, images in case_images():
        links = ' · '.join(f'[{i}]({image})' for i, image in enumerate(images, 1))
        lines.append(f'- {name} ({links})')
    lines.extend(['', '## Environment', '', f'- Host ({platform()}, {processor() or "processor unspecified"})'])
    lines.extend(f'- {key} ({value})' for key, value in versions.items())
    lines.extend(['', '[Raw scores](results.csv)', ''])
    return '\n'.join(lines)


def gallery(table):
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
    for name, _ in case_images():
        yield f'<li><a href="#{escape(name)}">{escape(name)}</a></li>'
    yield '</ul></nav>'
    for name, images in case_images():
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
    parser.add_argument('--aspose', action='store_true', help='Include the licensed Aspose run.')
    parser.add_argument('--syncfusion', action='store_true', help='Include the licensed Syncfusion DocIO run.')
    args = parser.parse_args()
    engines = tuple(engine for engine in ENGINES if engine not in LICENSED or getattr(args, engine))
    data = json.loads((OUTPUT / 'conversion.json').read_text())
    versions = data['versions']
    metrics = {(row['case'], row['engine']): row for row in data['results']}
    poppler = subprocess.run(['pdftoppm', '-v'], check=True, capture_output=True, text=True)
    versions['poppler'] = (poppler.stderr or poppler.stdout).splitlines()[0]
    expected = {(fixture['case'], engine) for fixture in BENCHMARK_FIXTURES for engine in engines[1:]}
    if metrics.keys() != expected or len(metrics) != len(data['results']):
        raise ValueError('Conversion logs do not cover the requested engines and fixtures')

    visuals = OUTPUT / 'visuals'
    visuals.mkdir(exist_ok=True)
    for old in visuals.glob('*.png'):
        old.unlink()
    rows = []
    for fixture in BENCHMARK_FIXTURES:
        print(f"Scoring {fixture['case']}", flush=True)
        rows.extend(score_case(fixture, engines, metrics, visuals))
    with (OUTPUT / 'results.csv').open('w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=rows[0])
        writer.writeheader()
        writer.writerows(rows)
    count = len(BENCHMARK_FIXTURES)
    table = list(summary_table(rows, engines, count))
    pages = sum(row['reference_pages'] for row in rows if row['engine'] == 'ours')
    summary = report(table, versions, count, pages)
    (OUTPUT / 'report.md').write_text(summary, encoding='utf-8')
    (OUTPUT / 'gallery.html').write_text('\n'.join(gallery(table)), encoding='utf-8')
    print(summary.split('\n## Cases', 1)[0])


if __name__ == '__main__':
    main()
