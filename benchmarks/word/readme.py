"""Refresh the benchmark summaries in the repository READMEs."""

from collections import defaultdict
from fcntl import LOCK_EX, flock
from pathlib import Path
from shutil import copymode
from statistics import mean
from tempfile import NamedTemporaryFile


HIGHER_IS_BETTER = {'Mean ink overlap', 'Matching page count', 'Warm docs/sec'}
MEMORY_COLUMNS = ('Median cold peak RSS', 'Max cold peak RSS')
HEADLINE_METRICS = {
    'Visual match': 'Mean ink overlap',
    'Page counts matched': 'Matching page count',
    'Fresh process': 'Median cold time',
    'Reused process': 'Median warm time',
    'Median peak memory': 'Median cold peak RSS',
}
SUITES = (
    ('classic', 'classic', 'Microsoft 365 Word page images'),
    ('sample', 'Hugging Face', 'Word for Mac PDFs'),
    ('word_issues', 'issues', 'Microsoft 365 Word PDFs'),
    ('office_issues', 'issues', 'LibreOffice PDFs'),
)


def markdown_table(headers, rows):
    lines = ['| ' + ' | '.join(headers) + ' |', '| --- |' + ' ---: |' * (len(headers) - 1)]
    lines.extend('| ' + ' | '.join(row) + ' |' for row in rows)
    return '\n'.join(lines)


def highlight_best(values, *, higher_is_better, eligible=None):
    scores = [
        float(value.split()[0].split('/')[0].rstrip('%'))
        if value != '–' and (eligible is None or eligible[index]) else None
        for index, value in enumerate(values)
    ]
    best = (max if higher_is_better else min)((score for score in scores if score is not None), default=None)
    return [f'**{value}**' if score is not None and score == best else value
            for value, score in zip(values, scores, strict=True)]


def metric_cells(results, engines, column):
    values = [results[engine][column] for engine in engines]
    eligible = ([results[engine]['Failed conversions'].startswith('0/') for engine in engines]
                if column == 'Total PDF size' else None)
    return highlight_best(values, higher_is_better=column in HIGHER_IS_BETTER, eligible=eligible)


def section_parts(text, name):
    start = f'<!-- {name}:start -->'
    end = f'<!-- {name}:end -->'
    before, _, rest = text.partition(start)
    body, _, after = rest.partition(end)
    if text.count(start) != 1 or text.count(end) != 1 or end not in rest:
        raise ValueError(f'Expected exactly one {name} section in the README')
    return before + start, body, end + after


def replace_section(text, name, body):
    before, _, after = section_parts(text, name)
    return before + '\n' + body + '\n' + after


def edit_main_readme(root, transform):
    # The benchmark and accuracy check can finish independently.
    output = root / 'out'
    output.mkdir(exist_ok=True)
    readme = root / 'README.md'
    with (output / '.readme.lock').open('a') as lock:
        flock(lock, LOCK_EX)
        original = readme.read_text(encoding='utf-8')
        updated = transform(original)
        if updated != original:
            with NamedTemporaryFile(mode='w', encoding='utf-8', dir=root, delete=False) as temporary:
                staged = Path(temporary.name)
                try:
                    temporary.write(updated)
                    temporary.close()
                    copymode(readme, staged)
                    staged.replace(readme)
                finally:
                    staged.unlink(missing_ok=True)


def read_table(lines, start):
    end = next((i for i in range(start, len(lines)) if not lines[i].startswith('|')), len(lines))
    headers, _, *rows = [[cell.strip().strip('*') for cell in line.strip('|').split('|')]
                         for line in lines[start:end]]
    return headers, rows, end


def refresh_engine_table(text, section, values, name, count):
    before, body, after = section_parts(text, section)
    lines = body.splitlines()
    table_start = next(i for i, line in enumerate(lines) if line.startswith('|'))
    headers, rows, table_end = read_table(lines, table_start)
    displayed = {row[0]: dict(zip(headers[1:], row[1:], strict=True)) for row in rows}
    existing = next((label for label in displayed if label.split()[0] == name.split()[0]), None)
    current = existing or name
    displayed[current] = {header: values[HEADLINE_METRICS.get(header, header)] for header in headers[1:]}

    # Recompute the highlights against the recorded results for the same document count.
    page_count_column = 'Page counts matched' if 'Page counts matched' in headers else 'Matching page count'
    comparison_count = next(result[page_count_column].split('/')[1] for engine, result in displayed.items()
                            if engine != current)
    compared = tuple(engine for engine in displayed if engine != current or str(count) == comparison_count)
    metrics = {engine: {HEADLINE_METRICS.get(column, column): value for column, value in result.items()}
               for engine, result in displayed.items()}
    for header in headers[1:]:
        key = HEADLINE_METRICS.get(header, header)
        for engine, value in zip(compared, metric_cells(metrics, compared, key), strict=True):
            displayed[engine][header] = value
    ordered = displayed if existing else sorted(
        displayed, key=lambda engine: -float(metrics[engine]['Mean ink overlap'].rstrip('%')))
    published = [(name if engine == current else engine, *(displayed[engine][header] for header in headers[1:]))
                 for engine in ordered]
    intro = '\n'.join(lines[:table_start]).strip('\n')
    tail = '\n'.join(lines[table_end:]).strip('\n')
    body = '\n\n'.join(part for part in (intro, markdown_table(headers, published), tail) if part)
    return before + '\n' + body + '\n' + after


def update_single_engine_readme(root, table, versions, count, engine):
    values = dict(zip(table[0], table[1], strict=True))
    converter = values['Converter']
    if versions.get(engine):
        converter += ' ' + versions[engine].split()[0]

    def transform(text):
        name = 'Kanera' if engine == 'ours' else values['Converter']
        text = refresh_engine_table(text, 'benchmark-summary', values, name, count)
        return refresh_engine_table(text, 'benchmark-results', values, converter, count)

    edit_main_readme(root, transform)


def fixture_suite(row):
    if row['suite'] == 'issue':
        return 'office_issues' if row['reference_application'].startswith('LibreOffice') else 'word_issues'
    return row['suite']


def comparison_counts(baseline, candidate):
    differences = [round(value - candidate[case], 2) for case, value in baseline.items()]
    higher = sum(value >= 0.5 for value in differences)
    lower = sum(value <= -0.5 for value in differences)
    return str(higher), str(lower), str(len(differences) - higher - lower)


def update_mini_pdf_breakdown(root, rows, label, kanera_scores):
    benchmark = root / 'benchmarks/word/README.md'
    original = benchmark.read_text(encoding='utf-8')
    before, body, after = section_parts(original, 'benchmark-breakdown')
    note = 'MiniPdf results come from a standalone run.'
    lines = [line for line in body.strip('\n').splitlines() if line != note]
    start = next(i for i, line in enumerate(lines) if line.startswith('|'))
    headers, previous, end = read_table(lines, start)
    groups = defaultdict(list)
    for row in rows:
        groups[fixture_suite(row)].append(row['ink_f1'])
    averages = {(f'{len(groups[suite])} {name}', reference): f'{mean(groups[suite]):.1f}%'
                for suite, name, reference in SUITES if groups[suite]}
    if label in headers:
        column = headers.index(label)
    else:
        column = len(headers)
        headers.append(label)
    for row in previous:
        row[column:column + 1] = [averages.get(tuple(row[:2]), '–')]
        row[2:] = highlight_best(row[2:], higher_is_better=True)

    comparison_start = next(i for i in range(end, len(lines)) if lines[i].startswith('|'))
    comparison_headers, counts, comparison_end = read_table(lines, comparison_start)
    candidate = {row['case']: row['ink_f1'] for row in rows}
    recorded = [row for row in counts if row[0] != label]
    if candidate.keys() == kanera_scores.keys():
        recorded.append((label, *comparison_counts(kanera_scores, candidate)))
    lines[comparison_start:comparison_end] = markdown_table(comparison_headers, recorded).splitlines()
    lines[start:end] = markdown_table(headers, previous).splitlines()
    updated = before + '\n' + '\n'.join(lines).rstrip() + '\n\n' + note + '\n' + after
    if updated != original:
        benchmark.write_text(updated, encoding='utf-8')


def update_accuracy_readme(root, report, run_date):
    count = report['documents']
    corpus = report['by_corpus']
    failures = report['conversion_failures']
    body = (
        f'**Full-corpus ink test.** Last run {run_date} (UTC).\n'
        f'Kanera scored **{report["accuracy_pct"]:.2f}%** across {count:,} documents\n'
        f'({corpus["huggingface"]["documents"]:,} Hugging Face and {corpus["upstream"]["documents"]:,} vendored, '
        f'{report["reference_pages"]:,} reference pages).\n'
        f'Page counts matched on {report["matching_page_counts"]:,} documents. '
        f'{failures:,} {"conversion" if failures == 1 else "conversions"} failed.\n'
        '[Full corpus accuracy](benchmarks/word/README.md#full-corpus-accuracy).'
    )
    edit_main_readme(root, lambda text: replace_section(text, 'full-corpus-accuracy', body))


def update_readmes(root, rows, table, engines, versions, run_date):
    public_columns = ('Converter', 'Mean ink overlap', 'Matching page count', 'Failed conversions',
                      'Median cold time', 'Median warm time', 'Warm p95 latency', 'Warm docs/sec',
                      *MEMORY_COLUMNS, 'Total PDF size')
    results = {engine: dict(zip(table[0], row, strict=True))
               for engine, row in zip(engines, table[1:], strict=True)}
    memory_note = 'Peak RSS is the peak resident memory during a successful cold conversion.'
    libreoffice_peak = results['libreoffice']['Max cold peak RSS']
    # A host Flatpak launch can yield only a few MiB of launcher RSS even for the largest case.
    if libreoffice_peak != '–' and float(libreoffice_peak.split()[0]) < 8:
        results['libreoffice'].update(dict.fromkeys(MEMORY_COLUMNS, '–'))
        memory_note += ' LibreOffice memory is unavailable because GNU time measured its launcher.'
    count = sum(row['engine'] == 'ours' for row in rows)
    ordered_engines = sorted(results, key=lambda key: -float(results[key]['Mean ink overlap'].rstrip('%')))
    headline_columns = ('Converter', *HEADLINE_METRICS)
    headline_names = ['Kanera' if engine == 'ours' else results[engine]['Converter'] for engine in ordered_engines]
    headline_metrics = [metric_cells(results, ordered_engines, column) for column in HEADLINE_METRICS.values()]
    headline = '\n\n'.join([
        '### Benchmark at a glance',
        f'{count} DOCX documents [Full comparison](#benchmark)',
        markdown_table(headline_columns, zip(headline_names, *headline_metrics, strict=True)),
        'Visual match measures how closely text and graphics line up with the reference pages. '
        'Conversion times are medians and include file I/O.',
    ])

    converters = []
    for engine in ordered_engines:
        name = results[engine]['Converter']
        version = versions.get(engine, '').split()
        if version:
            name += ' ' + (version[1] if engine == 'libreoffice' and len(version) > 1 else version[0])
        converters.append(name)
    metrics = [metric_cells(results, ordered_engines, column) for column in public_columns[1:]]
    published = zip(converters, *metrics, strict=True)
    date_line = f'**Last full benchmark run** {run_date} (UTC), {count} documents.'
    overview = '\n\n'.join([
        markdown_table(public_columns, published),
        memory_note,
        'The smallest total PDF size is determined among converters that completed every document.',
    ])

    suites = defaultdict(lambda: defaultdict(list))
    per_case = defaultdict(dict)
    for row in rows:
        suite = fixture_suite(row)
        suites[suite][row['engine']].append(row['ink_f1'])
        per_case[row['engine']][row['case']] = row['ink_f1']

    breakdown = []
    for suite, label, reference in SUITES:
        scores = suites.get(suite)
        if scores:
            averages = [f'{mean(scores[engine]):.1f}%' for engine in engines]
            breakdown.append([f'{len(scores["ours"])} {label}', reference,
                              *highlight_best(averages, higher_is_better=True)])
    columns = ('Cases', 'Reference', *(results[engine]['Converter'] for engine in engines))
    counts = []
    for engine in engines:
        if engine == 'ours':
            continue
        counts.append((results[engine]['Converter'], *comparison_counts(per_case['ours'], per_case[engine])))
    details = '\n\n'.join([
        date_line,
        markdown_table(columns, breakdown),
        'Per-document ink overlap differences under half a percentage point count as ties.',
        markdown_table(('Compared with', 'Kanera higher', 'Kanera lower', 'Ties'), counts),
    ])

    def transform(text):
        text = replace_section(text, 'benchmark-summary', headline)
        return replace_section(text, 'benchmark-results', overview)

    edit_main_readme(root, transform)
    benchmark = root / 'benchmarks/word/README.md'
    bench_text = benchmark.read_text(encoding='utf-8')
    bench_updated = replace_section(bench_text, 'benchmark-breakdown', details)
    if bench_text != bench_updated:
        benchmark.write_text(bench_updated, encoding='utf-8')
