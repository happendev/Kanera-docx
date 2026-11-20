"""Refresh the benchmark summaries in the repository READMEs."""

from collections import defaultdict
from statistics import mean


HIGHER_IS_BETTER = {'Mean ink overlap', 'Matching page count', 'Warm docs/sec'}


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


def replace_section(text, name, body):
    start = f'<!-- {name}:start -->'
    end = f'<!-- {name}:end -->'
    before, _, rest = text.partition(start)
    _, _, after = rest.partition(end)
    if text.count(start) != 1 or text.count(end) != 1 or end not in rest:
        raise ValueError(f'Expected exactly one {name} section in the README')
    return before + start + '\n' + body + '\n' + end + after


def update_readmes(root, rows, table, engines, versions, run_date):
    public_columns = ('Converter', 'Mean ink overlap', 'Matching page count', 'Failed conversions',
                      'Median cold time', 'Median warm time', 'Warm p95 latency', 'Warm docs/sec',
                      'Total PDF size')
    results = {engine: dict(zip(table[0], row, strict=True))
               for engine, row in zip(engines, table[1:], strict=True)}
    count = sum(row['engine'] == 'ours' for row in rows)
    ordered_engines = sorted(results, key=lambda key: -float(results[key]['Mean ink overlap'].rstrip('%')))
    headline_columns = ('Converter', 'Visual match', 'Page counts matched', 'Fresh process', 'Reused process')
    headline_names = ['Kanera' if engine == 'ours' else results[engine]['Converter'] for engine in ordered_engines]
    headline_metrics = [metric_cells(results, ordered_engines, column) for column in (
        'Mean ink overlap', 'Matching page count', 'Median cold time', 'Median warm time',
    )]
    headline = '\n\n'.join([
        '### Benchmark at a glance',
        f'{count} DOCX documents · Last full run {run_date} (UTC) · [Full comparison](#benchmark)',
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
        date_line,
        markdown_table(public_columns, published),
        'The smallest total PDF size is determined among converters that completed every document.',
    ])

    suites = defaultdict(lambda: defaultdict(list))
    per_case = defaultdict(dict)
    for row in rows:
        suite = row['suite']
        if suite == 'issue':
            suite = 'office_issues' if row['reference_application'].startswith('LibreOffice') else 'word_issues'
        suites[suite][row['engine']].append(row['ink_f1'])
        per_case[row['engine']][row['case']] = row['ink_f1']

    breakdown = []
    for suite, label, reference in (
        ('classic', 'classic', 'Microsoft 365 Word page images'),
        ('sample', 'Hugging Face', 'Word for Mac PDFs'),
        ('word_issues', 'issues', 'Microsoft 365 Word PDFs'),
        ('office_issues', 'issues', 'LibreOffice PDFs'),
    ):
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
        differences = [round(value - per_case[engine][case], 2) for case, value in per_case['ours'].items()]
        better = sum(value >= 0.5 for value in differences)
        worse = sum(value <= -0.5 for value in differences)
        counts.append((results[engine]['Converter'], str(better), str(worse), str(count - better - worse)))
    details = '\n\n'.join([
        date_line,
        markdown_table(columns, breakdown),
        'Per-document ink overlap differences under half a percentage point count as ties.',
        markdown_table(('Compared with', 'Kanera higher', 'Kanera lower', 'Ties'), counts),
    ])

    main = root / 'README.md'
    benchmark = root / 'benchmarks/word/README.md'
    main_text = main.read_text(encoding='utf-8')
    bench_text = benchmark.read_text(encoding='utf-8')
    main_updated = replace_section(main_text, 'benchmark-summary', headline)
    main_updated = replace_section(main_updated, 'benchmark-results', overview)
    bench_updated = replace_section(bench_text, 'benchmark-breakdown', details)
    if main_text != main_updated:
        main.write_text(main_updated, encoding='utf-8')
    if bench_text != bench_updated:
        benchmark.write_text(bench_updated, encoding='utf-8')
