"""Refresh the benchmark tables in the repository READMEs.

The main README gets compact summary tables only, plus a glance table derived from them. The full
results go in the benchmark READMEs under benchmarks/.
"""

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
    'Warm docs/sec': 'Warm docs/sec',
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


def sampled_note():
    from convert import FULL_TIMING_MS, SAMPLE_SIZE

    return (f'† Timed on a fixed {SAMPLE_SIZE}-document sample only, because timing every document would '
            f'take more than {FULL_TIMING_MS // 60_000} minutes; the other documents were converted untimed '
            'and count for accuracy and failures.')


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
    eligible = None
    if column == 'Total PDF size':
        eligible = [results[engine]['Failed conversions'].startswith('0/') for engine in engines]
    elif column == 'Warm docs/sec':
        # Warm throughput sums the warm times, so an engine warm-timed on the sample only (†) is not compared.
        eligible = [not value.endswith('†') for value in values]
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


def mark_sampled(text, name):
    """Keep the † note in a README section exactly when one of its table cells has a †."""
    _, body, _ = section_parts(text, name)
    note = sampled_note()
    parts = [part for part in body.strip('\n').split('\n\n') if part != note]
    if any('†' in line for line in body.splitlines() if line.startswith('|')):
        parts.append(note)
    return replace_section(text, name, '\n\n'.join(parts))


def edit_readme(root, path, transform):
    """Apply transform to a README under a lock, so benchmark runs that finish together keep each other's edits."""
    output = root / 'out'
    output.mkdir(exist_ok=True)
    with (output / '.readme.lock').open('a') as lock:
        flock(lock, LOCK_EX)
        original = path.read_text(encoding='utf-8')
        updated = transform(original)
        if updated != original:
            with NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent, delete=False) as temporary:
                staged = Path(temporary.name)
                try:
                    temporary.write(updated)
                    temporary.close()
                    copymode(path, staged)
                    staged.replace(path)
                finally:
                    staged.unlink(missing_ok=True)


def edit_main_readme(root, transform):
    """Edit the main README, then refresh its glance table from the summary tables."""
    edit_readme(root, root / 'README.md', lambda text: refresh_glance(transform(text)))


def edit_benchmark_readme(root, transform):
    edit_readme(root, root / 'benchmarks/word/README.md', transform)


# The glance table at the top of the main README: one row per summary table, Kanera against the
# most accurate converter that isn't Kanera or DocFoundry (the same engine).
GLANCE = (
    ('Word to PDF', 'benchmark-summary'),
    ('Word to page images', 'image-benchmark-summary'),
    ('Word to Markdown', 'markdown-benchmark-summary'),
    ('Word to text', 'text-benchmark-summary'),
)
SAME_ENGINE = ('Kanera', 'DocFoundry')


def number(value):
    value = value.replace('†', '').split()
    return float(value[0].split('/')[0].rstrip('%')) if value and value[0] != '–' else None


def compare_cells(ours, theirs, *, higher_is_better):
    """Our value against theirs, the better one in bold."""
    a, b = number(ours), number(theirs) if theirs else None
    if a is None or b is None:
        return ours, theirs or '–'
    if a == b:
        return ours, theirs
    if (a > b) == higher_is_better:
        return f'**{ours}**', theirs
    return ours, f'**{theirs}**'


def refresh_glance(text):
    if '<!-- benchmark-glance:start -->' not in text:
        return text
    rows = []
    for label, section in GLANCE:
        if f'<!-- {section}:start -->' not in text:
            continue
        _, body, _ = section_parts(text, section)
        lines = body.splitlines()
        start = next((i for i, line in enumerate(lines) if line.startswith('|')), None)
        if start is None:
            continue
        headers, table, _ = read_table(lines, start)
        ours = next((row for row in table if row[0].startswith('Kanera')), None)
        if ours is None:
            continue
        others = [row for row in table if not row[0].startswith(SAME_ENGINE) and number(row[1]) is not None]
        best = max(others, key=lambda row: number(row[1]), default=None)
        warm, memory = headers.index('Reused process'), headers.index('Median peak memory')
        rate = next(i for i, header in enumerate(headers) if header.endswith('/sec'))
        accuracy, _ = compare_cells(ours[1], best[1] if best else None, higher_is_better=True)
        other = f'{best[1]} {best[0]}' if best else '–'
        cells = [compare_cells(ours[i], best[i] if best else None, higher_is_better=i == rate)
                 for i in (warm, rate, memory)]
        rows.append((label, accuracy, other, *(f'{a} vs {b}' if best else a for a, b in cells)))
    table = markdown_table(('Conversion', 'Accuracy', 'Best other converter', 'Time per file', 'Files per second',
                            'Peak memory'), rows)
    return replace_section(text, 'benchmark-glance', table)


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


def update_measured_readme(root, table, versions, count, engines):
    measured = []
    for engine, row in zip(engines, table[1:], strict=True):
        values = dict(zip(table[0], row, strict=True))
        name = values['Converter']
        converter = name
        if versions.get(engine):
            version = versions[engine].split()
            converter += ' ' + (version[1] if engine == 'libreoffice' and len(version) > 1 else version[0])
        measured.append((values, 'Kanera' if engine == 'ours' else name, converter))

    def summary(text):
        for values, name, _ in measured:
            text = refresh_engine_table(text, 'benchmark-summary', values, name, count)
        return text

    def results(text):
        for values, _, converter in measured:
            text = refresh_engine_table(text, 'benchmark-results', values, converter, count)
        return mark_sampled(text, 'benchmark-results')

    # The main README has the compact summary; the benchmark README has every column.
    edit_main_readme(root, summary)
    edit_benchmark_readme(root, results)


def update_suite_scores(root, rows, engines):
    edit_benchmark_readme(root, lambda text: suite_scores(text, rows, engines))


def suite_scores(original, rows, engines):
    before, body, after = section_parts(original, 'benchmark-breakdown')
    lines = body.splitlines()
    start = next(i for i, line in enumerate(lines) if line.startswith('|'))
    headers, previous, end = read_table(lines, start)
    groups = defaultdict(list)
    for row in rows:
        groups[row['engine'], fixture_suite(row)].append(row['ink_f1'])
    for row in previous:
        suite = next(key for key, _, reference in SUITES if reference == row[1])
        for column, label in enumerate(headers[2:], 2):
            if label in engines:
                scores = groups[engines[label], suite]
                if scores:
                    row[column] = f'{mean(scores):.1f}%'
        row[2:] = highlight_best(row[2:], higher_is_better=True)
    lines[start:end] = markdown_table(headers, previous).splitlines()
    return before + '\n' + '\n'.join(lines).strip('\n') + '\n' + after


def fixture_suite(row):
    if row['suite'] == 'issue':
        return 'office_issues' if row['reference_application'].startswith('LibreOffice') else 'word_issues'
    return row['suite']


def comparison_counts(baseline, candidate):
    # Documents a sampled engine skipped are left out of its comparison.
    differences = [round(value - candidate[case], 2) for case, value in baseline.items() if case in candidate]
    higher = sum(value >= 0.5 for value in differences)
    lower = sum(value <= -0.5 for value in differences)
    return str(higher), str(lower), str(len(differences) - higher - lower)


def update_mini_pdf_breakdown(root, rows, label, kanera_scores):
    edit_benchmark_readme(root, lambda text: mini_pdf_breakdown(text, rows, label, kanera_scores))


def mini_pdf_breakdown(original, rows, label, kanera_scores):
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
    return before + '\n' + '\n'.join(lines).rstrip() + '\n\n' + note + '\n' + after


def update_accuracy_readme(root, report):
    count = report['documents']
    failures = report['conversion_failures']
    body = (
        f'Kanera scores **{report["accuracy_pct"]:.2f}%** on the full corpus of {count:,} documents '
        f'({report["reference_pages"]:,} reference pages), with matching page counts on '
        f'{report["matching_page_counts"]:,} and {failures:,} failed {"conversion" if failures == 1 else "conversions"}.'
    )
    edit_main_readme(root, lambda text: replace_section(text, 'full-corpus-accuracy', body))


def update_readmes(root, rows, table, engines, versions):
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
    headline = markdown_table(headline_columns, zip(headline_names, *headline_metrics, strict=True))

    converters = []
    for engine in ordered_engines:
        name = results[engine]['Converter']
        version = versions.get(engine, '').split()
        if version:
            name += ' ' + (version[1] if engine == 'libreoffice' and len(version) > 1 else version[0])
        converters.append(name)
    metrics = [metric_cells(results, ordered_engines, column) for column in public_columns[1:]]
    published = zip(converters, *metrics, strict=True)
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
        f'{count} documents.',
        markdown_table(columns, breakdown),
        'The head-to-head counts use the full comparison. Per-document ink overlap differences '
        'under half a percentage point count as ties.',
        markdown_table(('Compared with', 'Kanera higher', 'Kanera lower', 'Ties'), counts),
    ])

    def benchmark_transform(text):
        text = replace_section(text, 'benchmark-results', overview)
        text = replace_section(text, 'benchmark-breakdown', details)
        return mark_sampled(text, 'benchmark-results')

    edit_main_readme(root, lambda text: replace_section(text, 'benchmark-summary', headline))
    edit_benchmark_readme(root, benchmark_transform)
