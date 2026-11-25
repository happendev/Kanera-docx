"""Refresh the page-image benchmark's tables in benchmarks/image/README.md and the main README.

Built on the Word benchmark's README helpers (benchmarks/word/readme.py): the same locked edits,
table format, bold rules and headline columns.
"""
from collections import defaultdict
from statistics import mean

from readme import (HEADLINE_METRICS, HIGHER_IS_BETTER, MEMORY_COLUMNS, SUITES, comparison_counts, edit_main_readme,
                    edit_readme, fixture_suite, highlight_best, mark_sampled, markdown_table, number, read_table,
                    replace_section, section_parts)

from pages import DPI, ENGINES, ENGINE_DPI

SUMMARY = 'image-benchmark-summary'
RESULTS = 'image-benchmark-results'
BREAKDOWN = 'image-benchmark-breakdown'
README = 'benchmarks/image/README.md'
COLUMNS = ('Converter', 'Mean ink overlap', 'Matching page count', 'Failed conversions', 'Median cold time',
           'Median warm time', 'Warm p95 latency', 'Warm docs/sec', *MEMORY_COLUMNS, 'Total PNG size')
SIZE = 'Total PNG size'
# Marks the PNG size of an engine that cannot render at the benchmark's resolution.
OTHER_DPI = '‡'
MEMORY_NOTE = 'Peak RSS is the peak resident memory during a successful cold conversion.'
LAUNCHER_NOTE = ' LibreOffice memory is unavailable because GNU time measured its launcher.'
SIZE_NOTE = ('The smallest total PNG size is determined among converters that rendered every document at '
             f'{DPI} dpi.')


def other_dpi_note():
    engines = ', '.join(f'{ENGINES[engine]} at {dpi} dpi' for engine, dpi in ENGINE_DPI.items())
    return (f'{OTHER_DPI} Rendered at a fixed resolution with no setting for {DPI} dpi ({engines}), '
            'so its PNGs are smaller and are not compared for size.')


def metric_cells(results, names, column):
    """The column's cells for names, with the best comparable values in bold (as the Word benchmark does)."""
    values = [results[name][column] for name in names]
    eligible = None
    if column == SIZE:
        eligible = [results[name]['Failed conversions'].startswith('0/') and not value.endswith(OTHER_DPI)
                    for name, value in zip(names, values, strict=True)]
    elif column == 'Warm docs/sec':
        # Warm throughput sums the warm times, so an engine warm-timed on the sample only (†) is not compared.
        eligible = [not value.endswith('†') for value in values]
    return highlight_best(values, higher_is_better=column in HIGHER_IS_BETTER, eligible=eligible)


def metric(header):
    return HEADLINE_METRICS.get(header, header)


def converter_name(engine, values, versions):
    name = values['Converter']
    if engine == 'libreoffice':
        office = versions.get('libreoffice', '').split()
        poppler = versions.get('poppler', '').split()
        return (f'LibreOffice {office[1] if len(office) > 1 else ""} + pdftoppm {poppler[-1] if poppler else ""}'
                .replace('  ', ' ').strip())
    version = versions.get(engine, '').split()
    return f'{name} {version[0]}' if version else name


def measured_results(table, engines):
    """Each engine's results by column, with the memory note. LibreOffice's memory is dropped when GNU
    time measured only a launcher (a host Flatpak started from a container)."""
    results = {engine: dict(zip(table[0], row, strict=True)) for engine, row in zip(engines, table[1:], strict=True)}
    note = MEMORY_NOTE
    office = results.get('libreoffice')
    if office and office['Max cold peak RSS'] != '–' and number(office['Max cold peak RSS']) < 8:
        office.update(dict.fromkeys(MEMORY_COLUMNS, '–'))
        note += LAUNCHER_NOTE
    return results, note


def table_lines(text, section):
    _, body, _ = section_parts(text, section)
    lines = body.strip('\n').splitlines()
    start = next((i for i, line in enumerate(lines) if line.startswith('|')), None)
    return lines, start


def publish(headers, displayed):
    """A markdown table of displayed ({name: {header: value}}), most accurate first, best values in bold."""
    names = sorted(displayed, key=lambda name: -(number(displayed[name][headers[1]]) or 0))
    results = {name: {metric(header): displayed[name][header] for header in headers[1:]} for name in names}
    cells = [metric_cells(results, names, metric(header)) for header in headers[1:]]
    return markdown_table(headers, zip(names, *cells, strict=True))


def refresh_rows(text, section, measured):
    """Replace or add the measured rows ([(name, results by column)]) of a section's table and
    recompute its highlights. Rows match by the first word of the converter name."""
    lines, start = table_lines(text, section)
    headers, rows, end = read_table(lines, start)
    displayed = {row[0]: dict(zip(headers[1:], row[1:], strict=True)) for row in rows}
    for name, values in measured:
        existing = next((label for label in displayed if label.split()[0] == name.split()[0]), None)
        displayed = {name if label == existing else label: row for label, row in displayed.items()}
        displayed[name] = {header: values[metric(header)] for header in headers[1:]}
    intro = '\n'.join(lines[:start]).strip('\n')
    tail = '\n'.join(lines[end:]).strip('\n')
    return replace_section(text, section, '\n\n'.join(part for part in (intro, publish(headers, displayed), tail)
                                                      if part))


def results_notes(results, memory_note):
    notes = [memory_note, SIZE_NOTE]
    if any(values[SIZE].endswith(OTHER_DPI) for values in results.values()):
        notes.append(other_dpi_note())
    return notes


def breakdown(rows, results, engines):
    suites = defaultdict(lambda: defaultdict(list))
    per_case = defaultdict(dict)
    for row in rows:
        suites[fixture_suite(row)][row['engine']].append(row['ink_f1'])
        per_case[row['engine']][row['case']] = row['ink_f1']
    table = []
    for suite, label, reference in SUITES:
        scores = suites.get(suite)
        if scores:
            averages = [f'{mean(scores[engine]):.1f}%' for engine in engines]
            table.append([f'{len(scores["ours"])} {label}', reference, *highlight_best(averages, higher_is_better=True)])
    columns = ('Cases', 'Reference', *(results[engine]['Converter'] for engine in engines))
    counts = [(results[engine]['Converter'], *comparison_counts(per_case['ours'], per_case[engine]))
              for engine in engines if engine != 'ours']
    parts = [f'{len(per_case["ours"])} documents.', markdown_table(columns, table)]
    if counts:
        parts.extend([
            'Per-document ink overlap differences under half a percentage point count as ties.',
            markdown_table(('Compared with', 'Kanera higher', 'Kanera lower', 'Ties'), counts),
        ])
    return '\n\n'.join(parts)


def update_readmes(root, rows, table, engines, versions, *, complete):
    """Write the summary in the main README and the results in benchmarks/image/README.md.

    A complete comparison (every engine) rewrites both tables and the suite breakdown. A run of some
    engines replaces their rows and keeps the others, unless no table has been written yet.
    """
    results, memory_note = measured_results(table, engines)
    headline = [('Kanera' if engine == 'ours' else results[engine]['Converter'], results[engine]) for engine in engines]
    full = [(converter_name(engine, results[engine], versions), results[engine]) for engine in engines]

    def write(text, section, headers, measured, notes=()):
        if complete or table_lines(text, section)[1] is None:
            displayed = {name: {header: values[metric(header)] for header in headers[1:]} for name, values in measured}
            return replace_section(text, section, '\n\n'.join([publish(headers, displayed), *notes]))
        return refresh_rows(text, section, measured)

    def summary(text):
        return write(text, SUMMARY, ('Converter', *HEADLINE_METRICS), headline)

    def benchmark(text):
        text = write(text, RESULTS, COLUMNS, full, results_notes(results, memory_note))
        if complete or table_lines(text, BREAKDOWN)[1] is None:
            text = replace_section(text, BREAKDOWN, breakdown(rows, results, engines))
        return mark_sampled(text, RESULTS)

    edit_main_readme(root, summary)
    edit_readme(root, root / README, benchmark)
