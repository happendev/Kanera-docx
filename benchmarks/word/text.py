"""Score text extraction against the text of Word's reference PDFs.

Extracts every fixture with the built package (`docxToText` with headers and footers), then compares
the words with those PyMuPDF reads from Word's PDF of the same document. Words are compared as a
bag, so reading order does not count, and only fixtures with a reference PDF are scored.

Recall is the share of the reference's words that the extracted text has; precision is the share of
the extracted words that the reference has. Some differences are expected: a reference repeats its
headers on every page and shows the page numbers of that page, PDF text splits words at hyphenation
points and superscripts, and some fixtures were exported with comments and deletions in view.

Usage: text.py [--worst N]
"""
from argparse import ArgumentParser
from collections import Counter
from concurrent.futures import ProcessPoolExecutor
import json
import re
import subprocess
import unicodedata

from fetch import BENCHMARK_FIXTURES, HERE, OUTPUT, benchmark_reference, verify, verify_reference
from pdf_text import extract as extract_pdf_text

CASES = OUTPUT / 'cases'
TEXT = OUTPUT / 'text'
ROOT = HERE.parents[1]

EXTRACT = """
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { docxToText } from %s;
const [cases, out, fixtures] = process.argv.slice(1);
mkdirSync(out, { recursive: true });
const times = [];
for (const name of JSON.parse(fixtures)) {
  const data = readFileSync(`${cases}/${name}/input.docx`);
  const start = performance.now();
  const text = await docxToText(data, { includeHeadersFooters: true });
  times.push(performance.now() - start);
  writeFileSync(`${out}/${name}.txt`, text);
}
times.sort((a, b) => a - b);
console.log(JSON.stringify({ count: times.length, median: times[times.length >> 1] }));
"""

CJK = '\u3000-\u9fff\uac00-\ud7ff\uf900-\ufaff\uff00-\uffef'
TOKEN = re.compile('[' + CJK + ']|[^\\W_' + CJK + ']+')
# Letters that Arabic-script and Cyrillic fonts map to a different code point in their PDFs.
FOLD = str.maketrans({'\u06be': '\u0647', '\u06cc': '\u064a', '\u06a9': '\u0643', '\u0649': '\u064a', '\u04d9': '\u0259'})


def tokens(text):
    """Lower-case words, with each CJK character a word of its own."""
    text = unicodedata.normalize('NFKC', text).lower().replace('\u00ad', '').translate(FOLD)
    return TOKEN.findall(text)


def score(case):
    name, reference = case
    extracted = extract_pdf_text(reference)
    if extracted is None:
        return None
    ref = Counter(tokens(extracted[1]))
    ours = Counter(tokens((TEXT / f'{name}.txt').read_text(encoding='utf-8')))
    return {
        'case': name,
        'reference': sum(ref.values()),
        'extracted': sum(ours.values()),
        'common': sum((ref & ours).values()),
        'missing': [t for t, _ in (ref - ours).most_common(8)],
        'extra': [t for t, _ in (ours - ref).most_common(8)],
    }


def main():
    parser = ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--worst', type=int, default=15, help='list the N fixtures with the lowest scores')
    args = parser.parse_args()
    names = []
    references = {}
    for fixture in BENCHMARK_FIXTURES:
        name = fixture['case']
        verify(CASES / name / 'input.docx', fixture['docx_sha256'])
        names.append(name)
        if 'pdf_sha256' in fixture and fixture['reference_application'].startswith('Microsoft Word'):
            directory, reference = benchmark_reference(fixture)
            verify_reference(directory, reference)
            references[name] = directory / reference['pdf']
    module = json.dumps((ROOT / 'dist' / 'node.js').as_uri())
    run = subprocess.run(['node', '--input-type=module', '-e', EXTRACT % module,
                          str(CASES), str(TEXT), json.dumps(sorted(names))],
                         check=True, capture_output=True, text=True)
    timing = json.loads(run.stdout)
    with ProcessPoolExecutor() as pool:
        rows = [r for r in pool.map(score, sorted(references.items())) if r]
    ref = sum(r['reference'] for r in rows)
    ext = sum(r['extracted'] for r in rows)
    common = sum(r['common'] for r in rows)
    print(f"{timing['count']} documents extracted, median {timing['median']:.1f} ms each")
    print(f'{len(rows)} documents with a Word PDF: recall {common / ref:.2%}, precision {common / ext:.2%}')
    def f1(r):
        total = r['reference'] + r['extracted']
        return 2 * r['common'] / total if total else 1.0
    for r in sorted(rows, key=f1)[:args.worst]:
        print(f"{f1(r):6.1%}  {r['case'][:44]:44}  missing {' '.join(r['missing'])[:60]:60}  extra {' '.join(r['extra'])[:60]}")
    (OUTPUT / 'text-accuracy.json').write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding='utf-8')


if __name__ == '__main__':
    main()
