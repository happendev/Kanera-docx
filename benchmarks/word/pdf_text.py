"""Extract page words from a PDF in a child process.

Some converters write PDFs that crash PyMuPDF's text extraction outright (a segmentation fault in
MuPDF), which would take a scoring worker down with it. Running the extraction here confines a
crash to one PDF, which the scorer then counts as a failed conversion.

Usage: pdf_text.py <file.pdf>  ->  JSON {"pages": [[[token, x0, y0], ...], ...], "text": "..."}
"""
import json
import subprocess
import sys
import unicodedata
from pathlib import Path

import pymupdf


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


def extract(pdf):
    """Pages of (token, x0, y0) and the lower-cased text, or None when extraction crashes or hangs."""
    try:
        result = subprocess.run([sys.executable, '-I', str(Path(__file__).resolve()), str(pdf)],
                                capture_output=True, text=True, timeout=300)
    except subprocess.TimeoutExpired:
        return None
    if result.returncode:
        return None
    data = json.loads(result.stdout)
    return [[tuple(word) for word in page] for page in data['pages']], data['text']


if __name__ == '__main__':
    with pymupdf.open(sys.argv[1]) as document:
        text = ''.join(page.get_text() for page in document).lower()
        print(json.dumps({'pages': words(document), 'text': text}))
