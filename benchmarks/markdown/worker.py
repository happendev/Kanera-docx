"""Converts DOCX to Markdown, or to plain text for a .txt output, with a Python engine or Pandoc.

worker.py <markitdown|docling|pandoc> <input.docx> <output.md|output.txt>
worker.py <markitdown|docling|pandoc> --batch <jobs.json>   (prints one time in ms per job, or null)
MarkItDown writes only Markdown.
"""
import json
import subprocess
import sys
from pathlib import Path
from time import perf_counter

engine, first, second = sys.argv[1:4]
if engine == 'markitdown':
    from markitdown import MarkItDown
    converter = MarkItDown()
    def render(path, text):
        if text:
            raise ValueError('MarkItDown writes only Markdown')
        return converter.convert(str(path)).text_content
elif engine == 'docling':
    from docling.document_converter import DocumentConverter
    converter = DocumentConverter()
    def render(path, text):
        document = converter.convert(str(path)).document
        return document.export_to_text() if text else document.export_to_markdown()
elif engine == 'pandoc':
    import pypandoc
    PANDOC = pypandoc.get_pandoc_path()
    # Pandoc has no resident mode: every conversion is a process.
    def render(path, text):
        return subprocess.run([PANDOC, '-f', 'docx', '-t', 'plain' if text else 'gfm', str(path)], check=True,
                              capture_output=True, text=True, timeout=300).stdout
else:
    sys.exit(__doc__)


def convert(source, target):
    Path(target).write_text(render(source, target.endswith('.txt')), encoding='utf-8')


if first == '--batch':
    times = []
    for source, target in json.loads(Path(second).read_text()):
        start = perf_counter()
        try:
            convert(source, target)
            times.append((perf_counter() - start) * 1000)
        except Exception:
            times.append(None)
    print(json.dumps(times))
else:
    convert(first, second)
