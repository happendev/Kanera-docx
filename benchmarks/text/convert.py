"""Convert the Word benchmark's fixtures to plain text with each engine, timing cold and warm runs.

This is the Markdown benchmark's harness with --format text: benchmarks/markdown/convert.py takes
the same options. Run score.py afterwards.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'markdown'))
import convert  # noqa: E402

if __name__ == '__main__':
    convert.main(['--format', 'text', *sys.argv[1:]])
