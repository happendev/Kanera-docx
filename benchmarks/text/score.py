"""Score the plain text of each engine and write the report.

This is the Markdown benchmark's scorer with --format text: benchmarks/markdown/score.py takes the
same options.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'markdown'))
import score  # noqa: E402

if __name__ == '__main__':
    score.main(['--format', 'text', *sys.argv[1:]])
