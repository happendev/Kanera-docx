"""Render DOCX pages to PNG with rdocx, the one Rust engine of the Word benchmark that draws page images.

Usage: rdocx_runner.py <input.docx> <output dir> <dpi>
       rdocx_runner.py --batch <jobs.json> <dpi>
       rdocx_runner.py --version

Each job writes <output dir>/page-<n>.png. A batch prints the time of each job in milliseconds.
"""
from importlib.metadata import version
import json
from pathlib import Path
import sys
from time import perf_counter


def render(source, target, dpi):
    import rdocx

    # Like the Word benchmark's rdocx runner: system fonts on, because rdocx otherwise uses only its
    # bundled fallback fonts and the other converters use the host's fonts.
    pages = rdocx.Document.open(str(source)).render_all_pages(dpi, system_fonts=True)
    target = Path(target)
    target.mkdir(parents=True, exist_ok=True)
    for number, png in enumerate(pages, 1):
        (target / f'page-{number}.png').write_bytes(png)


def main(args):
    if args == ['--version']:
        print(json.dumps({'rdocx': version('rdocx')}))
        return 0
    if len(args) != 3:
        print(__doc__.strip(), file=sys.stderr)
        return 2
    source, target, dpi = args[0], args[1], float(args[2])
    if source == '--batch':
        times = []
        for job_source, job_target in json.loads(Path(target).read_text()):
            start = perf_counter()
            render(job_source, job_target, dpi)
            times.append((perf_counter() - start) * 1000)
        print(json.dumps(times))
    else:
        render(source, target, dpi)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
