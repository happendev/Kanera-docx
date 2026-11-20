"""Convert DOCX files with the Rust engines that ship Python bindings (dxpdf and rdocx).

Usage: runner.py <dxpdf|rdocx> <input.docx> <output.pdf>
       runner.py <dxpdf|rdocx> --batch <jobs.json>
       runner.py --version
"""
from importlib.metadata import version
import json
from pathlib import Path
import sys
from time import perf_counter

ENGINES = ('dxpdf', 'rdocx')


def converter(engine):
    if engine == 'dxpdf':
        import dxpdf

        # Default settings: Skia rendering with fontconfig font lookup.
        return lambda source, target: dxpdf.convert_file(str(source), str(target))

    import rdocx

    def convert(source, target):
        # rdocx falls back to bundled fonts unless system fonts are enabled; the other local
        # converters in the benchmark use the host's fonts, so opt in for a like-for-like run.
        document = rdocx.Document.open(str(source))
        Path(target).write_bytes(document.to_pdf(system_fonts=True))

    return convert


def main(args):
    if args == ['--version']:
        print(json.dumps({engine: version(engine) for engine in ENGINES}))
        return 0
    if len(args) != 3 or args[0] not in ENGINES:
        print(__doc__.strip(), file=sys.stderr)
        return 2
    engine, source, target = args
    convert = converter(engine)
    if source == '--batch':
        times = []
        for job_source, job_target in json.loads(Path(target).read_text()):
            start = perf_counter()
            convert(job_source, job_target)
            times.append((perf_counter() - start) * 1000)
        print(json.dumps(times))
    else:
        convert(source, target)
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
