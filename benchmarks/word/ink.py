from itertools import zip_longest
import subprocess

import numpy as np
from PIL import Image, ImageFilter

DPI = 96
INK_THRESHOLD = 216
INK_TOLERANCE_PX = 1


def render(pdf, raster_dir=None):
    raster_dir = raster_dir if raster_dir is not None else pdf.parent / 'raster'
    raster_dir.mkdir(exist_ok=True)
    prefix = raster_dir / pdf.stem
    for old in prefix.parent.glob(prefix.name + '-*.png'):
        old.unlink()
    subprocess.run(['pdftoppm', '-r', str(DPI), '-png', str(pdf), str(prefix)], check=True)
    pages = prefix.parent.glob(prefix.name + '-*.png')
    return sorted(pages, key=lambda p: int(p.stem.rsplit('-', 1)[1]))


def render_reference(directory, fixture, raster_dir):
    if 'pdf' in fixture:
        return render(directory / fixture['pdf'], raster_dir)

    raster_dir.mkdir(exist_ok=True)
    scale = DPI / fixture['reference_dpi']
    pages = []
    for number, page in enumerate(fixture['reference_images'], 1):
        path = raster_dir / f'reference-{number}.png'
        with Image.open(directory / page['path']) as image:
            size = (round(image.width * scale), round(image.height * scale))
            image.convert('RGB').resize(size, Image.Resampling.LANCZOS).save(path)
        pages.append(path)
    return pages


def ink_mask(path, shape):
    with Image.open(path) as source:
        im = source.convert('L')
    canvas = Image.new('L', shape, 255)
    canvas.paste(im, (0, 0))
    return canvas.point(lambda value: 255 if value < INK_THRESHOLD else 0)


def ink_score(reference, candidate):
    # Compare page-coordinate ink masks in both directions with a one-pixel tolerance.
    if reference is None or candidate is None:
        return 0
    with Image.open(reference) as a, Image.open(candidate) as b:
        shape = (max(a.width, b.width), max(a.height, b.height))
    word = ink_mask(reference, shape)
    pdf = ink_mask(candidate, shape)
    word_ink = np.asarray(word) > 0
    pdf_ink = np.asarray(pdf) > 0
    if not word_ink.any() or not pdf_ink.any():
        return float(not word_ink.any() and not pdf_ink.any())
    neighborhood = ImageFilter.MaxFilter(2 * INK_TOLERANCE_PX + 1)
    word_nearby = np.asarray(word.filter(neighborhood)) > 0
    pdf_nearby = np.asarray(pdf.filter(neighborhood)) > 0
    recall = np.count_nonzero(word_ink & pdf_nearby) / np.count_nonzero(word_ink)
    precision = np.count_nonzero(pdf_ink & word_nearby) / np.count_nonzero(pdf_ink)
    return float(2 * recall * precision / (recall + precision)) if recall + precision else 0


def page_scores(reference, candidate):
    return [ink_score(word, output) for word, output in zip_longest(reference, candidate)]
