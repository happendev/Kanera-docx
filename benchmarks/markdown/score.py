"""Score the Markdown of each engine and write the report.

Run convert.py first. Each engine's Markdown is rendered with one CommonMark renderer (GFM tables,
strikethrough, footnotes and $ math, with GitHub's rules) and scored on what a viewer would show,
so syntax the renderer does not understand counts against the engine.

References, none of them produced by any engine under test:
- Words, reading order, bold and italic: Word's own PDF of the document, read with PyMuPDF.
  Documents without a Word PDF use the text of the DOCX for words and order.
- Structure: a small reader of the DOCX XML in this file: headings by outline level, numbered and
  bulleted paragraphs, tables of two or more columns, external hyperlinks and pictures. Text boxes,
  table-of-contents entries and the Title and Subtitle styles are neutral.
- Neutral content: words of headers, footers, text boxes, charts and equations are not counted
  against an engine when Word's PDF does not show them, nor is their emphasis. An engine that leaves
  out chart or equation words, which Word draws, is not penalised either.
"""
from argparse import ArgumentParser
import csv
import json
import re
import statistics
import sys
import unicodedata
import zipfile
from collections import Counter
from datetime import datetime, timedelta
from urllib.parse import unquote
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import pymupdf as fitz
from lxml import etree, html as lhtml
from markdown_it import MarkdownIt
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.footnote import footnote_plugin
from rapidfuzz.distance import Indel

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT / 'benchmarks' / 'word'))
from fetch import BENCHMARK_FIXTURES, OUTPUT as WORD_OUTPUT, UPSTREAM_DIR  # noqa: E402

OUTPUT = ROOT / 'out' / 'benchmark-markdown'
MARKDOWN = OUTPUT / 'md'
CASES = WORD_OUTPUT / 'cases'
ENGINE_NAMES = {
    'kanera': '@kanera/docx2pdf', 'aspose': 'Aspose.Words (.NET)', 'syncfusion': 'Syncfusion DocIO (.NET)',
    'pandoc': 'Pandoc', 'docling': 'Docling', 'markitdown': 'MarkItDown', 'mammoth': 'mammoth + turndown',
    'libreoffice': 'LibreOffice',
}

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MATH = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'w': W, 'r': R}
q = lambda n: '{%s}%s' % (W, n)
STRICT = [(b'http://purl.oclc.org/ooxml/wordprocessingml/main', W.encode()),
          (b'http://purl.oclc.org/ooxml/officeDocument/relationships', R.encode()),
          (b'http://purl.oclc.org/ooxml/drawingml/main', b'http://schemas.openxmlformats.org/drawingml/2006/main'),
          (b'http://purl.oclc.org/ooxml/officeDocument/math', b'http://schemas.openxmlformats.org/officeDocument/2006/math')]

CJK = '\u3000-\u9fff\uac00-\ud7ff\uf900-\ufaff\uff00-\uffef'
TOKEN = re.compile('[' + CJK + ']|[^\\W_' + CJK + ']+')
FOLD = str.maketrans({'\u06be': '\u0647', '\u06cc': '\u064a', '\u06a9': '\u0643', '\u0649': '\u064a', '\u04d9': '\u0259'})
CJKNUM = '\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e\u5343\u58f9\u8cb3\u8d30\u53c3\u53c1\u8086\u4f0d\u9678\u9646\u67d2\u634c\u7396\u62fe\u96f6\u3007'
LABEL = re.compile(r'^(?:[(\[]?(?:\d+(?:\.\d+)*|[ivxlcdm]+|[a-z]|[' + CJKNUM + r']+)[.:)\]\u3001]\s*|\d+,\s*|\d+(?:\.\d+)*\s+|\u7b2c[\d' + CJKNUM + r']+[\u6761\u689d\u7ae0\u7bc0\u6b3e\u9805\u8282]\s*|[\u2022\u00b7\u25aa\u25cf\u25e6\u2013\u2014\u27a2\u2713\u2714\u25a0\u25a1\u2610\u2612-]\s+)+', re.I)


def tokens(text):
    text = unicodedata.normalize('NFKC', text).lower().replace('\u00ad', '').translate(FOLD)
    return TOKEN.findall(text)


def norm(text):
    return ' '.join(tokens(text))


def norm_label(text):
    """Normalised text without a leading list or heading number such as "1.2", "a)" or "[3]".
    The label is matched after NFKC, which makes full-width punctuation ASCII."""
    text = unicodedata.normalize('NFKC', text).strip()
    stripped = LABEL.sub('', text)
    return norm(stripped if norm(stripped) else text)


# ---------------------------------------------------------------- reference structure from DOCX

def url_key(url):
    """A URL compared without percent-encoding or a trailing slash before the end or the anchor."""
    url = unquote(url.strip())
    return re.sub(r'/+(?=#|$)', '', url)


def bind_data(z, names, doc):
    """Shows the custom XML value bound to a content control (w:dataBinding), as Word does."""
    stores = {}
    for name in names:
        m = re.match(r'customXml/itemProps(\d+)\.xml$', name)
        if m and f'customXml/item{m.group(1)}.xml' in names:
            props = etree.fromstring(z.read(name))
            item_id = props.get('{http://schemas.openxmlformats.org/officeDocument/2006/customXml}itemID', '')
            stores[item_id.upper()] = etree.fromstring(z.read(f'customXml/item{m.group(1)}.xml'))
    if not stores:
        return
    for sdt in doc.iter(q('sdt')):
        binding = sdt.find('w:sdtPr/w:dataBinding', NS)
        content = sdt.find('w:sdtContent', NS)
        if binding is None or content is None:
            continue
        store = stores.get((binding.get(q('storeItemID')) or '').upper())
        if store is None:
            continue
        ns = {k: v for k, _, v in re.findall(r'''xmlns:(\w+)=(['"])(.*?)\2''', binding.get(q('prefixMappings')) or '')}
        try:
            found = store.xpath(binding.get(q('xpath')), namespaces=ns)
        except etree.XPathError:
            continue
        if not found:
            continue
        value = found[0] if isinstance(found[0], str) else ''.join(found[0].itertext())
        texts = list(content.iter(q('t')))
        if texts:
            texts[0].text = value
            for t in texts[1:]:
                t.text = ''


def chart_words(z, names, body, folder, rel_target):
    """Words of the charts in the body (title, series names, categories, values). Word draws a
    chart rather than printing its data, so an engine that writes the data out as text or a table
    is neither rewarded nor penalised for these words."""
    words = []
    for chart in body.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}chart'):
        target = rel_target.get(chart.get('{%s}id' % R), (None,))[0]
        if not target:
            continue
        part = (folder + '/' + target).replace('//', '/')
        if part not in names:
            continue
        root = etree.fromstring(z.read(part))
        for el in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}v', '{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
            text = el.text or ''
            words += tokens(text)
            try:
                value = float(text)
            except ValueError:
                continue
            # A value may print with its number format applied: percentages, rounding or a date.
            words += tokens(f'{value * 100:g} {value:,.2f} {round(value)} {value:,.0f}')
            if 1 <= value < 2958466:
                day = datetime(1899, 12, 30) + timedelta(days=value)
                words += tokens(day.strftime('%b %B %y %Y %m %d ') + str(day.day) + ' ' + str(day.month))
    return words


def docx_oracle(path):
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    def read(n):
        if n not in names:
            return None
        data = z.read(n)
        for strict, normal in STRICT:
            data = data.replace(strict, normal)
        return etree.fromstring(data)
    main = 'word/document.xml'
    root_rels = read('_rels/.rels')
    if root_rels is not None:
        for r in root_rels:
            if r.get('Type', '').endswith('/officeDocument'):
                main = r.get('Target').lstrip('/')
    folder, _, base = main.rpartition('/')
    doc = read(main)
    bind_data(z, names, doc)
    rels = read(f'{folder}/_rels/{base}.rels')
    targets = {}
    if rels is not None:
        for r in rels:
            targets[r.get('Type', '').rsplit('/', 1)[-1]] = (folder + '/' + r.get('Target')).replace('//', '/') if not r.get('Target', '').startswith('/') else r.get('Target').lstrip('/')
    styles = read(targets.get('styles', 'word/styles.xml'))
    numbering = read(targets.get('numbering', 'word/numbering.xml'))

    style = {}
    if styles is not None:
        for s in styles.iter(q('style')):
            sid = s.get(q('styleId'))
            name = s.find('w:name', NS)
            based = s.find('w:basedOn', NS)
            outline = s.find('w:pPr/w:outlineLvl', NS)
            num = s.find('w:pPr/w:numPr/w:numId', NS)
            style[sid] = {
                'name': (name.get(q('val')) if name is not None else sid or '').lower(),
                'based': based.get(q('val')) if based is not None else None,
                'outline': int(outline.get(q('val'))) if outline is not None else None,
                'num': num.get(q('val')) if num is not None else None,
            }

    def chain(sid):
        seen = 0
        while sid in style and seen < 20:
            yield style[sid]
            sid = style[sid]['based']
            seen += 1

    abstract_fmt = {}
    num_abs = {}
    if numbering is not None:
        for a in numbering.iter(q('abstractNum')):
            fmts = {}
            for lvl in a.iter(q('lvl')):
                f = lvl.find('w:numFmt', NS)
                t = lvl.find('w:lvlText', NS)
                fmts[lvl.get(q('ilvl'))] = (f.get(q('val')) if f is not None else 'decimal', t.get(q('val')) if t is not None else '')
            abstract_fmt[a.get(q('abstractNumId'))] = fmts
        for n in numbering.iter(q('num')):
            a = n.find('w:abstractNumId', NS)
            if a is not None:
                num_abs[n.get(q('numId'))] = a.get(q('val'))

    links = set()
    rel_target = {}
    if rels is not None:
        for r in rels:
            rel_target[r.get('Id')] = (r.get('Target'), r.get('TargetMode'))

    def text_of(p):
        out = []
        for el in p.iter(q('t'), q('tab'), q('br'), q('cr'), q('noBreakHyphen')):
            if any(a.tag in (q('del'), q('moveFrom')) for a in el.iterancestors()):
                continue
            run = next((a for a in el.iterancestors() if a.tag == q('r')), None)
            if run is not None and run.find('w:rPr/w:vanish', NS) is not None and run.find('w:rPr/w:vanish', NS).get(q('val')) not in ('0', 'false'):
                continue
            if el.tag == q('t'):
                out.append(el.text or '')
            elif el.tag == q('noBreakHyphen'):
                out.append('-')
            else:
                out.append(' ')
        return ''.join(out)

    headings, neutral_headings, list_items, tables = [], set(), [], []
    images = 0
    body = doc.find('w:body', NS)
    source = []
    neutral_table = []
    boxes = []
    # Headers and footers are neutral: their words, and the page margins they print in (points).
    furniture = []
    for name in names:
        if re.match(r'word/(header|footer)\d*\.xml$', name):
            part = read(name)
            for para in part.iter(q('p')):
                # Each paragraph's own text: a text box's paragraphs are read on their own.
                own = [t.text or '' for t in para.iter(q('t'))
                       if next(a for a in t.iterancestors() if a.tag == q('p')) is para]
                furniture += tokens(''.join(own))
    tops, bottoms = [], []
    for pg in body.iter(q('pgMar')):
        tops.append(abs(int(float(pg.get(q('top'), '1440')))) / 20)
        bottoms.append(abs(int(float(pg.get(q('bottom'), '1440')))) / 20)
    margins = (min(tops, default=0), min(bottoms, default=0))
    for p in body.iter(q('p')):
        anc = list(p.iterancestors())
        tags = {a.tag for a in anc}
        if q('del') in tags or any(a.tag.endswith('}Fallback') for a in anc):
            continue
        if q('txbxContent') in tags:
            boxes += tokens(text_of(p))
            continue
        in_table = q('tc') in tags
        text = text_of(p)
        source.append(text)
        if not norm(text):
            continue
        ppr = p.find('w:pPr', NS)
        sid = ppr.find('w:pStyle', NS).get(q('val')) if ppr is not None and ppr.find('w:pStyle', NS) is not None else None
        sty = list(chain(sid)) if sid else list(chain(next((k for k, v in style.items() if v['name'] == 'normal'), None)))
        sname = sty[0]['name'] if sty else ''
        outline = None
        d = ppr.find('w:outlineLvl', NS) if ppr is not None else None
        if d is not None:
            outline = int(d.get(q('val')))
        else:
            for s in sty:
                if s['outline'] is not None:
                    outline = s['outline']
                    break
        m = re.match(r'heading (\d)$', sname)
        if outline is None and m:
            outline = int(m.group(1)) - 1
        if in_table:
            continue
        if sname.startswith('toc') or sname in ('title', 'subtitle'):
            neutral_headings.add(norm_label(text))
            continue
        if outline is not None and outline < 9:
            headings.append((outline + 1, norm_label(text)))
            continue
        num_id = None
        d = ppr.find('w:numPr/w:numId', NS) if ppr is not None else None
        ilvl = ppr.find('w:numPr/w:ilvl', NS) if ppr is not None else None
        if d is not None:
            num_id = d.get(q('val'))
        else:
            num_id = next((s['num'] for s in sty if s['num']), None)
        if num_id and num_id != '0' and num_id in num_abs:
            fmt, lvl_text = abstract_fmt.get(num_abs[num_id], {}).get(ilvl.get(q('val')) if ilvl is not None else '0', ('decimal', '%1'))
            if fmt != 'none' and lvl_text:
                list_items.append(norm_label(text))

    for tbl in body.iter(q('tbl')):
        anc = {a.tag for a in tbl.iterancestors()}
        if q('txbxContent') in anc and q('tbl') not in anc:
            # Text boxes are neutral, tables in them included.
            neutral_table += [t for p in tbl.iter(q('p')) for t in tokens(norm_label(text_of(p)))]
            continue
        if q('tbl') in anc or q('txbxContent') in anc:
            continue
        rows = tbl.findall('w:tr', NS)
        cols = max((len(r.findall('w:tc', NS)) for r in rows), default=0)
        cells = []
        for r in rows:
            for c in r.findall('w:tc', NS):
                # Each paragraph without the list label Word may show before it.
                t = ' '.join(norm_label(text_of(p)) for p in c.iter(q('p')))
                if t.strip():
                    cells.append(norm(t))
        if cols < 2:
            # A single column is as much a box around paragraphs as a table: neutral either way.
            neutral_table += [t for c in cells for t in c.split()]
        elif cells:
            tables.append(cells)

    for h in body.iter(q('hyperlink')):
        rid = h.get('{%s}id' % R)
        # Only links Word shows: an empty hyperlink element has nothing to click.
        if rid and rid in rel_target and rel_target[rid][1] == 'External' and norm(text_of(h)):
            anchor = h.get(q('anchor'))
            links.add(url_key(rel_target[rid][0] + ('#' + anchor if anchor else '')))
    for t in body.iter(q('instrText')):
        m = re.search(r'HYPERLINK\s+"([^"]+)"(?:.*?\\l\s+"([^"]+)")?', t.text or '')
        if m and not m.group(1).startswith('#'):
            links.add(url_key(m.group(1) + ('#' + m.group(2) if m.group(2) else '')))
    for el in body.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip', '{urn:schemas-microsoft-com:vml}imagedata'):
        if any(a.tag.endswith('}Fallback') for a in el.iterancestors()):
            continue
        images += 1
    # Equation words: Word draws an equation, and engines write it as LaTeX, MathML, a picture or text.
    math = []
    for eq in body.iter(MATH + 'oMath'):
        if any(a.tag == MATH + 'oMath' or a.tag.endswith('}Fallback') or a.tag == q('del') for a in eq.iterancestors()):
            continue
        parts = [t.text or '' for t in eq.iter(MATH + 't')]
        math += tokens(''.join(parts)) + [t for part in parts for t in tokens(part)]
    return {'headings': headings, 'neutral': neutral_headings, 'lists': list_items, 'tables': tables,
            'links': links, 'images': images, 'text': '\n'.join(source),
            'neutral_table': neutral_table, 'chart': chart_words(z, names, body, folder, rel_target),
            'furniture': furniture, 'boxes': boxes, 'math': math, 'margins': margins}


# ---------------------------------------------------------------- reference words from Word's PDF

QUAD = re.compile(r'<char [^>]*?quad="([^"]+)"')


def slanted(xml, count):
    """For each text line of a page's structured-text XML, whether each character is slanted: Word's
    italic for a font without an italic face is a skewed upright glyph. None when the lines do not line up."""
    lines = re.findall(r'<line [^>]*>(.*?)</line>', xml, re.S)
    if len(lines) != count:
        return [None] * count
    out = []
    for line in lines:
        flags = []
        for quad in QUAD.findall(line):
            ulx, uly, _, _, llx, lly, _, _ = map(float, quad.split())
            flags.append(lly - uly > 0 and (ulx - llx) / (lly - uly) > 0.1)
        out.append(flags)
    return out


def pdf_reference(pdf, margins=(0, 0), furniture=()):
    """Words of Word's PDF, with those Word drew mostly in a bold or italic face. Lines in the top
    and bottom page margins (headers, footers, page numbers) are left out, and so are the lines at the
    top and bottom of a page made only of header and footer words: a tall header pushes the body down."""
    words, bold, italic = [], [], []
    top, bottom = margins
    furniture = set(furniture)
    pages = []
    for page in fitz.open(pdf):
        # Collecting styles marks text drawn both filled and stroked: Word's bold for a font without a bold face.
        textpage = page.get_textpage(flags=fitz.TEXTFLAGS_DICT | fitz.TEXT_COLLECT_STYLES)
        lines = [line for block in page.get_text('dict', textpage=textpage)['blocks'] for line in block.get('lines', [])]
        for line, slant in zip(lines, slanted(textpage.extractXML(), len(lines))):
            line['slant'] = slant
        pages.append((page.rect.height, lines))
    def key(line):
        return ' '.join(tokens(''.join(span['text'] for span in line['spans']))), round(line['bbox'][1])
    # A running header or footer line repeats at the same height on other pages, if there are any.
    repeated = Counter(k for _, lines in pages for k in {key(line) for line in lines})
    least = 2 if len(pages) > 1 else 1
    for height, lines in pages:
        skip = set()
        for edge in (sorted(lines, key=lambda l: l['bbox'][1]), sorted(lines, key=lambda l: -l['bbox'][3])):
            for line in edge:
                y0, y1 = line['bbox'][1], line['bbox'][3]
                if y1 <= top + 1 or y0 >= height - bottom - 1:
                    continue
                if abs((y0 + y1) / 2 - height / 2) < height / 5 or repeated[key(line)] < least:
                    break
                if not furniture.issuperset(key(line)[0].split()):
                    break
                skip.add(id(line))
        for line in lines:
            y0, y1 = line['bbox'][1], line['bbox'][3]
            if y1 <= top + 1 or y0 >= height - bottom - 1 or id(line) in skip:
                continue
            # Words are read from the whole line: a change of face inside a word does not split it.
            text, styles = '', []
            slant = line.get('slant')
            for span in line['spans']:
                font = span['font'].lower()
                b = bool(span['flags'] & 16 or span.get('char_flags', 0) & 48 == 48
                         or any(k in font for k in ('bold', 'black', 'heavy', 'semibold', 'demi')))
                it = bool(span['flags'] & 2 or 'italic' in font or 'oblique' in font)
                styles += [(b, it or bool(slant and slant[len(text) + k])) for k in range(len(span['text']))]
                text += span['text']
            for m in TOKEN.finditer(text):
                t = tokens(m.group())
                words += t
                chars = styles[m.start():m.end()]
                if sum(c[0] for c in chars) * 2 > len(chars):
                    bold += t
                if sum(c[1] for c in chars) * 2 > len(chars):
                    italic += t
    return words, bold, italic


# ---------------------------------------------------------------- rendered Markdown

# GitHub's and Pandoc's rules for $ math: no space inside the delimiters, no digit after the closing one.
MD = (MarkdownIt('commonmark', {'html': True}).enable(['table', 'strikethrough']).use(footnote_plugin)
      .use(dollarmath_plugin, allow_space=False, allow_digits=False))
# markdown-it accepts only GIF, PNG, JPEG and WebP data URIs; a picture in another format is still a
# picture (a viewer shows it or a broken-image icon), not its base64 text.
MD.validateLink = lambda url: not re.match(r'(vbscript|javascript|file):|data:(?!image/)', url.strip(), re.I)


# Underscores count only against a letter or digit: a run of them is a form's blank line.
LEAK_INLINE = re.compile(r'\*\*|(?<!\w)__(?=[^\W_])|(?<=[^\W_])__(?!\w)|(?<![\w*])\*(?=\w)|\[\^?\w+\]\(|```')
LEAK_BLOCK = re.compile(r'^#{1,6} |^\|.*\|$|^\s*[-+] \S|^\s*\d+[.)] \S|^> ', re.M)


def syntax(text, paragraphs=None):
    """Markdown syntax in text a viewer shows: emphasis markers and link syntax anywhere, and
    heading, list, quote and table-row markers starting a plain paragraph."""
    found = Counter(m.group(0)[:3] for m in LEAK_INLINE.finditer(text))
    found.update(m.group(0).strip()[:3] for m in LEAK_BLOCK.finditer(text if paragraphs is None else paragraphs))
    return found


def li_text(li):
    """A list item's own text: no nested lists and no footnote markers."""
    out = []
    for t in li.xpath('.//text()[not(ancestor::sup[@class="footnote-ref"])]'):
        el = t.getparent()
        if t.is_tail:
            el = el.getparent()
        while el is not None and el is not li and el.tag not in ('ul', 'ol'):
            el = el.getparent()
        if el is li:
            out.append(t)
    return ''.join(out)


def rendered(md_text):
    raw_html = MD.render(md_text)
    root = lhtml.fromstring('<div>' + raw_html + '</div>')
    for br in root.iter('br'):
        br.tail = '\n' + (br.tail or '')
    for cell in root.iter('td', 'th', 'tr'):
        cell.tail = '\n' + (cell.tail or '')
    for el in root.xpath('//*[contains(concat(" ", @class, " "), " math ")]'):
        el.drop_tree()
    def plain(el):
        # Text without footnote markers, which Word shows as superscripts after the heading text.
        return ''.join(el.xpath('.//text()[not(ancestor::sup[@class="footnote-ref"])]'))
    heads = [(int(h.tag[1]), norm_label(plain(h))) for h in root.iter('h1', 'h2', 'h3', 'h4', 'h5', 'h6')]
    lis = []
    for li in root.iter('li'):
        if li.getparent() is not None and li.getparent().get('class') and 'footnote' in li.getparent().get('class'):
            continue
        if li.xpath('ancestor::section[contains(@class,"footnotes")]'):
            continue
        lis.append(norm_label(li_text(li)))
    def cell_text(c):
        # The cell's own text, line by line without a leading list label; nested tables count as cells of their own.
        parts = []
        for t in c.xpath('.//text()'):
            node = t.getparent()
            if t.is_tail:
                node = node.getparent()
            while node is not None and node.tag not in ('td', 'th'):
                node = node.getparent()
            if node is c:
                parts.append(t)
        text = ''.join(parts)
        return ' '.join(norm_label(line) for line in text.split('\n'))
    cells = [norm(cell_text(c)) for c in root.iter('td', 'th')]
    cells = [c for c in cells if c]
    blocks = ('p', 'li', 'td', 'th', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6')
    def emph(tags):
        # Emphasised words, each with the text of its block so that emphasis in a block the reference
        # calls a heading can be left out, whether the engine wrote a heading or a bold paragraph.
        out = []
        for el in root.iter(*tags):
            if any(a.tag in tags for a in el.iterancestors()):
                continue
            block = next((a for a in el.iterancestors() if a.tag in blocks), None)
            if block is not None and block.tag[0] == 'h':
                continue
            key = '' if block is None else norm_label(li_text(block) if block.tag == 'li' else plain(block))
            out += [(key, t) for t in tokens(el.text_content())]
        return out
    bold = emph(('strong', 'b'))
    italic = emph(('em', 'i'))
    hrefs = {url_key(a.get('href')) for a in root.iter('a') if a.get('href') and not a.get('href').startswith('#')}
    imgs = len(list(root.iter('img')))
    text = root.text_content()
    paragraphs = '\n'.join(p.text_content() for p in root.iter('p')
                           if not p.xpath('ancestor::li|ancestor::td|ancestor::th|ancestor::section'))
    leaks = syntax(text, paragraphs)
    return {'text': text, 'heads': heads, 'lis': lis, 'cells': cells, 'bold': bold, 'italic': italic,
            'hrefs': hrefs, 'imgs': imgs, 'leaks': leaks}


def raw_noise(md_text):
    data_uri = re.compile(r'data:[\w/+.-]+;base64,[A-Za-z0-9+/=\s]+')
    base64 = sum(len(m) for m in data_uri.findall(md_text))
    rest = data_uri.sub('', md_text)
    return {
        'md_bytes': len(rest.encode()),
        'base64_bytes': base64,
        'html_tags': len(re.findall(r'</?[a-zA-Z][\w-]*(?:\s[^<>]*)?/?>', rest)),
    }


def overlap(a, b):
    return sum((Counter(a) & Counter(b)).values())


def f1(ref, out, common):
    if not ref and not out:
        return None
    return 2 * common / (ref + out)


def reference_pdf(fixture):
    if fixture['corpus'] == 'huggingface':
        return CASES / fixture['case'] / 'word.pdf'
    if fixture.get('pdf') and 'Word' in fixture.get('reference_application', ''):
        return UPSTREAM_DIR / fixture['pdf']
    return None


def score_case(job):
    fixture, engines, failed = job
    case = fixture['case']
    oracle = docx_oracle(CASES / case / 'input.docx')
    pdf = reference_pdf(fixture)
    if pdf:
        words, bold, italic = pdf_reference(pdf, oracle['margins'], oracle['furniture'])
    else:
        words, bold, italic = tokens(oracle['text']), None, None
    source_syntax = syntax(oracle['text'])
    heading_texts = {h for _, h in oracle['headings']} | oracle['neutral']
    heading_tokens = Counter(t for h in [h for _, h in oracle['headings']] + list(oracle['neutral']) for t in h.split())
    rows = []
    for engine in engines:
        path = MARKDOWN / engine / f'{case}.md'
        ok = path.exists() and (engine, case) not in failed
        md_text = path.read_text(encoding='utf-8', errors='replace') if ok else ''
        r = rendered(md_text)
        out = tokens(r['text'])
        row = {'case': case, 'engine': engine, 'converted': ok, 'reference': 'word-pdf' if pdf else 'docx-text'}
        row.update(raw_noise(md_text))
        row['leaked_syntax'] = sum((r['leaks'] - source_syntax).values())
        # Words of charts, equations, headers, footers and text boxes beyond those in the reference are
        # neutral, and so are chart and equation words an engine leaves out.
        ref_count, out_count = Counter(words), Counter(out)
        extra = sum(((out_count - ref_count) & Counter(oracle['chart'] + oracle['math'] + oracle['furniture'] + oracle['boxes'])).values())
        missing = sum(((ref_count - out_count) & Counter(oracle['chart'] + oracle['math'])).values())
        row['ref_words'], row['out_words'], row['common_words'] = len(words) - missing, len(out) - extra, overlap(words, out)
        row['order'] = Indel.normalized_similarity(words[:20000], out[:20000]) if words or out else None
        if bold is not None:
            # Emphasis inside headings and title paragraphs is not compared: the heading is the emphasis.
            # Word's PDF has no blocks, so its heading words are taken out by count.
            furniture = Counter(oracle['furniture'] + oracle['boxes'] + oracle['chart'])
            for kind, ref_words in (('bold', bold), ('italic', italic)):
                ref_emph = list((Counter(ref_words) - heading_tokens).elements())
                out_emph = Counter(t for block, t in r[kind] if block not in heading_texts)
                out_emph = list((out_emph - ((out_emph - Counter(ref_emph)) & furniture)).elements())
                row[f'ref_{kind}'], row[f'out_{kind}'], row[f'common_{kind}'] = len(ref_emph), len(out_emph), overlap(ref_emph, out_emph)
        # A heading counts when its text and level match; headings in the neutral styles are ignored.
        ref_heads = Counter((l, h) for l, h in oracle['headings'] if h)
        shown = [(l, h) for l, h in r['heads'] if h and not (h in oracle['neutral'] and h not in {x for _, x in ref_heads})]
        # The levels may all shift by one amount: an engine may start a document's headings at ##.
        shift = max(range(-3, 4), key=lambda d: (sum((ref_heads & Counter((l + d, h) for l, h in shown)).values()), -abs(d)))
        out_heads = Counter((l + shift, h) for l, h in shown)
        row['ref_headings'], row['out_headings'], row['common_headings'] = sum(ref_heads.values()), sum(out_heads.values()), sum((ref_heads & out_heads).values())
        ref_lis = Counter(x for x in oracle['lists'] if x)
        out_lis = Counter(x for x in r['lis'] if x and not (x in oracle['neutral'] and x not in ref_lis))
        row['ref_list_items'], row['out_list_items'], row['common_list_items'] = sum(ref_lis.values()), sum(out_lis.values()), sum((ref_lis & out_lis).values())
        ref_cells = [t for cells in oracle['tables'] for c in cells for t in c.split()]
        out_cells = [t for c in r['cells'] for t in c.split()]
        extra = Counter(out_cells) - Counter(ref_cells)
        neutral = sum((extra & Counter(oracle['neutral_table'] + oracle['chart'])).values())
        row['ref_table_words'], row['out_table_words'], row['common_table_words'] = len(ref_cells), len(out_cells) - neutral, overlap(ref_cells, out_cells)
        row['ref_links'], row['common_links'] = len(oracle['links']), len(oracle['links'] & r['hrefs'])
        row['ref_images'], row['out_images'] = oracle['images'], r['imgs']
        parts = {
            'text': f1(row['ref_words'], row['out_words'], row['common_words']),
            'order': row['order'],
            'headings': f1(row['ref_headings'], row['out_headings'], row['common_headings']),
            'lists': f1(row['ref_list_items'], row['out_list_items'], row['common_list_items']),
            'tables': f1(row['ref_table_words'], row['out_table_words'], row['common_table_words']),
            'links': row['common_links'] / row['ref_links'] if row['ref_links'] else None,
            'images': min(row['ref_images'], row['out_images']) / max(row['ref_images'], row['out_images']) if row['ref_images'] or row['out_images'] else None,
        }
        if bold is not None:
            parts['bold'] = f1(row['ref_bold'], row['out_bold'], row['common_bold'])
            parts['italic'] = f1(row['ref_italic'], row['out_italic'], row['common_italic'])
        scores = [v for v in parts.values() if v is not None]
        row['accuracy'] = (statistics.mean(scores) if scores else 1.0) if ok else 0.0
        rows.append(row)
    return rows


def summarize(rows, conversion, engines):
    by_engine = {e: [r for r in rows if r['engine'] == e] for e in engines}
    timing = {}
    for row in conversion['results']:
        timing.setdefault(row['engine'], []).append(row)
    summary = {}
    for engine, rs in by_engine.items():
        pdf = [r for r in rs if r['reference'] == 'word-pdf']
        def micro(group, name):
            ref = sum(r.get(f'ref_{name}', 0) for r in group)
            out = sum(r.get(f'out_{name}', 0) for r in group)
            common = sum(r.get(f'common_{name}', 0) for r in group)
            return f1(ref, out, common) or 0.0
        t = timing.get(engine, [])
        cold = sorted(r['cold_ms'] for r in t if 'cold_ms' in r)
        warm = sorted(r['warm_ms'] for r in t if 'warm_ms' in r)
        rss = sorted(r['peak_rss_kib'] for r in t if r.get('peak_rss_kib'))
        summary[engine] = {
            'documents': len(rs),
            'failed': sum(not r['converted'] for r in rs),
            'accuracy': statistics.mean(r['accuracy'] for r in rs),
            'text': micro(rs, 'words'),
            'order': statistics.mean(r['order'] or 0 for r in rs),
            'headings': micro(rs, 'headings'),
            'lists': micro(rs, 'list_items'),
            'tables': micro(rs, 'table_words'),
            'bold': micro(pdf, 'bold'),
            'italic': micro(pdf, 'italic'),
            'links': sum(r['common_links'] for r in rs) / max(sum(r['ref_links'] for r in rs), 1),
            'images': sum(min(r['ref_images'], r['out_images']) for r in rs) / max(sum(max(r['ref_images'], r['out_images']) for r in rs), 1),
            'leaked_syntax': sum(r['leaked_syntax'] for r in rs),
            'html_tags': sum(r['html_tags'] for r in rs),
            'base64_mib': sum(r['base64_bytes'] for r in rs) / 2**20,
            'md_mib': sum(r['md_bytes'] for r in rs) / 2**20,
            'cold_median_ms': statistics.median(cold) if cold else None,
            'warm_median_ms': statistics.median(warm) if warm else None,
            'warm_p95_ms': warm[min(len(warm) - 1, int(0.95 * len(warm)))] if warm else None,
            'warm_docs_per_s': len(warm) / (sum(warm) / 1000) if warm else None,
            'rss_median_mib': statistics.median(rss) / 1024 if rss else None,
            'rss_max_mib': rss[-1] / 1024 if rss else None,
            'version': conversion['versions'].get(engine, engine),
        }
    return summary


def tables(summary):
    engines = sorted(summary, key=lambda e: -summary[e]['accuracy'])
    best = {}
    higher = ['accuracy', 'text', 'order', 'headings', 'lists', 'tables', 'bold', 'italic', 'links', 'images', 'warm_docs_per_s']
    lower = ['failed', 'leaked_syntax', 'cold_median_ms', 'warm_median_ms', 'warm_p95_ms', 'rss_median_mib', 'rss_max_mib']
    for k in higher + lower:
        values = [summary[e][k] for e in engines if summary[e][k] is not None]
        if values:
            best[k] = max(values) if k in higher else min(values)
    def cell(e, k, fmt):
        v = summary[e][k]
        if v is None:
            return '–'
        text = fmt(v)
        # A tie is a tie as shown.
        return f'**{text}**' if k in best and text == fmt(best[k]) else text
    pct = lambda v: f'{v:.1%}'
    n = summary[engines[0]]['documents']
    def name(e):
        # The converter and its version number; mammoth's turndown version is named too.
        numbers = re.findall(r'\d+(?:\.\d+)+', summary[e]['version'])
        if e == 'mammoth' and len(numbers) > 1:
            return f'mammoth {numbers[0]} + turndown {numbers[1]}'
        return f"{ENGINE_NAMES.get(e, e)} {numbers[0] if numbers else ''}".strip()
    accuracy = ['| Converter | Markdown accuracy | Text | Order | Headings | Lists | Tables | Bold | Italic | Links | Images | Leaked syntax | Failed |',
                '| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |']
    speed = ['| Converter | Median cold time | Median warm time | Warm p95 latency | Warm docs/sec | Median cold peak RSS | Max cold peak RSS |',
             '| --- | ---: | ---: | ---: | ---: | ---: | ---: |']
    for e in engines:
        accuracy.append(f"| {name(e)} | " + ' | '.join([cell(e, 'accuracy', pct), *(cell(e, k, pct) for k in ('text', 'order', 'headings', 'lists', 'tables', 'bold', 'italic', 'links', 'images')),
                                                          cell(e, 'leaked_syntax', str), cell(e, 'failed', lambda v: f'{v}/{n}')]) + ' |')
        ms = lambda v: f'{v:.1f} ms'
        mib = lambda v: f'{v:.1f} MiB'
        speed.append(f"| {name(e)} | " + ' | '.join([cell(e, 'cold_median_ms', ms), cell(e, 'warm_median_ms', ms), cell(e, 'warm_p95_ms', ms),
                                                       cell(e, 'warm_docs_per_s', lambda v: f'{v:.1f}'), cell(e, 'rss_median_mib', mib), cell(e, 'rss_max_mib', mib)]) + ' |')
    if 'libreoffice' in summary:
        speed.append('\nLibreOffice memory is unavailable because GNU time measured its launcher.')
    return '\n'.join(accuracy), '\n'.join(speed)


def replace_block(path, marker, content):
    text = path.read_text(encoding='utf-8')
    start, end = f'<!-- {marker}:start -->', f'<!-- {marker}:end -->'
    if start not in text:
        return
    before, rest = text.split(start, 1)
    _, after = rest.split(end, 1)
    path.write_text(f'{before}{start}\n{content}\n{end}{after}', encoding='utf-8')


def main():
    parser = ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument('--engines', help='comma-separated engines to score (default: every engine converted)')
    parser.add_argument('--no-readme', action='store_true', help='do not update the README tables')
    args = parser.parse_args()
    conversion = json.loads((OUTPUT / 'conversion.json').read_text())
    engines = args.engines.split(',') if args.engines else list(dict.fromkeys(r['engine'] for r in conversion['results']))
    failed = {(r['engine'], r['case']) for r in conversion['results'] if 'error' in r}
    with ProcessPoolExecutor() as pool:
        rows = [row for rows in pool.map(score_case, [(f, engines, failed) for f in BENCHMARK_FIXTURES]) for row in rows]
    with (OUTPUT / 'results.csv').open('w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(dict.fromkeys(k for r in rows for k in r)))
        writer.writeheader()
        writer.writerows(rows)
    summary = summarize(rows, conversion, engines)
    (OUTPUT / 'summary.json').write_text(json.dumps(summary, indent=1))
    accuracy, speed = tables(summary)
    report = f"{accuracy}\n\n{speed}\n"
    (OUTPUT / 'report.md').write_text(report, encoding='utf-8')
    print(report)
    if not args.no_readme and len(engines) > 1:
        replace_block(HERE / 'README.md', 'markdown-accuracy', accuracy)
        replace_block(HERE / 'README.md', 'markdown-speed', speed)
        replace_block(ROOT / 'README.md', 'markdown-benchmark', f'{accuracy}\n\n{speed}')


if __name__ == '__main__':
    main()
