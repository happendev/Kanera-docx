import { copyFile, mkdtemp, readFile, rm, stat, utimes } from 'node:fs/promises';
import { basename, join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { convertDocx as convertBrowser } from '../src/browser.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { ChartPointBudget } from '../src/ooxml/chart.js';
import { parseDocx } from '../src/docx/parser.js';
import { FontBudget, SfntFont } from '../src/fonts/sfnt.js';
import { loadCjkFont, fontVersion } from '../src/fonts/node-cjk.js';
import { subsetTrueType } from '../src/fonts/subset.js';
import { BundledFontData, FontRegistry } from '../src/fonts/registry.js';
import { convertDocx } from '../src/node.js';
import { nodeCompression } from '../src/util/node-compression.js';
import { makeDocx, p, zipStore } from './helpers/docx.js';
import { layout, lines } from './helpers/layout.js';
import { pdfText } from './helpers/pdf.js';
import { inlineImage, redPng } from './helpers/sample.js';

vi.mock('node:fs/promises', async (original) => {
  const fs = await original<typeof import('node:fs/promises')>();
  return { ...fs, readFile: vi.fn(fs.readFile) };
});

afterEach(() => vi.restoreAllMocks());

const pageField = '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>';
const ref = '<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>';
const right = '<w:jc w:val="right"/>';
const section = (width: number) => '<w:sectPr><w:pgSz w:w="' + width * 20 + '" w:h="12000"/><w:pgMar w:left="1000" w:right="1000"/></w:sectPr>';

describe('conversion caches', () => {
  it('keeps adjacent runs distinct while reusing their inherited formatting', () => {
    const styles = '<w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Flip"><w:rPr><w:b/><w:i/></w:rPr></w:style>';
    const run = (text: string, format = '') => '<w:r>' + (format ? '<w:rPr>' + format + '</w:rPr>' : '') + '<w:t>' + text + '</w:t></w:r>';
    const body = '<w:p><w:pPr><w:pStyle w:val="Base"/></w:pPr>' +
      run('A') + run('B') + run('C', '<w:b w:val="0"/>') + run('D', '<w:rStyle w:val="Flip"/>') + run('E') + '</w:p>';
    const paragraph = parseDocx(makeDocx({ body, styles }), nodeCompression.inflateRaw).body[0];
    if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph');
    const items = paragraph.content.filter((c) => c.type === 'text');
    expect(items.map((c) => [c.text, c.props.bold, c.props.italic])).toEqual([
      ['A', true, false], ['B', true, false], ['C', false, false], ['D', false, true], ['E', true, false],
    ]);
    expect(items[0].props).not.toBe(items[1].props);
  });

  it('resolves reused Word styles against each document’s theme', async () => {
    const resolve = async (accent: string, slot = 'accent1') => {
      const styles = `<w:style w:type="paragraph" w:styleId="Heading"><w:rPr><w:color w:themeColor="${slot}"/></w:rPr></w:style>`;
      const theme = `<a:clrScheme><a:accent1><a:srgbClr val="${accent}"/></a:accent1><a:accent2><a:srgbClr val="112233"/></a:accent2></a:clrScheme>`;
      const doc = makeDocx({ styles, theme, body: p('Title', '<w:pStyle w:val="Heading"/>') });
      return lines((await layout(doc)).pages[0].body)[0].ops[0].color;
    };
    expect(await resolve('FF0000')).toBe('FF0000');
    expect(await resolve('00FF00')).toBe('00FF00');
    expect(await resolve('0000FF')).toBe('0000FF');
    expect(await resolve('0000FF', 'accent2')).toBe('112233');
  });

  it('lays out the same separator and note at each section width', async () => {
    const footnotes = '<w:footnote w:type="separator" w:id="-1">' + p('Rule', right) + '</w:footnote><w:footnote w:id="1">' + p('Note', right) + '</w:footnote>';
    const { pages } = await layout(makeDocx({ body: ref + p('', section(700)) + ref, sectPr: section(400), footnotes }));
    expect(pages).toHaveLength(2);
    for (const label of ['Rule', 'Note']) {
      const [wide, narrow] = pages.map((page) => lines(page.body).find((line) => line.text === label)!.x);
      expect(wide - narrow).toBeCloseTo(300);
      expect(narrow).toBeLessThan(350);
    }
  });

  it('uses the owning section grid for inherited headers and footnote separators', async () => {
    const grid = (pitch: number, refs = '') => '<w:sectPr>' + refs + '<w:docGrid w:type="lines" w:linePitch="' + pitch + '"/></w:sectPr>';
    const single = '<w:spacing w:line="240" w:lineRule="auto"/>';
    const content = p('First', single) + p('Second', single);
    const footnotes = '<w:footnote w:id="1">' + p('Note') + '</w:footnote><w:footnote w:type="separator" w:id="-1">' + content + '</w:footnote>';
    const { pages } = await layout(makeDocx({ body: ref + p('', grid(400, '<w:headerReference w:type="default" r:id="h"/>')) + ref, sectPr: grid(800), footnotes, headers: { h: content } }));
    expect(pages).toHaveLength(2);
    for (const layer of ['body', 'headerFooter'] as const) {
      const spacing = pages.map((page) => {
        const text = lines(page[layer]);
        return text.find((line) => line.text === 'Second')!.y - text.find((line) => line.text === 'First')!.y;
      });
      expect(spacing[1] - spacing[0]).toBeCloseTo(20);
    }
  });

  it('resolves page fields independently in a reused separator and note body', async () => {
    const fieldLine = (label: string) => '<w:p><w:r><w:t>' + label + '</w:t></w:r>' + pageField + '</w:p>';
    const footnotes = '<w:footnote w:type="separator" w:id="-1">' + fieldLine('Rule') + '</w:footnote><w:footnote w:id="1">' + fieldLine('Note') + '</w:footnote>';
    const { pages, warnings } = await layout(makeDocx({ body: ref + p('', section(400)) + ref, sectPr: section(400), footnotes }));
    expect(warnings).toEqual([]);
    expect(pages.map((page) => lines(page.body).slice(-2).map((line) => line.text))).toEqual([['Rule1', 'Note1'], ['Rule2', 'Note2']]);
  });

  it('shapes a static header and footer once across pages and pagination passes', async () => {
    const glyph = vi.spyOn(FontRegistry.prototype, 'faceForChar');
    const sectPr = '<w:sectPr><w:headerReference w:type="default" r:id="h"/><w:footerReference w:type="default" r:id="f"/></w:sectPr>';
    const total = '<w:p>' + pageField.replace('PAGE', 'NUMPAGES') + '</w:p>';
    const { pages } = await layout(makeDocx({ body: total + p('body', '<w:pageBreakBefore/>').repeat(4), sectPr, headers: { h: p('Z') }, footers: { f: p('Q') } }));
    expect(pages).toHaveLength(5);
    for (const page of pages) expect(lines(page.headerFooter).map((line) => line.text)).toEqual(['Z', 'Q']);
    for (const char of ['Z', 'Q']) expect(glyph.mock.calls.filter((args) => args[3] === char.codePointAt(0))).toHaveLength(1);
  });

  it('updates nested header fields on every page while reusing a static footer at mirrored margins', async () => {
    const sectPr = '<w:sectPr><w:headerReference w:type="default" r:id="h"/><w:footerReference w:type="default" r:id="f"/><w:pgMar w:left="1000" w:right="2000"/></w:sectPr>';
    const table = '<w:tbl><w:tr><w:tc><w:p>' + pageField + '</w:p></w:tc></w:tr></w:tbl>';
    const { pages } = await layout(makeDocx({ body: p('body') + p('body', '<w:pageBreakBefore/>').repeat(2), sectPr, settings: '<w:mirrorMargins/>',
      headers: { h: table }, footers: { f: p('Footer') } }));
    expect(pages.map((page) => lines(page.headerFooter).map((line) => line.text))).toEqual([['1', 'Footer'], ['2', 'Footer'], ['3', 'Footer']]);
    expect(pages.map((page) => lines(page.headerFooter)[1].x)).toEqual([50, 100, 50]);
  });

  it('embeds one image and spends one pixel allocation for two relationship spellings', async () => {
    const xml = (text: string) => new TextEncoder().encode(text);
    const bytes = zipStore([
      ['word/document.xml', xml('<w:document><w:body><w:p>' + inlineImage('a', 8, 8) + inlineImage('b', 16, 16) + '</w:p></w:body></w:document>')],
      ['word/_rels/document.xml.rels', xml('<Relationships><Relationship Id="a" Target="media/red.png"/><Relationship Id="b" Target="MEDIA/RED.PNG"/></Relationships>')],
      ['word/media/red.png', redPng()],
    ]);
    const { pdf } = await convertDocx(bytes, { imageLimits: { maxTotalPixels: 8 } });
    expect(pdfText(pdf).match(/\/Subtype\s*\/Image\b/g)).toHaveLength(1);
  });

  it('charges chart points once for repeated placements and separately for distinct parts', () => {
    const reserve = vi.spyOn(ChartPointBudget.prototype, 'reserve');
    const drawing = (id: string, width: number) => '<w:r><w:drawing><wp:inline><wp:extent cx="' + width * 12700 + '" cy="1270000"/><a:graphic><a:graphicData>' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="' + id + '"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const chart = '<c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numLit><c:ptCount val="3"/><c:pt idx="2"><c:v>7</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart>';
    const doc = parseDocx(makeDocx({ body: '<w:p>' + drawing('a', 80) + drawing('a', 100) + drawing('b', 80) + '</w:p>', charts: { a: chart, b: chart } }), nodeCompression.inflateRaw);
    const drawings = doc.body.flatMap((b) => b.type === 'paragraph' ? b.content.flatMap((i) => i.type === 'drawing' ? [i.drawing] : []) : []);
    expect(drawings.map((d) => d.width)).toEqual([80, 100, 80]);
    expect(drawings[0].chart).toBeDefined();
    expect(drawings[0].chart).toBe(drawings[1].chart);
    expect(drawings[0].chart).not.toBe(drawings[2].chart);
    expect(reserve.mock.calls).toEqual([[3], [3]]);
  });

  it('shares an in-flight packaged font and its parsed tables, and refreshes changed files in an explicit font directory', async () => {
    const face = vi.spyOn(FontRegistry.prototype, 'resolve');
    const doc = makeDocx({ body: p('AV', '', '<w:rFonts w:ascii="Gelasio"/><w:kern w:val="2"/>') });
    const mappings = vi.spyOn(FontBudget.prototype, 'chargeMappingsTo');
    const kerning = vi.spyOn(FontBudget.prototype, 'chargeKerningTo');
    const reads = vi.mocked(readFile);
    reads.mockClear();
    const conversions = await Promise.all([convertDocx(doc), convertDocx(doc)]);
    expect(conversions.map((r) => r.pageCount)).toEqual([1, 1]);
    const fonts = () => face.mock.results.map((r) => r.value?.face).filter((f) => f?.familyName === 'Gelasio');
    const packagedFaces = [...new Set(fonts())];
    expect(packagedFaces).toHaveLength(2);
    expect(packagedFaces[0].font).toBe(packagedFaces[1].font);
    const gelasioReads = () => reads.mock.calls.filter(([file]) => basename(String(file)) === 'Gelasio-Regular.ttf');
    expect(gelasioReads()).toHaveLength(1);
    // The shared tables still count against each document, once per face.
    expect(new Set(mappings.mock.calls.map(([target]) => target)).size).toBe(2);
    expect(kerning.mock.calls).toHaveLength(2);
    expect(new Set(kerning.mock.calls.map(([target]) => target)).size).toBe(2);
    const fontDir = await mkdtemp(join(tmpdir(), 'kanera-fonts-'));
    try {
      const path = join(fontDir, 'Gelasio-Regular.ttf');
      await copyFile(new URL('../fonts/Gelasio-Regular.ttf', import.meta.url), path);
      await convertDocx(doc, { fontDir });
      await convertDocx(doc, { fontDir });
      expect(gelasioReads()).toHaveLength(2);
      expect(new Set(fonts().map((f) => f.font)).size).toBe(2);
      await utimes(path, new Date(0), new Date(0));
      await convertDocx(doc, { fontDir });
      expect(gelasioReads()).toHaveLength(3);
      expect(new Set(fonts().map((f) => f.font)).size).toBe(3);
    } finally {
      await rm(fontDir, { recursive: true, force: true });
    }
  });

  it('embeds the same CJK glyphs in concurrent conversions, including equation bounds', async () => {
    const fontDir = fileURLToPath(new URL('../fonts/', import.meta.url));
    const props = '<w:rFonts w:ascii="Noto Sans CJK JP" w:eastAsia="Noto Sans CJK JP"/>';
    const equation = '<w:p><m:oMath><m:r><w:rPr>' + props + '</w:rPr><m:t>漢</m:t></m:r></m:oMath></w:p>';
    const docs = [p('日本語', '', props), p('東京', '', props), equation].map(body => makeDocx({ body }));
    const full = await Promise.all(docs.map(doc => convertDocx(doc, { fontDir })));
    const actual = await Promise.all(docs.map(doc => convertDocx(doc)));
    for (const [i, result] of actual.entries()) {
      expect(pdfText(result.pdf)).toContain('NotoSansCJKJP-Regular');
      expect(Buffer.from(result.pdf).equals(full[i].pdf)).toBe(true);
      expect(result.warnings).toEqual(full[i].warnings);
    }
  });

  it('loads composite dependencies and rejects outlines from a changed font file', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'kanera-outlines-'));
    try {
      const path = join(dir, 'composite.ttf');
      await copyFile(new URL('../fonts/DejaVuSans.ttf', import.meta.url), path);
      const load = async () => {
        const data = await loadCjkFont(path, fontVersion(await stat(path, { bigint: true })));
        if (!(data instanceof BundledFontData) || !data.font.readSubsetGlyphs) throw new Error('Expected detached outlines');
        return { font: data.font, read: data.font.readSubsetGlyphs };
      };
      const original = new SfntFont(await readFile(path));
      const gids = [original.glyphId(0xe9), original.glyphId(0x20)];
      const expected = subsetTrueType(original, gids);
      const { font, read } = await load();
      const outlines = await read(gids);
      expect(original.compositeComponents(gids[0])).toHaveLength(2);
      expect(subsetTrueType(font, gids, outlines)).toEqual(expected);
      await utimes(path, new Date(0), new Date(0));
      await expect(read(gids)).rejects.toThrow(/changed during conversion/);
      const refreshed = await load();
      expect(subsetTrueType(refreshed.font, gids, await refreshed.read(gids))).toEqual(expected);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it('shares parsed browser fonts and resolves a relative font URL without a window', async () => {
    const requests: URL[] = [];
    let active = 0;
    let peak = 0;
    const fetchFont: typeof fetch = async (input) => {
      const url = new URL(String(input));
      requests.push(url);
      peak = Math.max(peak, ++active);
      const data = await readFile(new URL('../fonts/' + basename(url.pathname), import.meta.url));
      active--;
      return new Response(new Uint8Array(data));
    };
    const resolve = vi.spyOn(FontRegistry.prototype, 'resolve');
    const doc = makeDocx({ body: p('Body', '', '<w:rFonts w:ascii="Gelasio"/>') + p('Text', '', '<w:rFonts w:ascii="Carlito"/>') });
    const options = { fetch: fetchFont, fontBaseUrl: './assets/fonts' };
    await convertBrowser(doc, options);
    const count = requests.length;
    await convertBrowser(doc, options);
    expect(peak).toBeGreaterThan(1);
    expect(requests).toHaveLength(count);
    expect(requests.some((url) => url.pathname.endsWith('/assets/fonts/Gelasio-Regular.ttf'))).toBe(true);
    const faces = new Set(resolve.mock.results.map((r) => r.value.face).filter((face) => face.familyName === 'Gelasio'));
    expect(faces.size).toBe(2);
    expect(new Set([...faces].map((face) => face.font)).size).toBe(1);
  });

  it('retries a failed packaged font read on a later conversion', async () => {
    const doc = makeDocx({ body: p('Body', '', '<w:rFonts w:ascii="Comic Neue"/>') });
    const fs = await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises');
    const reads = vi.mocked(readFile);
    let attempts = 0;
    reads.mockImplementation(async (...args) => {
      if (basename(String(args[0])) === 'ComicNeue-Regular.ttf' && ++attempts === 1) throw new Error('temporary read failure');
      return fs.readFile(...args);
    });
    try {
      const first = await convertDocx(doc);
      expect(first.warnings.some((w) => w.includes('temporary read failure'))).toBe(true);
      const second = await convertDocx(doc);
      expect(second.warnings.some((w) => w.includes('temporary read failure'))).toBe(false);
      expect(attempts).toBe(2);
    } finally {
      reads.mockImplementation(fs.readFile);
    }
  });
});
