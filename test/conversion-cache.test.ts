import { readFile } from 'node:fs/promises';
import { basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { ChartPointBudget } from '../src/docx/chart.js';
import { parseDocx } from '../src/docx/parser.js';
import { FontBudget } from '../src/fonts/sfnt.js';
import { FontRegistry } from '../src/fonts/registry.js';
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

  it('shares an in-flight packaged font and its parsed tables, and reloads an explicit font directory', async () => {
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
    const fontDir = fileURLToPath(new URL('../fonts/', import.meta.url));
    await convertDocx(doc, { fontDir });
    await convertDocx(doc, { fontDir });
    expect(gelasioReads()).toHaveLength(3);
    expect(new Set(fonts().map((f) => f.font)).size).toBe(3);
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
