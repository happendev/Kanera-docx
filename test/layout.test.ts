/**
 * Layout rules that Word applies and that page breaks depend on. Expected numbers come from
 * Word's behaviour with the metric-compatible bundled fonts (Liberation Sans for Arial:
 * winAscent 1854, winDescent 434, lineGap 67 per 2048 units).
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { parseDocx } from '../src/docx/parser.js';
import { BUNDLED_FONTS } from '../src/fonts/bundled.js';
import { FontRegistry } from '../src/fonts/registry.js';
import { layoutDocument } from '../src/layout/engine.js';
import { makeDocx, p } from './helpers/docx.js';
import { glyphOps, layout, lines, pageText } from './helpers/layout.js';
import { inlineImage, redPng } from './helpers/sample.js';

const ARIAL_STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';

/** Arial 11pt single-spaced line: (1854 + 434 + 67) / 2048 * 11. */
const ARIAL_11 = ((1854 + 434 + 67) / 2048) * 11;
/** A4 with 1" margins. */
const BODY_HEIGHT = 841.9 - 144;

const words = (n: number) => Array.from({ length: n }, (_, i) => `word${i % 10}`).join(' ');

describe('line height and spacing', () => {
  it('uses the GDI line height of the font', async () => {
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: p('One') + p('Two') + p('Three') }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11, 2);
    expect(ys[2] - ys[1]).toBeCloseTo(ARIAL_11, 2);
  });

  it('scales the line by a "multiple" spacing, with the extra space below the text', async () => {
    const body = p('One', '<w:spacing w:line="360" w:lineRule="auto" w:after="200"/>') + p('Two') + p('Three', '<w:spacing w:before="100" w:line="480" w:lineRule="exact"/>');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    const descent = (434 / 2048) * 11;
    expect(ys[0]).toBeCloseTo(72 + ARIAL_11 - descent, 1);
    // A 1.5-line box, then 10pt after.
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11 * 1.5 + 10, 1);
    // 5pt before, then an exact 24pt line with the text at its bottom.
    expect(ys[2] - ys[1]).toBeCloseTo(descent + 5 + 24 - descent, 1);
  });

  it('separates paragraphs by the larger of space after and space before', async () => {
    const body = p('One', '<w:spacing w:after="200"/>') + p('Two', '<w:spacing w:before="300" w:after="400"/>') + p('Three', '<w:spacing w:before="100"/>');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11 + 15, 2);
    expect(ys[2] - ys[1]).toBeCloseTo(ARIAL_11 + 20, 2);
  });

  it('collapses spacing between paragraphs of the same style with contextual spacing', async () => {
    const styles =
      ARIAL_STYLES + '<w:style w:type="paragraph" w:styleId="List"><w:name w:val="List"/><w:pPr><w:spacing w:after="200"/><w:contextualSpacing/></w:pPr></w:style>';
    const body = p('A', '<w:pStyle w:val="List"/>') + p('B', '<w:pStyle w:val="List"/>') + p('C');
    const ys = lines((await layout(makeDocx({ styles, body }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11, 2);
    expect(ys[2] - ys[1]).toBeCloseTo(ARIAL_11 + 10, 2);
  });

  it('ignores a larger paragraph mark on a line that has text', async () => {
    const body = '<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:t>Small text, big mark</w:t></w:r></w:p>' + p('Next');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11, 2);
  });

  it('gives a line holding only a picture taller than the text no descent below it', async () => {
    const body = p('Above') + `<w:p>${inlineImage('img1', 100, 50)}</w:p>` + p('Below');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body, media: { img1: { name: 'red.png', data: redPng() } } }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    // Arial's 0.36pt external leading is the only addition to the 50pt picture.
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11 + 50, 0);
  });

  it('does not lay out an empty paragraph whose mark is hidden', async () => {
    const body = p('A') + '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr></w:p>' + p('B');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11, 2);
  });

  it('measures fonts flagged USE_TYPO_METRICS by their typographic metrics', async () => {
    const font = new Uint8Array(readFileSync(new URL('../fonts/LiberationSans-Regular.ttf', import.meta.url)));
    const view = new DataView(font.buffer);
    let os2 = 0;
    for (let i = 0; i < view.getUint16(4); i++) if (String.fromCharCode(...font.subarray(12 + 16 * i, 16 + 16 * i)) === 'OS/2') os2 = view.getUint32(12 + 16 * i + 8);
    view.setUint16(os2 + 62, view.getUint16(os2 + 62) | 0x80);
    const typo = view.getInt16(os2 + 68) - view.getInt16(os2 + 70) + view.getInt16(os2 + 72);
    const doc = parseDocx(makeDocx({ styles: ARIAL_STYLES, body: p('One', '', '<w:rFonts w:ascii="Typo Sans" w:hAnsi="Typo Sans"/>') + p('Two', '', '<w:rFonts w:ascii="Typo Sans" w:hAnsi="Typo Sans"/>') }));
    const fonts = new FontRegistry();
    for (const b of BUNDLED_FONTS) fonts.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, load: async () => readFileSync(new URL(`../fonts/${b.file}`, import.meta.url)) }, 'bundled');
    fonts.add({ family: 'Typo Sans', bold: false, italic: false, load: async () => font }, 'user');
    await fonts.prepare(doc.fontFamilies, doc.codePoints);
    const ys = lines(layoutDocument(doc, fonts).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo((typo / 2048) * 11, 2);
  });

  it('sizes superscripts in whole half-points and leaves the line height alone', async () => {
    const run = (sz: number, sup = false) =>
      `<w:r><w:rPr>${sup ? '<w:vertAlign w:val="superscript"/>' : ''}<w:sz w:val="${sz}"/></w:rPr><w:t>${sup ? '1' : 'x'}</w:t></w:r>`;
    const body = (sup: boolean) => `<w:p>${run(22)}${sup ? run(22, true) : ''}</w:p><w:p>${run(20)}${sup ? run(20, true) : ''}</w:p>` + p('After');
    const ops = glyphOps((await layout(makeDocx({ styles: ARIAL_STYLES, body: body(true) }))).pages[0].body);
    const plain = glyphOps((await layout(makeDocx({ styles: ARIAL_STYLES, body: body(false) }))).pages[0].body);
    const [x11, one11, x10, one10, after] = ops;
    // 11pt: 22 half-points * 2/3 = 14.67, rounded down to 7pt, raised by the remaining 4pt.
    expect(one11.size).toBe(7);
    expect(x11.y - one11.y).toBeCloseTo(4, 5);
    // 10pt: 6.5pt raised 3.5pt.
    expect(one10.size).toBe(6.5);
    expect(x10.y - one10.y).toBeCloseTo(3.5, 5);
    // The raised 7pt glyph reaches above Arial's ascent, yet the lines stay where they were.
    expect(4 + (1854 / 2048) * 7).toBeGreaterThan((1854 / 2048) * 11);
    expect([x11.y, x10.y, after.y]).toEqual(plain.map((o) => o.y));
  });

  it('sizes an empty paragraph by its mark', async () => {
    const body = p('Above') + '<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr></w:p>' + p('Below');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(((1854 + 434 + 67) / 2048) * 24 + ARIAL_11, 1);
  });
});

describe('pagination', () => {
  const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);

  it('fills pages and breaks on explicit page breaks', async () => {
    const body = p('First page') + '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Second page</w:t></w:r></w:p>';
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    expect(r.pages).toHaveLength(2);
    expect(pageText(r.pages[1])).toBe('Second page');
  });

  it(`fits ${linesPerPage} single-spaced lines on an A4 page`, async () => {
    const body = Array.from({ length: linesPerPage + 1 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    expect(lines(r.pages[0].body)).toHaveLength(linesPerPage);
    expect(pageText(r.pages[1])).toBe(`Line ${linesPerPage + 1}`);
  });

  it('lets the extra space of 1.5 line spacing hang past the bottom margin', async () => {
    // 37 lines fit only if the last line's half-line of extra space may overhang.
    const pitch = ARIAL_11 * 1.5;
    expect(36 * pitch + ARIAL_11).toBeLessThan(BODY_HEIGHT);
    expect(37 * pitch).toBeGreaterThan(BODY_HEIGHT);
    const body = Array.from({ length: 38 }, (_, i) => p(`Line ${i + 1}`, '<w:spacing w:line="360" w:lineRule="auto"/>')).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    expect(lines(r.pages[0].body)).toHaveLength(37);
    expect(pageText(r.pages[1])).toBe('Line 38');
  });

  it('keeps the whole 1.5-spaced line above a footer that pushes the body up', async () => {
    // Three footer lines from 708 twips reach above the 1" margin; 37 lines fit only with overhang.
    const bottom = 841.9 - 35.4 - 3 * ARIAL_11;
    expect(bottom).toBeLessThan(841.9 - 72);
    expect(36 * ARIAL_11 * 1.5 + ARIAL_11).toBeLessThan(bottom - 72);
    const sectPr =
      '<w:sectPr><w:footerReference w:type="default" r:id="ftr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const body = Array.from({ length: 38 }, (_, i) => p(`Line ${i + 1}`, '<w:spacing w:line="360" w:lineRule="auto"/>')).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, footers: { ftr1: p('a') + p('b') + p('c') }, body }));
    expect(lines(r.pages[0].body)).toHaveLength(36);
  });

  it('applies widow control by default and not when it is turned off', async () => {
    // Three lines of the 4-line paragraph fit on page 1; the fourth would be a widow.
    const filler = Array.from({ length: linesPerPage - 3 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const para = (pPr: string) => p(words(52), pPr);
    const on = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + para('') }));
    expect(lines(on.pages[0].body).length + lines(on.pages[1].body).length).toBe(linesPerPage - 3 + 4);
    expect(lines(on.pages[1].body)).toHaveLength(2);
    const off = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + para('<w:widowControl w:val="0"/>') }));
    expect(lines(off.pages[1].body)).toHaveLength(1);
  });

  it('moves a whole 3-line paragraph rather than leave a widow or an orphan', async () => {
    const filler = Array.from({ length: linesPerPage - 2 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + p(words(36)) }));
    expect(lines(r.pages[1].body)).toHaveLength(3);
  });

  it('keeps a heading with the next paragraph', async () => {
    const filler = Array.from({ length: linesPerPage - 1 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + p('Heading', '<w:keepNext/>') + p('Body') }));
    expect(pageText(r.pages[1])).toBe('Heading\nBody');
  });

  it('keeps the lines of a paragraph together with keepLines', async () => {
    const filler = Array.from({ length: linesPerPage - 2 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + p(words(36), '<w:keepLines/><w:widowControl w:val="0"/>') }));
    expect(lines(r.pages[1].body)).toHaveLength(3);
  });

  it('does not leave a blank page between a page break and a new-page section break', async () => {
    const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const brk = '<w:r><w:t>One</w:t></w:r><w:r><w:br w:type="page"/></w:r>';
    const next = await layout(makeDocx({ styles: ARIAL_STYLES, body: `<w:p>${brk}</w:p><w:p><w:pPr>${sect}</w:pPr></w:p>` + p('Two') }));
    expect(next.pages.map(pageText)).toEqual(['One', 'Two']);
    const same = await layout(makeDocx({ styles: ARIAL_STYLES, body: `<w:p><w:pPr>${sect}</w:pPr>${brk}</w:p>` + p('Two') }));
    expect(same.pages.map(pageText)).toEqual(['One', 'Two']);
  });

  it('starts a section on a new page with its own size', async () => {
    const landscape = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>';
    const body = `<w:p><w:pPr>${landscape}</w:pPr><w:r><w:t>Landscape</w:t></w:r></w:p>` + p('Portrait');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    expect(r.pages.map((pg) => [Math.round(pg.width), Math.round(pg.height)])).toEqual([
      [842, 595],
      [595, 842],
    ]);
    expect(lines(r.pages[0].body)[0].x).toBeCloseTo(36, 1);
  });
});

describe('numbering', () => {
  const numbering =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="2160" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';
  const item = (text: string, numId: number, ilvl = 0) => p(text, `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);

  it('draws Symbol bullets as real bullets at the hanging indent', async () => {
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, numbering, body: item('Apples', 1) + item('Pears', 1) }));
    const [first] = lines(r.pages[0].body);
    expect(first.text.replace(/\s/g, '')).toBe('•Apples');
    expect(first.ops[0].x).toBeCloseTo(72 + 18, 1);
    // The tab after the bullet runs to the left indent.
    const apples = glyphOps(r.pages[0].body).find((o) => o.texts.join('').startsWith('A'))!;
    expect(apples.x).toBeCloseTo(72 + 36, 1);
  });

  it('counts multi-level lists and restarts sub-levels', async () => {
    const body = item('One', 2) + item('Sub a', 2, 1) + item('Sub b', 2, 1) + item('Deep', 2, 2) + item('Two', 2) + item('Sub again', 2, 1);
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, numbering, body }));
    const labels = lines(r.pages[0].body).map((l) => l.text.split(/\s/)[0]);
    expect(labels).toEqual(['1.One', 'a)Sub', 'b)Sub', 'i.Deep', '2.Two', 'a)Sub']);
  });
});

describe('tables', () => {
  const cell = (text: string, w: number, tcPr = '') => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${tcPr}</w:tcPr>${p(text)}</w:tc>`;
  const borders =
    '<w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="8" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="8" w:space="0" w:color="000000"/></w:tblBorders>';
  const table = (rows: string[][], tblPr = '') =>
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>${tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows
      .map((r) => `<w:tr>${cell(r[0], 3000)}${cell(r[1], 4000)}</w:tr>`)
      .join('')}</w:tbl>`;

  it('places cell text in its column with the cell margins', async () => {
    const settings = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: table([['A1', 'B1'], ['A2', 'B2']]) + p('After') }));
    const ls = lines(r.pages[0].body);
    expect(ls.map((l) => l.ops.map((o) => o.texts.join('')).join('|'))).toEqual(['A1|B1', 'A2|B2', 'After']);
    const [a1, b1] = ls[0].ops;
    // Word 2013+ puts the table edge on the margin, so text is inset by the 5.4pt cell margin.
    expect(a1.x).toBeCloseTo(72 + 5.4, 1);
    expect(b1.x - a1.x).toBeCloseTo(150, 1);
    // Rows grow by the 1pt top border.
    expect(ls[1].y - ls[0].y).toBeCloseTo(ARIAL_11 + 1, 2);
    const strokes = r.pages[0].body.filter((o) => o.kind === 'line');
    expect(strokes.length).toBeGreaterThanOrEqual(6);
  });

  it('adds the top border to a minimum row height', async () => {
    const tbl = table([['A1', 'B1'], ['A2', 'B2']]).replace('<w:tr>', '<w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>');
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: tbl }))).pages[0].body);
    expect(ls[1].y - ls[0].y).toBeCloseTo(30 + 1, 2);
  });

  it('keeps widow/orphan control when a row splits across pages', async () => {
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    const filler = Array.from({ length: linesPerPage - 2 }, (_, i) => p(`Line ${i + 1}`)).join('');
    // Room is left for the 1pt border, "Short" and one more line: the two-line paragraph moves whole.
    const twoLines = (pPr: string) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('Short')}${p(words(6), pPr)}</w:tc></w:tr></w:tbl>`;
    const on = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + twoLines('') }));
    expect(lines(on.pages[0].body).at(-1)!.text).toBe('Short');
    expect(lines(on.pages[1].body)).toHaveLength(2);
    const off = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + twoLines('<w:widowControl w:val="0"/>') }));
    expect(lines(off.pages[1].body)).toHaveLength(1);
  });

  it('applies row-level table property exceptions', async () => {
    let row = 0;
    const tbl = table([['A1', 'B1'], ['A2', 'B2']], '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>').replace(/<w:tr>/g, (m) => (row++ > 0 ? '<w:tr><w:tblPrEx><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tblPrEx>' : m));
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: tbl }));
    const red = r.pages[0].body.filter((o) => o.kind === 'rect' && o.fill === 'FF0000');
    // Only the first row's two cells keep the table's shading.
    expect(red).toHaveLength(2);
  });

  it('floats a table positioned with tblpPr and wraps text beside it', async () => {
    const tbl = table([['A1', 'B1'], ['A2', 'B2']], '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="0"/>');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: tbl + p('Beside') + p('Next') }));
    const ls = lines(r.pages[0].body);
    const beside = ls.find((l) => l.text === 'Beside')!;
    const next = ls.find((l) => l.text === 'Next')!;
    // Floats are painted in front of the body text.
    const a1 = glyphOps(r.pages[0].front).find((o) => o.texts.join('') === 'A1')!;
    // The 350pt table sits against the right margin; the text runs down the 92pt gap on its left.
    expect(a1.x).toBeCloseTo(72 + (595.3 - 144) - 350 + 5.4, 0);
    expect(beside.x).toBeCloseTo(72, 1);
    expect(a1.y - beside.y).toBeCloseTo(1, 1);
    expect(next.y - beside.y).toBeCloseTo(ARIAL_11, 2);
  });

  it('repeats header rows on the next page', async () => {
    const rows = [['Head A', 'Head B'], ...Array.from({ length: 70 }, (_, i) => [`r${i}`, `v${i}`])];
    const tbl = table(rows).replace('<w:tr>', '<w:tr><w:trPr><w:tblHeader/></w:trPr>');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: tbl }));
    expect(r.pages.length).toBe(2);
    expect(lines(r.pages[1].body)[0].text).toBe('Head AHead B');
  });
});

describe('headers, footers and fields', () => {
  it('draws the header and footer on every page with page numbers', async () => {
    const sectPr =
      '<w:sectPr><w:headerReference w:type="default" r:id="hdr1"/><w:footerReference w:type="default" r:id="ftr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const footer =
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';
    const r = await layout(
      makeDocx({
        styles: ARIAL_STYLES,
        sectPr,
        headers: { hdr1: p('Report header') },
        footers: { ftr1: footer },
        body: p('One') + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p('Two'),
      }),
    );
    expect(r.pages).toHaveLength(2);
    r.pages.forEach((pg, i) => {
      const hf = lines(pg.headerFooter);
      expect(hf.map((l) => l.text)).toEqual(['Report header', `Page ${i + 1} of 2`]);
      // The header's first line box starts at the header distance (35.4pt).
      expect(hf[0].y).toBeCloseTo(35.4 + ARIAL_11 - (434 / 2048) * 11, 1);
    });
  });

  it('uses a different first page header when titlePg is set', async () => {
    const sectPr =
      '<w:sectPr><w:headerReference w:type="default" r:id="hdr1"/><w:headerReference w:type="first" r:id="hdr2"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr>';
    const r = await layout(
      makeDocx({ styles: ARIAL_STYLES, sectPr, headers: { hdr1: p('Default'), hdr2: p('Cover') }, body: p('One') + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' }),
    );
    expect(r.pages.map((pg) => lines(pg.headerFooter)[0]?.text)).toEqual(['Cover', 'Default']);
  });
});

describe('fonts', () => {
  it('substitutes metric-compatible fonts for missing Office fonts', async () => {
    const body = ['Calibri', 'Cambria', 'Times New Roman', 'Courier New', 'Georgia', 'Segoe UI', 'Some Unknown Font']
      .map((f) => p(f, '', `<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`))
      .join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    const families = lines(r.pages[0].body).map((l) => l.ops[0].face.familyName);
    expect(families).toEqual(['Carlito', 'Caladea', 'Liberation Serif', 'Liberation Mono', 'Gelasio', 'Selawik', 'Liberation Sans']);
  });

  it('synthesises bold and italic only when the family has no such face', async () => {
    const body = p('Bold', '', '<w:b/>') + p('Bold DejaVu italic', '', '<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:b/><w:i/>');
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    expect(ls[0].ops[0].syntheticBold).toBe(false);
    expect(ls[1].ops[0].syntheticItalic).toBe(true);
  });
});

describe('justification', () => {
  // Liberation Sans has Arial's widths: "a" 1139 and the space 569 units of 2048.
  const a = (1139 / 2048) * 11;
  const space = (569 / 2048) * 11;
  const fiveWords = 5 * 4 * a + 4 * space;
  // The fifth word overflows by 2pt; the line's four spaces can give up a quarter of their width.
  const indentRight = Math.round((595.3 - 144 - (fiveWords - 2)) * 20);
  const para = (jc: string) => p(Array(10).fill('aaaa').join(' '), `<w:ind w:right="${indentRight}"/><w:jc w:val="${jc}"/>`);

  it('shrinks the spaces of a justified line to fit one more word', async () => {
    const justified = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: para('both') }))).pages[0].body);
    expect(justified[0].text.match(/aaaa/g)).toHaveLength(5);
    // The fifth word ends at the right indent: the spaces gave up the 2pt.
    const glyphs = justified[0].ops.flatMap((o) => o.xs.map((x, i) => ({ x: o.x + x, text: o.texts[i] })));
    const lastA = glyphs.filter((g) => g.text === 'a').at(-1)!;
    expect(lastA.x + a).toBeCloseTo(72 + fiveWords - 2, 1);
    const left = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: para('left') }))).pages[0].body);
    expect(left[0].text.match(/aaaa/g)).toHaveLength(4);
  });

  it('only shrinks spaces in Word 2013+ compatibility mode', async () => {
    const settings = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>';
    const word2010 = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: para('both') }))).pages[0].body);
    expect(word2010[0].text.match(/aaaa/g)).toHaveLength(4);
  });
});

describe('fields', () => {
  it('draws table-of-contents entries without the Hyperlink character style', async () => {
    const styles = ARIAL_STYLES + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>';
    const link = (text: string) => `<w:hyperlink w:anchor="_Toc1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${text}</w:t></w:r></w:hyperlink>`;
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${link('Entry')}</w:p>` +
      `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p>${link('Elsewhere')}</w:p>`;
    const ls = lines((await layout(makeDocx({ styles, body }))).pages[0].body);
    expect(ls.find((l) => l.text === 'Entry')!.ops[0].color).not.toBe('0563C1');
    expect(ls.find((l) => l.text === 'Elsewhere')!.ops[0].color).toBe('0563C1');
  });

  it('draws legacy form checkboxes in their saved state', async () => {
    const box = (state: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:sizeAuto/>${state}</w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body = `<w:p>${box('<w:default w:val="0"/>')}<w:r><w:t xml:space="preserve"> No </w:t></w:r>${box('<w:default w:val="1"/>')}<w:r><w:t xml:space="preserve"> Yes</w:t></w:r></w:p>`;
    expect(pageText((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0])).toBe('\u2610 No \u2612 Yes');
  });
});
