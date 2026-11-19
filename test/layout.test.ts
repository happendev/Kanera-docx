/**
 * Layout rules that Word applies and that page breaks depend on. Expected numbers come from
 * Word's behaviour with the metric-compatible bundled fonts (Liberation Sans for Arial:
 * winAscent 1854, winDescent 434, lineGap 67 per 2048 units).
 */
import { nodeCompression } from '../src/util/node-compression.js';
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { parseDocx } from '../src/docx/parser.js';
import { FULL_FONTS } from '../src/fonts/bundled.js';
import { FontRegistry } from '../src/fonts/registry.js';
import { layoutDocument } from '../src/layout/engine.js';
import { A4_SECT, makeDocx, p } from './helpers/docx.js';
import { glyphOps, layout, lines, pageText } from './helpers/layout.js';
import { anchoredImage, inlineImage, redPng } from './helpers/sample.js';

const ARIAL_STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';

/** Arial 11pt single-spaced line: (1854 + 434 + 67) / 2048 * 11. */
const ARIAL_11 = ((1854 + 434 + 67) / 2048) * 11;
/** A4 with 1" margins. */
const BODY_HEIGHT = 841.9 - 144;

const words = (n: number) => Array.from({ length: n }, (_, i) => `word${i % 10}`).join(' ');
const compat = (mode: number) => `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat>`;

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

  it('measures spacing in lines when given, but keeps the points beside a zero', async () => {
    const gap = async (spacing: string) => {
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: p('A') + p('B', spacing) }))).pages[0].body);
      return ls[1].y - ls[0].y;
    };
    expect(await gap('<w:spacing w:before="480" w:beforeLines="0"/>')).toBeCloseTo(ARIAL_11 + 24, 2);
    expect(await gap('<w:spacing w:before="480" w:beforeLines="100"/>')).toBeCloseTo(ARIAL_11 + 12, 2);
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
    const doc = parseDocx(makeDocx({ styles: ARIAL_STYLES, body: p('One', '', '<w:rFonts w:ascii="Typo Sans" w:hAnsi="Typo Sans"/>') + p('Two', '', '<w:rFonts w:ascii="Typo Sans" w:hAnsi="Typo Sans"/>') }), nodeCompression.inflateRaw);
    const fonts = new FontRegistry();
    for (const b of FULL_FONTS) fonts.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, covers: b.covers, load: async () => readFileSync(new URL(`../fonts/${b.file}`, import.meta.url)) }, 'bundled');
    fonts.add({ family: 'Typo Sans', bold: false, italic: false, load: async () => font }, 'user');
    await fonts.prepare(doc.fontRequests);
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

  it('sizes lines by their text and a line of only spaces and tabs by its mark', async () => {
    const big = (t: string) => `<w:r><w:rPr><w:sz w:val="48"/></w:rPr>${t}</w:r>`;
    const body = `<w:p><w:r><w:t>Text</w:t></w:r>${big('<w:t xml:space="preserve">   </w:t>')}</w:p><w:p>${big('<w:t xml:space="preserve">  </w:t>')}${big('<w:tab/>')}</w:p>` + p('After');
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    const text = ls.find((l) => l.text.startsWith('Text'))!;
    const after = ls.find((l) => l.text === 'After')!;
    // Neither the 24pt spaces after "Text" nor the 24pt spaces and tab of the next line count.
    expect(after.y - text.y).toBeCloseTo(2 * ARIAL_11, 2);
  });

  it('drops automatic space before at the start of the document and of a cell, and after at the end of a cell', async () => {
    const auto = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>';
    const cell = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('Cell', auto)}</w:tc></w:tr></w:tbl>`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: p('First', auto) + p('Before') + cell + p('After') }))).pages[0].body);
    const y = (t: string) => ls.find((l) => l.text === t)!.y;
    expect(y('First')).toBeCloseTo(72 + ARIAL_11 - (434 / 2048) * 11, 1);
    // Between paragraphs the 14pt automatic spacing applies.
    expect(y('Before') - y('First')).toBeCloseTo(ARIAL_11 + 14, 1);
    expect(y('Cell') - y('Before')).toBeCloseTo(ARIAL_11, 1);
    expect(y('After') - y('Cell')).toBeCloseTo(ARIAL_11, 1);
  });

  it('applies Normal and its default spacing when no style is marked default and there are no paragraph defaults', async () => {
    const styles =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>';
    const ys = lines((await layout(makeDocx({ styles, body: p('One') + p('Two') }))).pages[0].body).map((l) => l.y);
    // 16pt Arial at 1.15 lines, then 8pt after.
    expect(ys[1] - ys[0]).toBeCloseTo(((1854 + 434 + 67) / 2048) * 16 * 1.15 + 8, 1);
  });

  it('sizes an empty paragraph by its mark', async () => {
    const body = p('Above') + '<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr></w:p>' + p('Below');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(((1854 + 434 + 67) / 2048) * 24 + ARIAL_11, 1);
  });

  it('shades the gap between two shaded paragraphs, but not above plain text or past a bottom border', async () => {
    const shd = (fill: string) => `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
    const yellow = async (pPr: string, next: string) => {
      const body = p('Shaded', shd('FFFF00') + '<w:spacing w:after="240"/>' + pPr) + next;
      const fill = (await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body.find((o) => o.kind === 'rect' && o.fill === 'FFFF00');
      return fill?.kind === 'rect' ? fill.h : 0;
    };
    const blue = p('Blue', shd('0000FF'));
    // One 11pt line, then the 12pt after.
    expect(await yellow('', blue)).toBeCloseTo(ARIAL_11 + 12, 1);
    expect(await yellow('', p('Plain'))).toBeCloseTo(ARIAL_11, 1);
    expect(await yellow('<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>', blue)).toBeLessThan(ARIAL_11 + 6);
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

  it('lets the extra space of a 1.5-spaced line reach past a tall footer but not into the margin', async () => {
    // Five footer lines and 10pt after them put the footer top at 733.25pt: the 35th line's text ends
    // above it at 729.75pt, and its line box ends at 736.07pt, still above the 769.9pt margin.
    const pitch = ARIAL_11 * 1.5;
    const footerTop = 841.9 - 35.4 - 5 * ARIAL_11 - 10;
    expect(72 + 34 * pitch + ARIAL_11).toBeLessThan(footerTop);
    expect(72 + 35 * pitch).toBeGreaterThan(footerTop);
    const sectPr =
      '<w:sectPr><w:footerReference w:type="default" r:id="ftr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const footer = ['a', 'b', 'c', 'd'].map((t) => p(t)).join('') + p('e', '<w:spacing w:after="200"/>');
    const body = Array.from({ length: 38 }, (_, i) => p(`Line ${i + 1}`, '<w:spacing w:line="360" w:lineRule="auto"/>')).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, footers: { ftr1: footer }, body }));
    expect(lines(r.pages[0].body)).toHaveLength(35);
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

  it('gives no room to an empty paragraph that ends a section before a continuous one', async () => {
    const continuous = A4_SECT.replace('<w:pgSz', '<w:type w:val="continuous"/><w:pgSz');
    const body = p('One') + `<w:p><w:pPr>${A4_SECT}</w:pPr></w:p>` + p('Two');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, sectPr: continuous }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11, 2);
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
  it('reads page margins written with units', async () => {
    const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="20mm" w:right="1440" w:bottom="1440" w:left="1in" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const [line] = lines((await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, body: p('Margins') }))).pages[0].body);
    expect(line.x).toBeCloseTo(72, 2);
    expect(line.y).toBeCloseTo((20 / 25.4) * 72 + ARIAL_11 - (434 / 2048) * 11, 2);
  });

  const MARGINS = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>';
  const oddEven = '<w:headerReference w:type="default" r:id="hOdd"/><w:headerReference w:type="even" r:id="hEven"/>';
  const twoSections = (settings: string | undefined, second: string) =>
    makeDocx({
      styles: ARIAL_STYLES,
      settings,
      headers: { hOdd: p('Odd header'), hEven: p('Even header') },
      body: p('One') + `<w:p><w:pPr><w:sectPr>${oddEven}${MARGINS}</w:sectPr></w:pPr></w:p>` + p('Two'),
      sectPr: `<w:sectPr>${oddEven}${second}${MARGINS}</w:sectPr>`,
    });

  it('prints an empty page before a section that restarts numbering on the same parity, with odd and even headers', async () => {
    const r = await layout(twoSections('<w:evenAndOddHeaders/>', '<w:pgNumType w:start="1"/>'));
    expect(r.pages.map((pg) => pg.number)).toEqual([1, 2, 1]);
    expect(lines(r.pages[1].body)).toEqual([]);
    expect(r.pages[1].headerFooter).toEqual([]);
    expect(lines(r.pages[2].headerFooter).map((l) => l.text)).toEqual(['Odd header']);
    // With one header for all pages, page numbers may repeat parity.
    expect((await layout(twoSections(undefined, '<w:pgNumType w:start="1"/>'))).pages.map((pg) => pg.number)).toEqual([1, 1]);
  });

  it('prints an empty page before an odd-page section that would start on an even page', async () => {
    const r = await layout(twoSections(undefined, '<w:type w:val="oddPage"/>'));
    expect(r.pages.map((pg) => pg.number)).toEqual([1, 2, 3]);
    expect(r.pages[1].headerFooter).toEqual([]);
    expect(pageText(r.pages[2])).toBe('Two');
  });

  it('keeps an empty section-break paragraph on a full page before a new-page section', async () => {
    const filler = Array.from({ length: linesPerPage }, (_, i) => p(`Line ${i + 1}`)).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + `<w:p><w:pPr>${A4_SECT}</w:pPr></w:p>` + p('Next section') }));
    expect(r.pages).toHaveLength(2);
    expect(pageText(r.pages[1])).toBe('Next section');
  });

  it("measures a new section's space before against the empty section-break paragraph", async () => {
    const body = p('Before', '<w:spacing w:after="240"/>') + `<w:p><w:pPr>${A4_SECT}</w:pPr></w:p>` + p('Heading', '<w:spacing w:before="280"/>');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    // The section's first paragraph keeps the space before that exceeds the previous paragraph's
    // space after, and that paragraph is the empty one (0pt), not "Before" (12pt).
    expect(lines(r.pages[1].body)[0].y).toBeCloseTo(72 + 14 + ARIAL_11 - (434 / 2048) * 11, 2);
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

  it("lends a bullet's ascent to the line but never its descent", async () => {
    // Symbol (ascent 2059/2048, kept although the bullet is drawn from another font) is taller
    // than Arial's ascent plus line gap (1854 + 67) and pushes the baseline down by the difference.
    // Courier New descends further than Arial but leaves the line as tall as plain text.
    const courier = '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>';
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, numbering: numbering + courier, body: item('One', 1) + item('Two', 1) + item('Three', 3) + item('Four', 3) + p('Five') }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    expect(ys[0]).toBeCloseTo(72 + (2059 / 2048) * 11, 2);
    expect(ys[1] - ys[0]).toBeCloseTo(((2059 + 434) / 2048) * 11, 2);
    expect(ys[3] - ys[2]).toBeCloseTo(ARIAL_11, 2);
    expect(ys[4] - ys[3]).toBeCloseTo(ARIAL_11, 2);
  });

  it('counts multi-level lists and restarts sub-levels', async () => {
    const body = item('One', 2) + item('Sub a', 2, 1) + item('Sub b', 2, 1) + item('Deep', 2, 2) + item('Two', 2) + item('Sub again', 2, 1);
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, numbering, body }));
    const labels = lines(r.pages[0].body).map((l) => l.text.split(/\s/)[0]);
    expect(labels).toEqual(['1.One', 'a)Sub', 'b)Sub', 'i.Deep', '2.Two', 'a)Sub']);
  });

  it('scales only the text of a "multiple" line, not what a taller bullet adds', async () => {
    const spaced = (text: string) => p(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:spacing w:line="360" w:lineRule="auto"/>');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, numbering, body: spaced('One') + spaced('Two') }))).pages[0].body).map((l) => l.y);
    // Symbol's ascent raises the line to (2059 + 434) / 2048; the extra half line is Arial's.
    expect(ys[1] - ys[0]).toBeCloseTo(((2059 + 434) / 2048) * 11 + ARIAL_11 * 0.5, 2);
  });

  it('leaves automatic spacing out between items of one list but keeps it around the list', async () => {
    const auto = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>';
    const listItem = (text: string, numId: number) => p(text, `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>${auto}`);
    const body = p('Before', auto) + listItem('One', 1) + listItem('Two', 1) + listItem('Three', 2);
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, numbering, body }))).pages[0].body);
    const y = (t: string) => ls.find((l) => l.text.endsWith(t))!.y;
    const bulletLine = ((2059 + 434) / 2048) * 11;
    expect(y('One') - y('Before')).toBeCloseTo(bulletLine + 14, 1);
    expect(y('Two') - y('One')).toBeCloseTo(bulletLine, 1);
    // A different list keeps the 14pt between them.
    expect(y('Three') - y('Two')).toBeCloseTo(ARIAL_11 + 14, 1);
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

  it("sizes the columns from the cells when every cell's width contradicts the grid", async () => {
    // A Word-saved package (it has rsids), so the grid is otherwise taken as Word computed it.
    const settings = compat(15) + '<w:rsids><w:rsidRoot w:val="00A1B2C3"/></w:rsids>';
    const auto = `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${p('B')}</w:tc>`;
    const tbl = (second: string, tr = '<w:tr>') =>
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${tr}${cell('A', 1000)}${second}</w:tr></w:tbl>`;
    const gap = async (second: string, tr?: string) => {
      const [a, b] = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: tbl(second, tr) }))).pages[0].body)[0].ops;
      return b.x - a.x;
    };
    expect(await gap(cell('B', 5000))).toBeCloseTo(50, 1);
    // An automatic cell leaves the columns to the grid, and so does a row Word saved.
    expect(await gap(auto)).toBeCloseTo(150, 1);
    expect(await gap(cell('B', 5000), '<w:tr w:rsidR="00A1B2C3">')).toBeCloseTo(150, 1);
  });

  it('draws a double border as two lines of its width with as wide a gap, and makes room for it', async () => {
    const top = (style: string) => `<w:tcBorders><w:top w:val="${style}" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>`;
    const doc = (style: string) =>
      makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('A', 3000, top(style))}</w:tr></w:tbl>` });
    const double = await layout(doc('double'));
    const single = await layout(doc('single'));
    expect(lines(double.pages[0].body)[0].y - lines(single.pages[0].body)[0].y).toBeCloseTo(1.5, 2);
    const strokes = double.pages[0].body.filter((o) => o.kind === 'line');
    expect(strokes.map((o) => o.kind === 'line' && o.width)).toEqual([0.75, 0.75]);
    expect(strokes.map((o) => o.kind === 'line' && o.y1).reduce((a, b) => Math.abs(Number(a) - Number(b)))).toBeCloseTo(1.5, 2);
  });

  it("adds the cells' top and bottom margins to a minimum row height", async () => {
    const tcMar = '<w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/></w:tcMar>';
    const row = (text: string, trPr = '') => `<w:tr>${trPr && `<w:trPr>${trPr}</w:trPr>`}${cell(text, 3000, tcMar)}${cell(text, 4000, tcMar)}</w:tr>`;
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="4000"/></w:tblGrid>${row('Tall', '<w:trHeight w:val="600" w:hRule="atLeast"/>')}${row('Plain')}${row('End')}</w:tbl>`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: tbl }))).pages[0].body);
    const y = (t: string) => ls.find((l) => l.text.startsWith(t))!.y;
    // 30pt for the content, then the 2pt margins above and below it.
    expect(y('Plain') - y('Tall')).toBeCloseTo(30 + 4, 1);
    expect(y('End') - y('Plain')).toBeCloseTo(ARIAL_11 + 4, 1);
  });

  it('hangs the cell margin left of the indent in a package without a settings part', async () => {
    // Such a package opens in Word 2007 mode, where the indent positions the first cell's text.
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: table([['A1', 'B1']]) }));
    expect(lines(r.pages[0].body)[0].ops[0].x).toBeCloseTo(72, 1);
  });

  const WORD_2013 = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';
  const margins = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
  /** A one-cell, 150pt table holding a nested table (columns in twips) and then `after`. */
  const nesting = (tblInd: number, cols: number[], after: string) => {
    const inner = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblInd w:w="${tblInd}" w:type="dxa"/>${margins}</w:tblPr><w:tblGrid>${cols.map((c) => `<w:gridCol w:w="${c}"/>`).join('')}</w:tblGrid><w:tr>${cols
      .map((c, i) => `<w:tc><w:tcPr><w:tcW w:w="${c}" w:type="dxa"/></w:tcPr>${p('N' + i)}</w:tc>`)
      .join('')}</w:tr></w:tbl>`;
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${margins}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${inner}${after}</w:tc></w:tr></w:tbl>`;
  };

  it('ends a cell at its nested table when only an empty paragraph follows it', async () => {
    const gap = async (after: string) => {
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: WORD_2013, body: nesting(0, [2000], after) + p('After') }))).pages[0].body);
      return ls.find((l) => l.text === 'After')!.y - ls.find((l) => l.text === 'N0')!.y;
    };
    // The paragraph that has to close the cell adds a line only when it has text.
    expect((await gap(p('X'))) - (await gap('<w:p/>'))).toBeCloseTo(ARIAL_11, 2);
  });

  it('keeps a nested table inside its cell, shrinking it to fit', async () => {
    // The nested table asks to start 25pt left of the cell's text and is 200pt wide, with 139.2pt
    // of room: it moves in to the text edge and its columns shrink in proportion.
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, settings: WORD_2013, body: nesting(-500, [2000, 2000], '<w:p/>') }));
    const [n0, n1] = lines(r.pages[0].body).find((l) => l.text === 'N0N1')!.ops;
    expect(n0.x).toBeCloseTo(72 + 5.4 + 5.4, 1);
    expect(n1.x - n0.x).toBeCloseTo((150 - 10.8) / 2, 1);
  });

  it('collapses a row of empty cells marked hideMark', async () => {
    const spacer = (tcPr: string) =>
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('A1', 3000)}</w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr><w:p/></w:tc></w:tr></w:tbl>${p('After')}`;
    const pitch = async (tcPr: string) => {
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: spacer(tcPr) }))).pages[0].body);
      return ls[1].y - ls[0].y;
    };
    expect(await pitch('<w:hideMark/>')).toBeCloseTo(ARIAL_11, 2);
    expect(await pitch('')).toBeCloseTo(2 * ARIAL_11, 2);
  });

  it('adds the top border to a minimum row height', async () => {
    const tbl = table([['A1', 'B1'], ['A2', 'B2']]).replace('<w:tr>', '<w:tr><w:trPr><w:trHeight w:val="600"/></w:trPr>');
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: tbl }))).pages[0].body);
    expect(ls[1].y - ls[0].y).toBeCloseTo(30 + 1, 2);
  });

  it("adds the bottom border of the row above by what it is wider than the cell's top margin", async () => {
    const pitch = async (marginTop: number) => {
      const mar = `<w:tcMar><w:top w:w="${marginTop}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tcMar>`;
      const upper = cell('Upper', 3000, `<w:tcBorders><w:bottom w:val="single" w:sz="16" w:space="0" w:color="000000"/></w:tcBorders>${mar}`);
      const lower = cell('Lower', 3000, `<w:tcBorders><w:top w:val="nil"/></w:tcBorders>${mar}`);
      const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${upper}</w:tr><w:tr>${lower}</w:tr></w:tbl>`;
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: tbl }))).pages[0].body);
      return ls[1].y - ls[0].y;
    };
    // The 2pt border adds in full without a margin, and fits inside a 3pt one.
    expect(await pitch(0)).toBeCloseTo(ARIAL_11 + 2, 2);
    expect(await pitch(60)).toBeCloseTo(ARIAL_11 + 3, 2);
  });

  it('moves a row whose content fits but whose minimum height does not to the next page whole', async () => {
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    const filler = Array.from({ length: linesPerPage - 3 }, (_, i) => p(`Line ${i + 1}`)).join('');
    // Three lines of room are left: enough for the two paragraphs but not for the 60pt minimum height.
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p('Tall')}${p('Second')}</w:tc></w:tr></w:tbl>`;
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: filler + tbl }));
    expect(lines(r.pages[0].body).at(-1)!.text).toBe(`Line ${linesPerPage - 3}`);
    expect(lines(r.pages[1].body)[0].text).toBe('Tall');
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

  it('insets every cell by the cell spacing', async () => {
    const settings = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';
    const run = async (tblPr: string) => lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: table([['A1', 'B1'], ['A2', 'B2']], tblPr) }))).pages[0].body);
    const plain = await run('');
    const spaced = await run('<w:tblCellSpacing w:w="40" w:type="dxa"/>');
    // Word's spacing value is half the gap: cells sit 4pt apart and 4pt in from the table's
    // edges, so rows are 4pt further apart and text moves 4pt in.
    expect(spaced[0].ops[0].x - plain[0].ops[0].x).toBeCloseTo(4, 2);
    expect(spaced[1].y - spaced[0].y - (plain[1].y - plain[0].y)).toBeCloseTo(4, 2);
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
  it("gives cell text the size of a table style without one only before Word 2013", async () => {
    const styles = ARIAL_STYLES.replace('<w:sz w:val="22"/>', '<w:sz w:val="20"/>')
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
      + '<w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/><w:rPr><w:color w:val="1F3864"/></w:rPr></w:style>';
    const body = `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell('Cell', 4000)}</w:tr></w:tbl>` + p('After');
    const sizes = async (mode: number) => lines((await layout(makeDocx({ styles, settings: compat(mode), body }))).pages[0].body).map((l) => l.ops[0].size);
    // Word 2010 gives the cell the document defaults' 10pt over the 11pt Normal style; Word 2013
    // (overrideTableStyleFontSizeAndJustification) keeps Normal's size.
    expect(await sizes(14)).toEqual([10, 11]);
    expect(await sizes(15)).toEqual([11, 11]);
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

  it('wraps header text around the header\'s own picture, but not its empty lines', async () => {
    const sectPr =
      '<w:sectPr><w:headerReference w:type="default" r:id="hdr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    // A 30pt picture across the header, ending above the 72pt top margin.
    const logo = anchoredImage('none', 451, 30, { h: 'column', x: 0, v: 'paragraph', y: 0 }, '<wp:wrapSquare wrapText="bothSides"/>');
    const baseline = ARIAL_11 - (434 / 2048) * 11;
    const withText = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, headers: { hdr1: `<w:p>${logo}<w:r><w:t>Dated</w:t></w:r></w:p>` }, body: p('Body') }));
    // The line moves below the picture, and the taller header pushes the body down.
    expect(lines(withText.pages[0].headerFooter)[0].y).toBeCloseTo(35.4 + 30 + baseline, 1);
    expect(lines(withText.pages[0].body)[0].y).toBeCloseTo(35.4 + 30 + ARIAL_11 + baseline, 1);
    // Blank lines stay where they are: the two-line header leaves the body at the top margin.
    const empty = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, headers: { hdr1: `<w:p>${logo}</w:p><w:p/>` }, body: p('Body') }));
    expect(lines(empty.pages[0].body)[0].y).toBeCloseTo(72 + baseline, 1);
  });

  it('positions a picture laid out in a header table cell from the cell, and grows the row to hold it', async () => {
    const sectPr =
      '<w:sectPr><w:headerReference w:type="default" r:id="hdr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const logo = anchoredImage('none', 20, 60, { h: 'margin', x: 10, v: 'margin', y: 5 }, '<wp:wrapSquare wrapText="bothSides"/>');
    const header = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${logo}</w:p></w:tc></w:tr></w:tbl>`;
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, headers: { hdr1: header }, body: p('Body') }));
    // Measured from the page margin the picture would sit over the body and push "Body" aside;
    // in its cell it ends 65pt below the header's top, which becomes the header's height.
    const body = lines(r.pages[0].body)[0];
    expect(body.x).toBeCloseTo(72, 1);
    expect(body.y).toBeCloseTo(35.4 + 65 + ARIAL_11 - (434 / 2048) * 11, 1);
  });

  it('places a framed page number beside the footer paragraph that follows it', async () => {
    const sectPr =
      '<w:sectPr><w:footerReference w:type="default" r:id="ftr1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const number = '<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, footers: { ftr1: number + p('Footer text') }, body: p('Body') }));
    const hf = lines(r.pages[0].headerFooter);
    const text = hf.find((l) => l.text === 'Footer text')!;
    const page = hf.find((l) => l.text === '1')!;
    // The frame takes no room: the footer is one line, with the number on it against the right margin.
    expect(text.y).toBeCloseTo(841.9 - 35.4 - (434 / 2048) * 11, 1);
    // w:y offsets the frame 1 twip below the top of that paragraph.
    expect(page.y).toBeCloseTo(text.y + 0.05, 2);
    expect(page.x).toBeCloseTo(595.3 - 72 - (1139 / 2048) * 11, 1);
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

describe('floating objects', () => {
  const media = { img1: { name: 'red.png', data: redPng() } };
  const square = '<wp:wrapSquare wrapText="bothSides"/>';

  it('keeps the first-line indent from the edge of an object the line wraps beside', async () => {
    const pic = anchoredImage('img1', 50, 40, { h: 'margin', x: 0, v: 'paragraph', y: 0 }, square);
    const body = `<w:p><w:pPr><w:ind w:firstLine="360"/></w:pPr>${pic}<w:r><w:t>${words(30)}</w:t></w:r></w:p>`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body);
    expect(ls[0].x).toBeCloseTo(72 + 50 + 18, 1);
    expect(ls[1].x).toBeCloseTo(72 + 50, 1);
  });

  it('wraps text above the anchor around an object placed against the margin', async () => {
    // The picture belongs to the second paragraph but sits at the top of the page, beside the first.
    const pic = anchoredImage('img1', 151.3, 30, { h: 'margin', x: 300, v: 'margin', y: 0 }, square);
    const body = p(words(12)) + `<w:p>${pic}<w:r><w:t>Anchor</w:t></w:r></w:p>`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body);
    const starts = ls[0].ops.flatMap((o) => o.xs.map((x) => o.x + x));
    expect(Math.max(...starts)).toBeLessThan(72 + 300);
    expect(ls.find((l) => l.text === 'Anchor')!.y).toBeGreaterThan(ls[1].y);
  });
});

describe('fonts', () => {
  it('draws a mathematical letter no font has as its plain letter in the style it stands for', async () => {
    const styles = ARIAL_STYLES.replace(/Arial/g, 'Times New Roman');
    const r = await layout(makeDocx({ styles, body: p('\u{1D6FC}\u{1D431}') }));
    const ops = glyphOps(r.pages[0].body);
    const glyph = (t: string) => ops.find((o) => o.texts.includes(t))!;
    // Italic alpha and bold x, copied as the mathematical letters they are.
    expect(glyph('\u{1D6FC}').face.italic).toBe(true);
    expect(glyph('\u{1D431}').face.bold).toBe(true);
    expect(ops.flatMap((o) => o.gids)).not.toContain(0);
  });

  it('draws Aptos with Morrow Sans, keeping its line height on other fallbacks too', async () => {
    const styles = ARIAL_STYLES.replace(/Arial/g, 'Aptos');
    const docx = makeDocx({ styles, body: p('One') + p('Two') + p('Display', '', '<w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/>') });
    const pitch = ((1923 + 577) / 2048) * 11;
    const morrow = await layoutWithWarnings(docx);
    const ls = lines(morrow.result.pages[0].body);
    expect(ls.map((l) => l.ops[0].face.font.names.family)).toEqual(['Morrow Sans', 'Morrow Sans', 'Morrow Sans Display']);
    expect(ls[1].y - ls[0].y).toBeCloseTo(pitch, 2);
    expect(morrow.warnings).toEqual([]);
    // Without Morrow Sans (a caller's own font set), Selawik stands in with Aptos's line height.
    const fallback = await layoutWithWarnings(docx, (file) => !file.startsWith('Morrow'));
    const fl = lines(fallback.result.pages[0].body);
    expect(fl[0].ops[0].face.familyName).toBe('Selawik');
    expect(fl[1].y - fl[0].y).toBeCloseTo(pitch, 2);
  });

  it('substitutes metric-compatible fonts for missing Office fonts', async () => {
    const body = ['Calibri', 'Cambria', 'Times New Roman', 'Courier New', 'Georgia', 'Segoe UI', 'Some Unknown Font']
      .map((f) => p(f, '', `<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`))
      .join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    const families = lines(r.pages[0].body).map((l) => l.ops[0].face.familyName);
    expect(families).toEqual(['Carlito', 'Caladea', 'Liberation Serif', 'Liberation Mono', 'Gelasio', 'Selawik', 'Liberation Sans']);
  });

  it('uses narrow widths and Word line metrics for Arial Narrow, including full face names', async () => {
    const text = 'Width check';
    const body = p(text) + p(text, '', '<w:rFonts w:ascii="Arial Narrow"/>')
      + p(text, '', '<w:rFonts w:ascii="Arial Narrow Bold Italic"/>');
    const { result, warnings } = await layoutWithWarnings(makeDocx({ styles: ARIAL_STYLES, body }));
    const [regular, narrow, boldItalic] = lines(result.pages[0].body);
    expect(narrow.ops[0].face.familyName).toBe('Liberation Sans Narrow');
    expect(narrow.ops[0].xs.at(-1)!).toBeLessThan(regular.ops[0].xs.at(-1)!);
    expect(boldItalic.ops[0].face.font.names.postscriptName).toBe('LiberationSansNarrow-BoldItalic');
    expect(boldItalic.y - narrow.y).toBeCloseTo(11 * 2350 / 2048, 2);
    expect(warnings).toEqual([]);
  });

  it('warns when a custom font directory lacks Liberation Sans Narrow', async () => {
    const docx = makeDocx({ styles: ARIAL_STYLES, body: p('Narrow', '', '<w:rFonts w:ascii="Arial Narrow"/>') });
    const { warnings } = await layoutWithWarnings(docx, (file) => !file.startsWith('LiberationSansNarrow'));
    expect(warnings).toContain('Font "Arial Narrow" not available; using Liberation Sans');
  });

  it('keeps Menlo and Monaco monospaced without font-table hints', async () => {
    const body = ['Menlo', 'Monaco Bold Oblique']
      .map((family) => p('WiéΩЖ', '', `<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>`)).join('');
    const { result, warnings } = await layoutWithWarnings(makeDocx({ styles: ARIAL_STYLES, body }));
    const ops = glyphOps(result.pages[0].body);
    expect(ops.map((op) => op.face.font.names.postscriptName)).toEqual(['DejaVuSansMono', 'DejaVuSansMono-BoldOblique']);
    for (const op of ops) {
      expect(op.gids.every((gid) => gid > 0)).toBe(true);
      op.xs.forEach((x, i) => expect(x).toBeCloseTo(i * op.xs[1], 5));
    }
    expect(warnings).toEqual([
      'Font "Menlo" not available; using DejaVu Sans Mono',
      'Font "Monaco Bold Oblique" not available; using DejaVu Sans Mono',
    ]);
  });

  it('synthesises bold and italic only when the family has no such face', async () => {
    const body = p('Bold', '', '<w:b/>') + p('Bold DejaVu italic', '', '<w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:b/><w:i/>');
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    expect(ls[0].ops[0].syntheticBold).toBe(false);
    expect(ls[1].ops[0].syntheticItalic).toBe(true);
  });

  async function layoutWithWarnings(docx: Uint8Array, use: (file: string) => boolean = () => true) {
    const doc = parseDocx(docx, nodeCompression.inflateRaw);
    const fonts = new FontRegistry();
    for (const b of FULL_FONTS.filter((f) => use(f.file))) fonts.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, covers: b.covers, load: async () => readFileSync(new URL(`../fonts/${b.file}`, import.meta.url)) }, 'bundled');
    await fonts.prepare(doc.fontRequests);
    return { result: layoutDocument(doc, fonts), warnings: [...fonts.warnings] };
  }

  it('draws the characters bundled Caladea lacks from a fallback font', async () => {
    // Caladea 1.002 matches Cambria's widths but has no comma-below letters, these combining marks
    // or superscript four; they must come from another font rather than as missing glyphs.
    const marks = [0x300, 0x301, 0x302, 0x303, 0x304, 0x306, 0x307, 0x308, 0x30a, 0x30b, 0x30c, 0x326, 0x327, 0x328];
    const text = 'Ș ș Ț ț ' + marks.map((m) => 'a' + String.fromCodePoint(m)).join(' ') + ' x⁴';
    const fallback = text.replace(/[ ax]/g, '');
    for (const style of ['', '<w:b/>', '<w:i/>', '<w:b/><w:i/>']) {
      const body = p(text, '', `<w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/>${style}`);
      const { result, warnings } = await layoutWithWarnings(makeDocx({ styles: ARIAL_STYLES, body }));
      const ops = glyphOps(result.pages[0].body);
      const drawn = ops.flatMap((o) => o.texts.map((t, i) => ({ t, gid: o.gids[i], face: o.face.familyName })));
      expect(drawn.filter((g) => fallback.includes(g.t)).map((g) => g.t).join('')).toBe(fallback);
      expect(drawn.filter((g) => g.gid === 0)).toEqual([]);
      expect(drawn.filter((g) => g.t === 'a').every((g) => g.face === 'Caladea')).toBe(true);
      expect(warnings.filter((w) => w.includes('No font has a glyph'))).toEqual([]);
    }
  });

  it('maps symbol-font characters to Unicode and warns about those no font can draw', async () => {
    const sym = (font: string, char: string) => `<w:r><w:sym w:font="${font}" w:char="${char}"/></w:r>`;
    // Webdings clipboard, Wingdings ➢ (drawn as the older U+27A2 look-alike), Wingdings 3 rightwards
    // arrow, and a Wingdings 2 code with no Unicode equivalent.
    const body = `<w:p>${sym('Webdings', 'F0A4')}${sym('Wingdings', 'F0D8')}${sym('Wingdings 3', 'F067')}${sym('Wingdings 2', 'F0FF')}</w:p>`;
    const { result, warnings } = await layoutWithWarnings(makeDocx({ styles: ARIAL_STYLES, body }));
    const ops = glyphOps(result.pages[0].body);
    expect(ops.flatMap((o) => o.texts)).toEqual(['📋', '➢', '→', '']);
    expect(warnings).toEqual([
      'No font has a glyph for Webdings character U+F0A4 (U+1F4CB); it needs the Webdings font or one with U+1F4CB',
      'Wingdings 2 character U+F0FF has no Unicode equivalent; it needs the Wingdings 2 font',
    ]);
  });

  it('keeps Wingdings line metrics whichever font draws the symbol', async () => {
    // Wingdings 0x6C (U+26AB) comes from DejaVu Sans, 0x80 (U+1F10B) only from Noto Sans CJK JP
    // (loaded here for the Japanese line); both lines are as tall as Wingdings makes them.
    const nextLineY = async (char: string) => {
      const body = `<w:p><w:r><w:sym w:font="Wingdings" w:char="${char}"/></w:r></w:p>` + p('Next') + p('日本語');
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
      expect(ls[0].ops[0].gids[0]).toBeGreaterThan(0);
      return ls[1].y;
    };
    expect(await nextLineY('F080')).toBeCloseTo(await nextLineY('F06C'), 2);
  });
  it("gives missing Office fonts Word's line height and a substitute of about their width", async () => {
    const run = (f: string) => `<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`;
    const body = p('Tahoma', '', run('Tahoma')) + p('Tahoma', '', run('Tahoma')) + p('Verdana', '', run('Verdana')) + p('Verdana', '', run('Verdana'));
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    // Tahoma is about as wide as Liberation Sans; Verdana as DejaVu Sans.
    expect(ls.map((l) => l.ops[0].face.familyName)).toEqual(['Liberation Sans', 'Liberation Sans', 'DejaVu Sans', 'DejaVu Sans']);
    expect(ls[1].y - ls[0].y).toBeCloseTo(((2049 + 423) / 2048) * 11, 1);
    expect(ls[3].y - ls[2].y).toBeCloseTo(((2059 + 430) / 2048) * 11, 1);
  });

  it("keeps Century Gothic's line height whichever font stands in", async () => {
    const run = '<w:rFonts w:ascii="Century Gothic" w:hAnsi="Century Gothic"/>';
    const { result } = await layoutWithWarnings(makeDocx({ styles: ARIAL_STYLES, body: p('One', '', run) + p('Two', '', run) }));
    const [one, two] = lines(result.pages[0].body);
    expect(one.ops[0].face.familyName).toBe('Liberation Sans');
    expect(two.y - one.y).toBeCloseTo(((2060 + 451) / 2048) * 11, 1);
  });

  it('sizes the line of a symbol the font lacks by Segoe UI Symbol, which Word draws it from', async () => {
    const height = async (middle: string) => {
      const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: p('Above') + middle + p('Below') }))).pages[0].body).map((l) => l.y);
      return ys[2] - ys[0];
    };
    const plain = await height(p('Task'));
    // 2210 + 514 of 2048 units against Arial's 1854 + 434.
    expect((await height(p('☐ Task'))) - plain).toBeCloseTo(((2210 - 1854 + 514 - 434) / 2048) * 11, 1);
    // A legacy form check box is a box Word draws in the line as it is.
    const field = (type: string, inner = '') => `<w:r><w:fldChar w:fldCharType="${type}">${inner}</w:fldChar></w:r>`;
    const formBox =
      '<w:p>' + field('begin', '<w:ffData><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData>') +
      '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>' + field('separate') + field('end') +
      '<w:r><w:t xml:space="preserve"> Task</w:t></w:r></w:p>';
    expect(await height(formBox)).toBeCloseTo(plain, 2);
    // MS Gothic has the box itself (it is Word's content-control check box font).
    const gothic = '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:eastAsia="MS Gothic"/>';
    expect(await height(p('☐ Task', '', gothic))).toBeCloseTo(await height(p('中 Task', '', gothic)), 2);
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
  const settings = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';

  it('shrinks the spaces of a justified line to fit one more word', async () => {
    const justified = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: para('both') }))).pages[0].body);
    expect(justified[0].text.match(/aaaa/g)).toHaveLength(5);
    // The fifth word ends at the right indent: the spaces gave up the 2pt.
    const glyphs = justified[0].ops.flatMap((o) => o.xs.map((x, i) => ({ x: o.x + x, text: o.texts[i] })));
    const lastA = glyphs.filter((g) => g.text === 'a').at(-1)!;
    expect(lastA.x + a).toBeCloseTo(72 + fiveWords - 2, 1);
    const left = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: para('left') }))).pages[0].body);
    expect(left[0].text.match(/aaaa/g)).toHaveLength(4);
  });

  it("also shrinks the spaces of a paragraph's last line", async () => {
    const five = p(Array(5).fill('aaaa').join(' '), `<w:ind w:right="${indentRight}"/><w:jc w:val="both"/>`);
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: five }))).pages[0].body);
    expect(ls).toHaveLength(1);
    const glyphs = ls[0].ops.flatMap((o) => o.xs.map((x, i) => ({ x: o.x + x, text: o.texts[i] })));
    expect(glyphs.filter((g) => g.text === 'a').at(-1)!.x + a).toBeCloseTo(72 + fiveWords - 2, 1);
  });

  it('wraps a word whose squeeze would change the spaces more than half the stretch without it', async () => {
    // Ten words and a one-letter word overflowing by 4pt: the ten spaces could give up 7.6pt, but
    // without the last word the nine spaces would stretch by only (a + space - 4) / 9 each.
    const width = 10 * 4 * a + 10 * space + a - 4;
    const body = p(Array(10).fill('aaaa').join(' ') + ' a', `<w:ind w:right="${Math.round((595.3 - 144 - width) * 20)}"/><w:jc w:val="both"/>`);
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body }))).pages[0].body);
    expect(ls.map((l) => l.text.trim().split(/\s+/).length)).toEqual([10, 1]);
  });

  it('only shrinks spaces in Word 2013+ compatibility mode', async () => {
    const word2010Settings = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>';
    const word2010 = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: word2010Settings, body: para('both') }))).pages[0].body);
    expect(word2010[0].text.match(/aaaa/g)).toHaveLength(4);
    // A package without a settings part opens in Word 2007 mode.
    const noSettings = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: para('both') }))).pages[0].body);
    expect(noSettings[0].text.match(/aaaa/g)).toHaveLength(4);
  });
});

describe('line breaking', () => {
  it('wraps fixed-width spaces like text instead of letting them hang past the margin', async () => {
    // Room for "aaaa " and one and a half em spaces: the two em spaces move to the next line
    // together, ahead of the "b" they precede.
    const width = 4 * ((1139 / 2048) * 11) + (569 / 2048) * 11 + 1.5 * 11;
    const body = p('aaaa \u2003\u2003b', `<w:ind w:right="${Math.round((595.3 - 144 - width) * 20)}"/>`);
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    const b = ls.at(-1)!.ops.flatMap((o) => o.xs.map((x, i) => ({ x: o.x + x, text: o.texts[i] }))).find((g) => g.text === 'b')!;
    expect(ls).toHaveLength(2);
    expect(b.x).toBeCloseTo(72 + 22, 1);
  });
  it('lets text on a right tab hang into the right indent, and from Word 2013 stops it at the margin', async () => {
    const para = (tab: number) => `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="${tab}"/></w:tabs><w:ind w:right="1134"/></w:pPr><w:r><w:t>Left</w:t><w:tab/><w:t>99</w:t></w:r></w:p>`;
    const end = async (mode: number, tab: number) => {
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(mode), body: para(tab) }))).pages[0].body);
      expect(ls.map((l) => l.text)).toEqual(['Left99']);
      const last = ls[0].ops.at(-1)!;
      return last.x + last.xs.at(-1)! + (1139 / 2048) * 11;
    };
    // A contents line: the right indent wraps the title, while the page number ends at the tab.
    expect(await end(14, 9026)).toBeCloseTo(72 + 9026 / 20, 1);
    expect(await end(15, 9026)).toBeCloseTo(72 + 9026 / 20, 1);
    // A tab past the margin: Word 2013 ends the text at the margin, Word 2010 at the tab.
    expect(await end(15, 9200)).toBeCloseTo(595.3 - 72, 1);
    expect(await end(14, 9200)).toBeCloseTo(72 + 9200 / 20, 1);
  });

  it('keeps a line of nothing but tabs whole past the margin, but wraps tabs that follow text', async () => {
    const tabs = (text: string) => `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="9300"/></w:tabs></w:pPr><w:r>${text ? `<w:t>${text}</w:t>` : ''}<w:tab/><w:tab/></w:r></w:p>`;
    const body = p('Before') + tabs('') + p('Middle') + tabs('Text') + p('After');
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(14), body }))).pages[0].body);
    const y = (t: string) => ls.find((l) => l.text === t)!.y;
    expect(y('Middle') - y('Before')).toBeCloseTo(2 * ARIAL_11, 1);
    // After "Text" the first tab, past the margin, starts a line that keeps the second.
    expect(y('After') - y('Middle')).toBeCloseTo(3 * ARIAL_11, 1);
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

  it('formats a page number like its result, not like the field characters', async () => {
    const big = '<w:rPr><w:sz w:val="48"/></w:rPr>';
    const body =
      `<w:p><w:r>${big}<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r>${big}<w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:t>1</w:t></w:r><w:r>${big}<w:fldChar w:fldCharType="end"/></w:r></w:p>${p('Next')}`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    expect(ls[0].ops[0].size).toBe(11);
    expect(ls[1].y - ls[0].y).toBeCloseTo(ARIAL_11, 2);
  });

  it('shows the selected entry of a legacy drop-down form field', async () => {
    const dropdown = (result: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:ddList>${result}<w:listEntry w:val="Red"/><w:listEntry w:val="Green"/></w:ddList></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body = `<w:p>${dropdown('<w:result w:val="1"/>')}</w:p><w:p>${dropdown('')}</w:p>`;
    expect(pageText((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0])).toBe('Green\nRed');
  });

  it('draws legacy form checkboxes in their saved state', async () => {
    const box = (state: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:checkBox><w:sizeAuto/>${state}</w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body = `<w:p>${box('<w:default w:val="0"/>')}<w:r><w:t xml:space="preserve"> No </w:t></w:r>${box('<w:default w:val="1"/>')}<w:r><w:t xml:space="preserve"> Yes</w:t></w:r></w:p>`;
    expect(pageText((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0])).toBe('\u2610 No \u2612 Yes');
  });

  it('fills in an empty table of contents only when the document asks Word to update its fields', async () => {
    const toc = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const heading = (text: string, level: number) => p(text, `<w:outlineLvl w:val="${level}"/>`);
    const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const body = toc + heading('One', 0) + pageBreak + heading('Two', 1) + p('Body text');
    const first = async (settings: string) => lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15) + settings, body }))).pages[0].body);
    const filled = await first('<w:updateFields w:val="true"/>');
    // Each entry shows the page its heading lands on, at the right margin; a level-2 entry is
    // indented by the 11pt Normal font size.
    expect(filled.map((l) => [l.text.replace(/\.+/, ' … '), Math.round(l.x)])).toEqual([['One … 1', 72], ['Two … 2', 83], ['One', 72]]);
    expect((await first('')).map((l) => l.text)).toEqual(['One']);
  });

  it('shows the data a content control is bound to, or its placeholder when the data is empty', async () => {
    const store = '{A1B2C3D4-0000-0000-0000-000000000001}';
    const core = '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}';
    const control = (storeId: string, xpath: string, cached: string, pr = '') =>
      `<w:p><w:sdt><w:sdtPr>${pr}<w:dataBinding w:xpath="${xpath}" w:storeItemID="${storeId}"/></w:sdtPr><w:sdtContent><w:r><w:t>${cached}</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
    const body =
      control(store, '/ns0:root[1]/ns0:Company[1]', 'Old') +
      control(core, '/ns1:coreProperties[1]/ns0:creator[1]', 'Wei', '<w:placeholder><w:docPart w:val="Author"/></w:placeholder>') +
      control(store, '/ns0:root[1]/ns0:Gone[1]', 'Kept') +
      control(store, '/ns0:root[1]/ns0:Color[1]', 'r', '<w:dropDownList><w:listItem w:displayText="Red" w:value="r"/></w:dropDownList>');
    const docx = makeDocx({
      styles: ARIAL_STYLES,
      body,
      customXml: { [store]: '<root xmlns="urn:test"><Company>Contoso</Company><Color>r</Color></root>' },
      coreProperties: '<dc:creator></dc:creator>',
      glossary: '<w:docPart><w:docPartPr><w:name w:val="Author"/></w:docPartPr><w:docPartBody><w:p><w:r><w:t>Your name</w:t></w:r></w:p></w:docPartBody></w:docPart>',
    });
    expect(pageText((await layout(docx)).pages[0])).toBe('Contoso\nYour name\nKept\nRed');
  });
});

describe('run formatting', () => {
  it('draws no underline for a w:u that only sets a colour', async () => {
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: p('Plain', '', '<w:u w:color="FF0000"/>') + p('Under', '', '<w:u w:val="single"/>') }));
    const [, under] = lines(r.pages[0].body);
    const rules = r.pages[0].body.filter((o) => o.kind === 'line');
    expect(rules).toHaveLength(1);
    expect(rules[0].kind === 'line' && rules[0].y1).toBeGreaterThan(under.y);
  });
});
