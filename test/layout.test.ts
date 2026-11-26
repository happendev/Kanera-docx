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
import { formatNumber } from '../src/layout/chart.js';
import { applyColorMods } from '../src/ooxml/values.js';
import type { DrawOp as DrawOpLike, PathOp, ShadeOp } from '../src/layout/types.js';
import { convertDocx } from '../src/node.js';
import type { DrawingInline, Paragraph } from '../src/ooxml/model.js';
import { ghostscript, ghostscriptPixels, pdfStreams } from './helpers/pdf.js';
import { solidPixels } from '../src/image/pixels.js';
import { ImageBudget } from '../src/image/budget.js';

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

  it('evens out the columns of a section that a continuous section break ends', async () => {
    const sect = (cols: number) => `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:num="${cols}" w:space="708"/></w:sectPr>`;
    const items = Array.from({ length: 6 }, (_, i) => p(`Item ${i + 1}`)).join('');
    const body = p('Before', sect(1)) + items + p('', sect(2)) + p('After');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body, sectPr: sect(1) }));
    const ls = lines(r.pages[0].body);
    // Three items in each column, side by side, and the next section below them.
    expect(ls.map((l) => l.text)).toEqual(['Before', 'Item 1Item 4', 'Item 2Item 5', 'Item 3Item 6', 'After']);
    expect(ls[4].y - ls[3].y).toBeLessThan(2 * ARIAL_11);
  });

  it('lays a picture turned on its side out in its turned box', async () => {
    const turned = inlineImage('img1', 50, 100).replace('<a:xfrm>', '<a:xfrm rot="5400000">');
    const body = p('Above') + `<w:p>${turned}</w:p>` + p('Below');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body, media: { img1: { name: 'red.png', data: redPng() } } }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    // 100pt wide and 50pt tall once turned, as in Word.
    expect(ys[1] - ys[0]).toBeCloseTo(ARIAL_11 + 50, 0);
  });

  describe('a line holding only pictures', () => {
    const media = { img1: { name: 'red.png', data: redPng() } };
    const em = (units: number) => units / 2048;
    // The picture's height (pt) in a run of `run` half-points, under a mark of `mark` half-points.
    const pictureLine = async (h: number, run: number, mark: number, spacing = '') => {
      const pic = inlineImage('img1', 20, h).replace('<w:r><w:drawing>', `<w:r><w:rPr><w:sz w:val="${run}"/></w:rPr><w:drawing>`);
      const body = p('Above') + `<w:p><w:pPr>${spacing}<w:rPr><w:sz w:val="${mark}"/></w:rPr></w:pPr>${pic}</w:p>` + p('Below');
      const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body).map((l) => l.y);
      return ys[1] - ys[0] - ARIAL_11;
    };

    it("is as tall as its pictures' run text, not the mark's", async () => {
      // A 2pt picture in a 12pt run under an 8pt mark: the run's text height, and the mark's line gap.
      expect(await pictureLine(2, 24, 16)).toBeCloseTo(em(1854 + 434) * 12 + em(67) * 8, 1);
    });

    it("scales the shorter of the run's and the mark's text heights by multiple spacing", async () => {
      const half = '<w:spacing w:line="360" w:lineRule="auto"/>';
      // A 10pt run under a 14pt mark adds half the run's height; a 14pt run under a 10pt mark half the mark's.
      expect(await pictureLine(50, 20, 28, half)).toBeCloseTo(50 + em(67) * 14 + 0.5 * (em(1854 + 434) * 10 + em(67) * 14), 1);
      expect(await pictureLine(50, 28, 20, half)).toBeCloseTo(50 + em(67) * 10 + 0.5 * (em(1854 + 434) * 10 + em(67) * 10), 1);
    });

    it('sits its pictures at the bottom of the line, with no descent below them', async () => {
      // A 4pt picture in an 11pt line: the next line's ascent and line gap are all that lie below it.
      const body = p('Above') + `<w:p>${inlineImage('img1', 20, 4)}</w:p>` + p('Below');
      const page = (await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0];
      const img = page.body.find((o) => o.kind === 'image');
      expect(img && img.y + img.h).toBeCloseTo(lines(page.body)[1].y - em(1854 + 67) * 11, 1);
    });
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

  it('spaces an auto-spaced paragraph by its stated spacing when HTML auto spacing is turned off', async () => {
    const auto = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="0"/>';
    const gap = async (settings: string) => {
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: p('First') + p('Second', auto) }))).pages[0].body);
      return ls[1].y - ls[0].y;
    };
    expect(await gap('<w:compat/>')).toBeCloseTo(ARIAL_11 + 14, 1);
    expect(await gap('<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>')).toBeCloseTo(ARIAL_11 + 5, 1);
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

  it("reaches the paragraph box 1.5pt past the text, and runs side borders through the space after", async () => {
    const left = '<w:pBdr><w:left w:val="single" w:sz="8" w:space="4" w:color="FF0000"/></w:pBdr><w:spacing w:after="200"/>';
    const body = p('Ruled', left) + p('Shaded', '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>');
    const ops = (await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body;
    const rule = ops.find((o) => o.kind === 'line');
    const fill = ops.find((o) => o.kind === 'rect');
    if (rule?.kind !== 'line' || fill?.kind !== 'rect') throw new Error('missing decorations');
    // 4pt of space, the 1pt line's centre, and Word's 1.5pt beyond the 72pt margin.
    expect(rule.x1).toBeCloseTo(72 - 4 - 0.5 - 1.5, 2);
    // The line, then the 10pt space after: there is no bottom border to close the box above it.
    expect(rule.y2 - rule.y1).toBeCloseTo(ARIAL_11 + 10, 1);
    expect(fill.x).toBeCloseTo(72 - 1.5, 2);
    expect(fill.w).toBeCloseTo(841.9 / 1.414 - 144 + 3, 0);
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

  it('adds no blank page for a page break whose line does not fit below a full page', async () => {
    const full = Array.from({ length: linesPerPage }, (_, i) => p(`Line ${i + 1}`)).join('');
    const body = full + '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Next</w:t></w:r></w:p>';
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body }));
    expect(r.pages).toHaveLength(2);
    expect(pageText(r.pages[1])).toBe('Next');
    // A paragraph of nothing but the break takes its line on the next page, which stays blank.
    const alone = await layout(makeDocx({ styles: ARIAL_STYLES, body: full + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p('Next') }));
    expect(alone.pages).toHaveLength(3);
    expect(pageText(alone.pages[2])).toBe('Next');
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

  it('hangs horizontal borders from their edge, draws the heavier of two that meet, and ends the table on its last content', async () => {
    const borders = (top: string, bottom: string) => `<w:tcBorders><w:top ${top} w:space="0" w:color="000000"/><w:bottom ${bottom} w:space="0" w:color="000000"/></w:tcBorders>`;
    const double = 'w:val="double" w:sz="6"';
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('A', 3000, borders(double, double))}</w:tr><w:tr>${cell('B', 3000, borders('w:val="single" w:sz="4"', 'w:val="thick" w:sz="18"'))}</w:tr></w:tbl>`;
    const body = (await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: tbl + p('After') }))).pages[0].body;
    const ys = new Map(lines(body).map((l) => [l.text, l.y]));
    const strokes = body.flatMap((o) => (o.kind === 'line' ? [o] : []));
    // The double between the rows outweighs the second row's single top border, which is not drawn.
    expect(strokes.some((o) => o.width === 0.5)).toBe(false);
    const doubles = [...new Set(strokes.filter((o) => o.width === 0.75).map((o) => o.y1))].sort((a, b) => a - b);
    // Each double starts on its edge and its second line lies 1.5pt below.
    expect(doubles).toHaveLength(4);
    expect(doubles[1] - doubles[0]).toBeCloseTo(1.5, 2);
    expect(doubles[2] - doubles[0]).toBeCloseTo(ys.get('B')! - ys.get('A')!, 2);
    // The 2.25pt bottom border starts 0.375pt above the last row's content edge; the table keeps
    // room for it before the next paragraph.
    const thick = strokes.find((o) => o.width === 2.25)!;
    expect(ys.get('After')! - ys.get('B')!).toBeCloseTo(ARIAL_11 + 2.25, 2);
    expect(thick.y1 - doubles[2]).toBeCloseTo(2.25 + ARIAL_11 + 0.75, 2);
  });

  it('ends a thick bottom border at the side borders’ outer edges, and the side borders and shading at its foot', async () => {
    const edge = (side: string, sz: number) => `<w:${side} w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/>`;
    const tcPr = `<w:tcBorders>${edge('left', 4)}${edge('bottom', 18)}${edge('right', 4)}</w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="DDEEDD"/>`;
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell('A', 3000, tcPr)}${cell('B', 3000, tcPr)}</w:tr></w:tbl>`;
    const body = (await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: tbl + p('After') }))).pages[0].body;
    const strokes = body.flatMap((o) => (o.kind === 'line' ? [o] : []));
    const thick = strokes.filter((o) => o.width === 2.25);
    const sides = strokes.filter((o) => o.width === 0.5);
    const xs = sides.map((o) => o.x1);
    // The thick border spans the side borders' outer edges and no further.
    expect(Math.min(...thick.map((o) => Math.min(o.x1, o.x2)))).toBeCloseTo(Math.min(...xs) - 0.25, 3);
    expect(Math.max(...thick.map((o) => Math.max(o.x1, o.x2)))).toBeCloseTo(Math.max(...xs) + 0.25, 3);
    // The side borders and the shading stop at its foot rather than show below it.
    const foot = thick[0].y1 + 2.25 / 2;
    for (const o of sides) expect(Math.max(o.y1, o.y2)).toBeCloseTo(foot, 3);
    for (const o of body) if (o.kind === 'rect' && o.fill) expect(o.y + o.h).toBeCloseTo(foot, 3);
    expect(thick.every((o) => o.cap === 'butt') && sides.every((o) => o.cap === 'butt')).toBe(true);
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

  it('turns cell text that runs up or down the row, and grows the row to fit its line', async () => {
    const turn = (dir: string) => `<w:textDirection w:val="${dir}"/>`;
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('Plain', 2000)}${cell('Running up the row', 2000, turn('btLr'))}${cell('Down', 2000, turn('tbRl'))}</w:tr></w:tbl>`;
    const body = (await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: tbl + p('After') }))).pages[0].body;
    const turns = body.flatMap((o) => (o.kind === 'transform' ? [o.matrix] : []));
    // Up: tops of the lines to the left, from the cell's bottom; down: tops to the right, from its top.
    expect(turns.map((m) => m.slice(0, 4))).toEqual([
      [0, -1, 1, 0],
      [0, 1, -1, 0],
    ]);
    // The up line starts at the second cell's left text edge, the down line at the third's right
    // (these styles give cells no side margins).
    expect(turns[1][4] - turns[0][4]).toBeCloseTo(200, 1);
    const ls = lines(body);
    const y = (t: string) => ls.find((l) => l.text === t)!.y;
    // One unwrapped line of the longer text, about 100pt, sets the row's height.
    expect(y('After') - y('Plain')).toBeGreaterThan(90);
    expect(turns[0][5] - turns[1][5]).toBeGreaterThan(90);
  });

  it("adds the row's bottom margin to an exact row height and starts its cells at its largest top margin", async () => {
    const mar = (top: number, bottom: number) => `<w:tcMar><w:top w:w="${top}" w:type="dxa"/><w:bottom w:w="${bottom}" w:type="dxa"/></w:tcMar>`;
    const exact = '<w:trPr><w:trHeight w:val="1000" w:hRule="exact"/></w:trPr>';
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${exact}${cell('Left', 3000, mar(40, 200))}${cell('Right', 4000, mar(200, 0))}</w:tr><w:tr>${exact}${cell('Next', 3000, mar(0, 0))}${cell('Other', 4000, mar(0, 0))}</w:tr></w:tbl>`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: tbl }))).pages[0].body);
    const y = (t: string) => ls.find((l) => l.text.includes(t))!.y;
    expect(y('Left')).toBeCloseTo(y('Right'), 2);
    // The 50pt row, then its 10pt bottom margin; the next row's cells have no top margin.
    expect(y('Next') - y('Left')).toBeCloseTo(50 + 10 - 10, 1);
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

  it('moves a row on whole where its minimum height does not fit, though its content could split there', async () => {
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    const filler = Array.from({ length: linesPerPage - 3 }, (_, i) => p(`Line ${i + 1}`)).join('');
    // Three lines of room are left: less than the 60pt minimum height, and the row's ten lines overflow anyway.
    const paras = Array.from({ length: 10 }, (_, i) => p(`Cell ${i + 1}`)).join('');
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${paras}</w:tc></w:tr></w:tbl>`;
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: filler + tbl }));
    expect(lines(r.pages[0].body).at(-1)!.text).toBe(`Line ${linesPerPage - 3}`);
    expect(lines(r.pages[1].body)[0].text).toBe('Cell 1');
  });

  it('keeps a footnote on the page of the table row that refers to it', async () => {
    const ref = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
    const footnotes = `<w:footnote w:id="1">${p('The note')}</w:footnote>`;
    const row = (i: number, k: number) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Row ${i}</w:t></w:r>${i === k ? ref : ''}</w:p></w:tc></w:tr>`;
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    // Near the page's end the row fits but its note doesn't: the row moves on with its note.
    for (let k = linesPerPage - 6; k <= linesPerPage; k++) {
      const rows = Array.from({ length: 2 * linesPerPage }, (_, i) => row(i + 1, k)).join('');
      const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;
      const r = await layout(makeDocx({ styles: ARIAL_STYLES, footnotes, body: tbl + p('') }));
      const page = r.pages.findIndex((pg) => lines(pg.body).some((l) => /^Row (\d+)/.exec(l.text)?.[1] === String(k)));
      expect(pageText(r.pages[page]), `row ${k}`).toContain('The note');
    }
  });

  it('splits a row that continues a vertical merge when the merged content overflows', async () => {
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    const filler = Array.from({ length: linesPerPage - 6 }, (_, i) => p(`Line ${i + 1}`)).join('');
    const tc = (vMerge: string, body: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge${vMerge}/></w:tcPr>${body}</w:tc>`;
    const paras = (label: string, n: number) => Array.from({ length: n }, (_, i) => p(`${label} ${i + 1}`)).join('');
    // The second row's own cell fits in the room left, but the cell merged down into it does not.
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>`
      + `<w:tr>${tc(' w:val="restart"', paras('Merged', 10))}${cell('Top', 3000)}</w:tr>`
      + `<w:tr>${tc('', '<w:p/>')}<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${paras('Own', 3)}</w:tc></w:tr></w:tbl>`;
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: filler + tbl }));
    const text = (page: number) => lines(r.pages[page].body).map((l) => l.text).join(' ');
    expect(text(0)).toContain('Own 1');
    expect(text(1)).not.toContain('Own 1');
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

  it("leaves a cell's orphaned first line before Word 2013, and moves its row from Word 2013", async () => {
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    const filler = Array.from({ length: linesPerPage - 1 }, (_, i) => p(`Line ${i + 1}`)).join('');
    // Room is left for the 1pt border and one line of a four-line statement beside an empty cell.
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr>${cell(words(8), 1500)}${cell('', 1500)}</w:tr></w:tbl>`;
    const word2010 = await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(14), body: filler + tbl }));
    expect(lines(word2010.pages[0].body).at(-1)!.text.trim()).toBe('word0 word1');
    expect(lines(word2010.pages[1].body)).toHaveLength(3);
    const word2013 = await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(15), body: filler + tbl }));
    expect(lines(word2013.pages[0].body).at(-1)!.text).toBe(`Line ${linesPerPage - 1}`);
    expect(lines(word2013.pages[1].body)).toHaveLength(4);
  });

  it("leaves room for a split row's top border on the page the row starts on", async () => {
    const linesPerPage = Math.floor(BODY_HEIGHT / ARIAL_11);
    const filler = Array.from({ length: linesPerPage - 2 }, (_, i) => p(`Line ${i + 1}`)).join('');
    // Two lines and 2.4pt of room are left, so below the 3pt top border only one line fits.
    const top = '<w:tcBorders><w:top w:val="single" w:sz="24" w:space="0" w:color="000000"/></w:tcBorders>';
    const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${top}</w:tcPr>${p('One')}${p('Two')}${p('Three')}</w:tc></w:tr></w:tbl>`;
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: filler + tbl }));
    expect(lines(r.pages[0].body).at(-1)!.text).toBe('One');
    expect(lines(r.pages[1].body)[0].text).toBe('Two');
  });

  it("widens an autofit column to its longest word while the table fits, and breaks the word otherwise", async () => {
    // A file Word saved (it has revision ids) keeps its grid; only too-narrow columns widen.
    const settings = '<w:rsids><w:rsidRoot w:val="00000001"/></w:rsids>';
    const texts = async (text: string, tblPr = '') => {
      const tbl = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${tblPr}</w:tblPr><w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="600"/></w:tblGrid><w:tr>${cell(text, 600)}${cell('B', 600)}</w:tr></w:tbl>`;
      return lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: tbl }))).pages[0].body).map((l) => l.text);
    };
    // "Unbreakable" is wider than its 30pt column.
    expect(await texts('Unbreakable')).toEqual(['UnbreakableB']);
    expect((await texts('Unbreakable', '<w:tblLayout w:type="fixed"/>')).length).toBeGreaterThan(1);
    expect((await texts('W'.repeat(60))).length).toBeGreaterThan(1);
  });

  it('insets every cell by the cell spacing', async () => {
    const settings = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';
    const run = async (tblPr: string) => lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings, body: table([['A1', 'B1'], ['A2', 'B2']], tblPr) }))).pages[0].body);
    const plain = await run('');
    const spaced = await run('<w:tblCellSpacing w:w="40" w:type="dxa"/>');
    // Word's spacing value is half the gap: cells sit 4pt apart and 4pt in from the table's
    // edges. A spaced cell also holds its 1pt borders inside it, so rows are 5pt further apart
    // and text moves 5pt in.
    expect(spaced[0].ops[0].x - plain[0].ops[0].x).toBeCloseTo(5, 2);
    expect(spaced[1].y - spaced[0].y - (plain[1].y - plain[0].y)).toBeCloseTo(5, 2);
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

  it("centres cell text in the space below the row's top border", async () => {
    const rule = '<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="12" w:space="0" w:color="000000"/>';
    const row = (text: string) => `<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>${cell(text, 4000, '<w:vAlign w:val="center"/>')}</w:tr>`;
    const body = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${rule}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${row('Middle')}${row('Next')}</w:tbl>`;
    const first = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body)[0];
    // A 20pt row below a 1.5pt border.
    expect(first.y).toBeCloseTo(72 + 1.5 + (20 - ARIAL_11) / 2 + ARIAL_11 - (434 / 2048) * 11, 1);
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
    // A banner placed on the page and wrapped above and below does move a blank line, to 70pt here.
    const banner = anchoredImage('none', 451, 50, { h: 'margin', x: 0, v: 'page', y: 20 }, '<wp:wrapTopAndBottom/>');
    const below = await layout(makeDocx({ styles: ARIAL_STYLES, sectPr, headers: { hdr1: `<w:p>${banner}</w:p>` }, body: p('Body') }));
    expect(lines(below.pages[0].body)[0].y).toBeCloseTo(70 + ARIAL_11 + baseline, 1);
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

  /** Tight wrapping along an outline given as fractions of the object's extent. */
  const tight = (pts: Array<[number, number]>) =>
    `<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0">${pts.map(([x, y], i) => `<wp:${i ? 'lineTo' : 'start'} x="${x * 21600}" y="${y * 21600}"/>`).join('')}</wp:wrapPolygon></wp:wrapTight>`;
  const box = tight([[0, 0], [0, 1], [1, 1], [1, 0], [0, 0]]);
  const starts = (line: { ops: Array<{ x: number; xs: number[] }> }) => line.ops.flatMap((o) => o.xs.map((x) => o.x + x));

  it('wraps tight text along the outline and fills the gaps on both sides of a line', async () => {
    // The outline covers the middle half of a picture 150-250pt from the margin.
    const pic = anchoredImage('img1', 100, 60, { h: 'margin', x: 150, v: 'paragraph', y: 0 }, tight([[0.25, 0], [0.25, 1], [0.75, 1], [0.75, 0], [0.25, 0]]));
    const body = `<w:p>${pic}<w:r><w:t>${words(60)}</w:t></w:r></w:p>`;
    const xs = starts(lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body)[0]);
    expect(Math.min(...xs)).toBeCloseTo(72, 1);
    expect(Math.min(...xs.filter((x) => x > 72 + 175))).toBeCloseTo(72 + 225, 0);
  });

  it('wraps an outline beside the text of a line, not the space that "multiple" spacing adds below it', async () => {
    // The picture starts 15pt down: below the first line's 12.6pt of text but within its double-spaced line.
    const first = async (wrap: string) => {
      const pic = anchoredImage('img1', 100, 50, { h: 'margin', x: 0, v: 'paragraph', y: 15 }, wrap);
      const body = `<w:p><w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>${pic}<w:r><w:t>${words(40)}</w:t></w:r></w:p>`;
      return lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body)[0].x;
    };
    expect(await first(box)).toBeCloseTo(72, 1);
    expect(await first(square)).toBeCloseTo(172, 1);
  });

  it('moves a line with no room beside an outline down a line at a time', async () => {
    // A triangle over the column whose right side opens downwards: no room on the first line, 57pt
    // on the second. A box would send the line below the whole picture.
    const pic = anchoredImage('img1', 451.3, 100, { h: 'margin', x: 0, v: 'paragraph', y: 0 }, tight([[0, 0], [1, 0], [0, 1], [0, 0]]));
    const body = `<w:p>${pic}<w:r><w:t>${words(4)}</w:t></w:r></w:p>`;
    const first = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body)[0];
    expect(first.y).toBeCloseTo(72 + ARIAL_11 + ARIAL_11 - (434 / 2048) * 11, 1);
    expect(first.x).toBeCloseTo(72 + 451.3 * (1 - ARIAL_11 / 100), 0);
  });

  it('leaves a page break beside an object on its page', async () => {
    // The picture fills the rest of the page below "One", leaving no room for the break's line.
    const pic = anchoredImage('img1', 451.3, 680, { h: 'margin', x: 0, v: 'margin', y: 15 }, square);
    const body = `<w:p>${pic}<w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>` + p('Two');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body, media }));
    expect(r.pages).toHaveLength(2);
    expect(pageText(r.pages[1])).toBe('Two');
  });

  it('sizes a line of pictures ended by a line break as a line of pictures', async () => {
    const pic = inlineImage('img1', 40, 40);
    const next = async (body: string) => lines((await layout(makeDocx({ styles: ARIAL_STYLES, body, media }))).pages[0].body).find((l) => l.text === 'Next')!.y;
    expect(await next(`<w:p>${pic}<w:r><w:br/><w:t>Next</w:t></w:r></w:p>`)).toBeCloseTo(await next(`<w:p>${pic}</w:p>${p('Next')}`), 2);
  });

  it('aligns an object together with its effects', async () => {
    // A 10pt shadow on the right keeps inside the margin.
    const pic = anchoredImage('img1', 100, 40, { h: 'margin', x: 0, v: 'paragraph', y: 0 }, square)
      .replace('<wp:posOffset>0</wp:posOffset></wp:positionH>', '<wp:align>right</wp:align></wp:positionH>')
      .replace(square, `<wp:effectExtent l="0" t="0" r="${10 * 12700}" b="0"/>${square}`);
    const page = (await layout(makeDocx({ styles: ARIAL_STYLES, body: `<w:p>${pic}</w:p>`, media }))).pages[0];
    const img = [...page.behind, ...page.body, ...page.front].find((o) => o.kind === 'image');
    expect(img?.x).toBeCloseTo(595.3 - 72 - 10 - 100, 1);
  });

  it('draws a legacy callout from its shape type, with level text inside the type\'s text area', async () => {
    // An oval callout (type 63) turned 90°, its tail up and to the right. The shape states no fill
    // or outline, so it takes VML's white fill and black outline.
    const type = '<v:shapetype id="_x0000_t63" coordsize="21600,21600" o:spt="63" adj="1350,25920">' +
      '<v:path textboxrect="3163,3163,18437,18437"/></v:shapetype>';
    const shape = '<v:shape type="#_x0000_t63" style="width:100pt;height:80pt;rotation:90" adj="20574,-7904">' +
      `<v:textbox><w:txbxContent>${p(words(40))}</w:txbxContent></v:textbox></v:shape>`;
    const page = (await layout(makeDocx({ styles: ARIAL_STYLES, body: `<w:p><w:r><w:pict>${type}${shape}</w:pict></w:r></w:p>` }))).pages[0];
    const ops = [...page.behind, ...page.body, ...page.front];
    expect(ops.some((o) => o.kind === 'path' && o.stroke === '000000')).toBe(true);
    // The text is laid out once the rotation ends, clipped to the type's text area.
    const turn = ops.findIndex((o) => o.kind === 'transform');
    const end = ops.findIndex((o, i) => i > turn && o.kind === 'restore');
    const area = ops.find((o, i) => i > end && o.kind === 'clip');
    expect(turn).toBeGreaterThanOrEqual(0);
    expect(ops.findIndex((o) => o.kind === 'glyphs')).toBeGreaterThan(end);
    if (area?.kind !== 'clip') throw new Error('no text area');
    expect(area.w).toBeCloseTo((100 * (18437 - 3163)) / 21600, 1);
    // The fill starts at the tail's tip, given in the shape's 21600 units.
    const fill = ops.find((o): o is PathOp => o.kind === 'path' && o.fill === 'FFFFFF');
    const [tipX, tipY] = fill!.segs[0].pts;
    expect(tipX - area.x).toBeCloseTo((100 * (20574 - 3163)) / 21600, 1);
    expect(tipY - area.y).toBeCloseTo((80 * (-7904 - 3163)) / 21600, 1);
  });

  it("draws a shape's theme shadow blurred, at 300 dpi, under the shape and its outline, as Word does", async () => {
    const theme =
      '<a:fmtScheme name="Office"><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst>' +
      '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>' +
      '</a:effectLst></a:effectStyle></a:effectStyleLst></a:fmtScheme>';
    const shape = (rot: string) =>
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><a:graphic><a:graphicData>' +
      `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:xfrm${rot}><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"/>` +
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
      '<wps:style><a:lnRef idx="1"><a:srgbClr val="000000"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef><a:effectRef idx="2"><a:srgbClr val="000000"/></a:effectRef><a:fontRef idx="minor"/></wps:style>' +
      '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const draw = async (rot = '') => {
      const ops = (await layout(makeDocx({ theme, body: shape(rot) }))).pages[0].body;
      const at = ops.findIndex((o) => o.kind === 'image' && o.image.id.startsWith('shadow-'));
      const fill = ops.find((o): o is PathOp => o.kind === 'path' && o.fill === 'FF0000')!;
      return { ops, at, image: ops[at] as Extract<DrawOpLike, { kind: 'image' }>, fill };
    };
    const { ops, at, image, fill } = await draw();
    // Behind the shape, and not cut out from under it.
    expect(at).toBeGreaterThanOrEqual(0);
    expect(ops.indexOf(fill)).toBeGreaterThan(at);
    expect(ops.slice(0, at).some((o) => o.kind === 'clip')).toBe(false);
    const xs = fill.segs.flatMap((q) => q.pts.filter((_, i) => i % 2 === 0));
    const ys = fill.segs.flatMap((q) => q.pts.filter((_, i) => i % 2 === 1));
    // 1.81pt down; reaching the 3.15pt blur past the 1pt outline's outer edge.
    expect(image.x).toBeCloseTo(Math.min(...xs) - 0.5 - 3.15, 2);
    expect(image.y).toBeCloseTo(Math.min(...ys) + 1.811 - 0.5 - 3.15, 2);
    const pixels = solidPixels(image.image.solid!, image.image.data, new ImageBudget());
    expect(pixels.width).toBe(Math.ceil((100 + 1 + 6.3) * (300 / 72)));
    // 35% black inside; half of it on the outline's outer edge.
    const row = Math.floor(pixels.height / 2) * pixels.width;
    expect(pixels.data[4 * (row + Math.floor(pixels.width / 2)) + 3]).toBe(89);
    const edge = Math.round(3.15 * (300 / 72));
    expect(pixels.data[4 * (row + edge) + 3]).toBeGreaterThan(35);
    expect(pixels.data[4 * (row + edge) + 3]).toBeLessThan(55);
    // A shape turned a quarter keeps its shadow below it on the page: in its own turned frame, to its right.
    const turned = await draw(' rot="5400000"');
    const local = turned.image.y - Math.min(...turned.fill.segs.flatMap((q) => q.pts.filter((_, i) => i % 2 === 1)));
    expect(local).toBeCloseTo(-0.5 - 3.15, 2);
    expect(turned.image.x - Math.min(...turned.fill.segs.flatMap((q) => q.pts.filter((_, i) => i % 2 === 0)))).toBeCloseTo(1.811 - 0.5 - 3.15, 2);
  });

  it('fits a text box set to resize to its text to the text, whatever height it was saved at', async () => {
    const box = (bodyPr: string) =>
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="1397000"/><a:graphic><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:prstGeom prst="rect"/><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
      `<wps:txbx><w:txbxContent>${p('Fit')}</w:txbxContent></wps:txbx><wps:bodyPr>${bodyPr}</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const next = async (bodyPr: string) => lines((await layout(makeDocx({ styles: ARIAL_STYLES, body: box(bodyPr) + p('Next') }))).pages[0].body).find((l) => l.text === 'Next')!.y;
    // Saved 110pt tall, the box shrinks to its line, the 3.6pt insets above and below it and its
    // 0.75pt outline, which Word counts the insets from inside of.
    expect((await next('')) - (await next('<a:spAutoFit/>'))).toBeCloseTo(110 - (ARIAL_11 + 7.2 + 0.75), 1);
  });

  it('fits a text box to its page count once the count is known', async () => {
    // "Pages 1" (40.4pt in Arial 11) fits the 42.25pt text width; "Pages 12" (46.5pt) wraps.
    const text = '<w:p><w:r><w:t xml:space="preserve">Pages </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>';
    const box = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="728980" cy="254000"/><a:graphic><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:prstGeom prst="rect"/><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
      `<wps:txbx><w:txbxContent>${text}</w:txbxContent></wps:txbx><wps:bodyPr><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const pages = Array.from({ length: 11 }, (_, i) => `<w:p><w:r><w:br w:type="page"/><w:t>Page ${i + 2}</w:t></w:r></w:p>`).join('');
    const r = await layout(makeDocx({ styles: ARIAL_STYLES, body: box + pages }));
    expect(r.pages).toHaveLength(12);
    const outline = r.pages[0].body.find((o): o is PathOp => o.kind === 'path' && o.stroke === '000000')!;
    const ys = outline.segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 1));
    expect(Math.max(...ys) - Math.min(...ys)).toBeCloseTo(2 * ARIAL_11 + 7.2 + 0.75, 1);
  });

  it('makes room for a character border around its text, beside and above and below it', async () => {
    const bdr = '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="FF0000"/>';
    const run = (text: string, rPr = '') => `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    // "Bo" and "x" are runs in the same border: one box, with room for the 0.5pt border at its ends.
    const page = async (border: string) =>
      (await layout(makeDocx({ styles: ARIAL_STYLES, body: `<w:p>${run('Bo', border)}${run('x', border)}${run(' after')}</w:p>${p('Next')}` }))).pages[0].body;
    const [boxed, plain] = [await page(bdr), await page('')];
    const glyphX = (body: DrawOpLike[], text: string) => {
      const op = glyphOps(body).find((o) => o.texts.join('').includes(text))!;
      return op.x + op.xs[op.texts.indexOf(text)];
    };
    expect(glyphX(boxed, 'B') - glyphX(plain, 'B')).toBeCloseTo(0.5, 5);
    expect(glyphX(boxed, 'x') - glyphX(plain, 'x')).toBeCloseTo(0.5, 5);
    expect(glyphX(boxed, 'a') - glyphX(plain, 'a')).toBeCloseTo(1, 5);
    // The line grows by the border's width above and below the text.
    const next = (body: DrawOpLike[]) => lines(body).find((l) => l.text === 'Next')!.y;
    expect(next(boxed) - next(plain)).toBeCloseTo(1, 5);
    // One box: its two upright sides at its ends, none between the runs.
    const sides = boxed.filter((o) => o.kind === 'line' && o.color === 'FF0000' && o.x1 === o.x2);
    expect(sides).toHaveLength(2);
  });

  it("keeps a shape's text inside its themed outline and starts text too tall for the box at its top", async () => {
    // Office 2007's line styles: a style's line reference 2 is 2pt wide.
    const theme = '<a:fmtScheme name="T"><a:lnStyleLst><a:ln w="9525"/><a:ln w="25400"/><a:ln w="38100"/></a:lnStyleLst></a:fmtScheme>';
    const black = '<a:srgbClr val="000000"/>';
    const box = (ln: string, anchor: string, after: number) =>
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="254000"/><a:graphic><a:graphicData>' +
      `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:prstGeom prst="rect"/>${ln}</wps:spPr>` +
      `<wps:style><a:lnRef idx="2">${black}</a:lnRef><a:fillRef idx="0">${black}</a:fillRef><a:effectRef idx="0">${black}</a:effectRef><a:fontRef idx="minor">${black}</a:fontRef></wps:style>` +
      `<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:after="${after}"/></w:pPr><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
      `<wps:bodyPr anchor="${anchor}"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const draw = async (ln: string, anchor = 't', after = 0) => {
      const body = (await layout(makeDocx({ styles: ARIAL_STYLES, theme, body: box(ln, anchor, after) }))).pages[0].body;
      return { text: lines(body).find((l) => l.text === 'Box')!, outline: body.find((o): o is PathOp => o.kind === 'path' && o.stroke === '000000') };
    };
    const themed = await draw(`<a:ln><a:solidFill>${black}</a:solidFill></a:ln>`);
    const bare = await draw('<a:ln><a:noFill/></a:ln>');
    // An outline that gives no width takes its style's, and the text keeps clear of half of it.
    expect(themed.outline?.strokeWidth).toBeCloseTo(2, 5);
    expect(bare.outline).toBeUndefined();
    expect(themed.text.x - bare.text.x).toBeCloseTo(1, 5);
    expect(themed.text.y - bare.text.y).toBeCloseTo(1, 5);
    // A line and 20pt of space after it overflow the 20pt box: centred, it starts at the top all the same.
    expect((await draw(`<a:ln><a:solidFill>${black}</a:solidFill></a:ln>`, 'ctr', 400)).text.y).toBeCloseTo(themed.text.y, 5);
  });

  it("draws a custom shape's quadratic curve as the same curve, skipping a malformed one", () => {
    const geom = '<a:custGeom><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
      '<a:quadBezTo><a:pt x="50" y="100"/><a:pt x="100" y="0"/></a:quadBezTo><a:quadBezTo><a:pt x="10" y="10"/></a:quadBezTo></a:path></a:pathLst></a:custGeom>';
    const body = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="1270000"/><a:graphic><a:graphicData>' +
      `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr>${geom}</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const para = parseDocx(makeDocx({ body }), nodeCompression.inflateRaw).body[0] as Paragraph;
    const d = (para.content.find((i) => i.type === 'drawing') as DrawingInline).drawing;
    // Control points two thirds of the way from each end to the quadratic's (50, 100).
    expect(d.path![0].segs).toEqual([
      { op: 'M', pts: [0, 0] },
      { op: 'C', pts: [expect.closeTo(1 / 3, 6), expect.closeTo(2 / 3, 6), expect.closeTo(2 / 3, 6), expect.closeTo(2 / 3, 6), 1, 0] },
    ]);
  });

  it("draws a custom shape's arc on its ellipse, its angles those of rays from the centre", () => {
    const segs = (path: string) => {
      const geom = `<a:custGeom><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="200" y="50"/></a:moveTo>${path}</a:path></a:pathLst></a:custGeom>`;
      const body = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2540000" cy="1270000"/><a:graphic><a:graphicData>' +
        `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr>${geom}</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
      const para = parseDocx(makeDocx({ body }), nodeCompression.inflateRaw).body[0] as Paragraph;
      return (para.content.find((i) => i.type === 'drawing') as DrawingInline).drawing.path![0].segs;
    };
    const close = (...v: number[]) => v.map((n) => expect.closeTo(n, 4));
    // Half way round the ellipse centred on (100, 50) from its right end, a quarter turn per curve, then on from where it ends.
    const k = (4 / 3) * Math.tan(Math.PI / 8) * 0.5;
    expect(segs('<a:arcTo wR="100" hR="50" stAng="0" swAng="10800000"/><a:lnTo><a:pt x="100" y="0"/></a:lnTo>')).toEqual([
      { op: 'M', pts: [1, 0.5] },
      { op: 'C', pts: close(1, 0.5 + k, 0.5 + k, 1, 0.5, 1) },
      { op: 'C', pts: close(0.5 - k, 1, 0, 0.5 + k, 0, 0.5) },
      { op: 'L', pts: [0.5, 0] },
    ]);
    // 45° ends where the ray at 45° meets the ellipse, (100 + 20√5, 50 + 20√5), not at its parameter's 45°.
    const eighth = segs('<a:arcTo wR="100" hR="50" stAng="0" swAng="2700000"/>');
    expect(eighth).toHaveLength(2);
    expect(eighth[1].pts.slice(4)).toEqual(close((100 + 20 * Math.sqrt(5)) / 200, (50 + 20 * Math.sqrt(5)) / 100));
    // Anticlockwise, a negative sweep goes up through the top.
    expect(segs('<a:arcTo wR="100" hR="50" stAng="0" swAng="-5400000"/>')[1].pts.slice(4)).toEqual(close(0.5, 0));
  });

  it("reads a VML fill's opacity as a fraction, in 65536ths or as a percentage, zero included", () => {
    const alpha = (fill: string) => {
      const body = `<w:p><w:r><w:pict><v:rect style="width:50pt;height:50pt" fillcolor="#FF0000">${fill}</v:rect></w:pict></w:r></w:p>`;
      const para = parseDocx(makeDocx({ body }), nodeCompression.inflateRaw).body[0] as Paragraph;
      return (para.content.find((i) => i.type === 'drawing') as DrawingInline).drawing.fill?.alpha;
    };
    expect(alpha('<v:fill opacity="0"/>')).toBe(0);
    expect(alpha('<v:fill opacity="50%"/>')).toBeCloseTo(0.5, 6);
    expect(alpha('<v:fill opacity="32768f"/>')).toBeCloseTo(0.5, 6);
    expect(alpha('<v:fill opacity=".25"/>')).toBeCloseTo(0.25, 6);
    expect(alpha('')).toBe(1);
    expect(alpha('<v:fill opacity="junk"/>')).toBe(1);
  });
});

describe('page borders', () => {
  it('draws double and dashed page borders in their style', async () => {
    const borders = '<w:pgBorders w:offsetFrom="page"><w:top w:val="double" w:sz="12" w:space="24"/><w:left w:val="dashed" w:sz="8" w:space="24"/>' +
      '<w:bottom w:val="double" w:sz="12" w:space="24"/><w:right w:val="single" w:sz="8" w:space="24"/></w:pgBorders>';
    const page = (await layout(makeDocx({ body: p('Text'), sectPr: A4_SECT.replace('</w:sectPr>', `${borders}</w:sectPr>`) }))).pages[0];
    const rules = [...page.behind, ...page.front].filter((o): o is Extract<DrawOpLike, { kind: 'line' }> => o.kind === 'line');
    const horizontal = rules.filter((l) => l.y1 === l.y2);
    // Top and bottom: two 1.5pt lines each, 1.5pt apart; the outer one runs on to the corner.
    expect(horizontal).toHaveLength(4);
    expect(horizontal.every((l) => l.width === 1.5)).toBe(true);
    const top = horizontal.filter((l) => l.y1 < 100).sort((a, b) => a.y1 - b.y1);
    expect(top[1].y1 - top[0].y1).toBeCloseTo(3, 6);
    expect(top[1].x1 - top[0].x1).toBeCloseTo(3, 6);
    const vertical = rules.filter((l) => l.x1 === l.x2).sort((a, b) => a.x1 - b.x1);
    expect(vertical).toHaveLength(2);
    expect(vertical[0].dash).toBeDefined();
    expect(vertical[1].dash).toBeUndefined();
  });
});

describe('equations', () => {
  const mr = (t: string) => `<m:r><m:t xml:space="preserve">${t}</m:t></m:r>`;
  const display = (math: string) => `<w:p><m:oMathPara><m:oMath>${math}</m:oMath></m:oMathPara></w:p>`;

  it('stacks a displayed fraction on the maths axis, centred on its own line', async () => {
    const frac = `<m:f><m:num>${mr('a+b')}</m:num><m:den>${mr('2')}</m:den></m:f>`;
    const page = (await layout(makeDocx({ styles: ARIAL_STYLES, body: display(mr('y=') + frac) }))).pages[0];
    const ops = glyphOps(page.body);
    const at = (t: string) => ops.find((o) => o.texts.includes(t))!;
    const bar = page.body.find((o) => o.kind === 'rect');
    if (bar?.kind !== 'rect') throw new Error('no fraction bar');
    // The bar centres on the axis, 0.29em above the baseline, between numerator and denominator.
    expect(at('y').y - (bar.y + bar.h / 2)).toBeCloseTo(0.29 * 11, 1);
    expect(at('a').y).toBeLessThan(bar.y);
    expect(at('2').y).toBeGreaterThan(bar.y + bar.h);
    // A displayed equation centres: as much room before "y" as after the bar.
    const left = at('y').x - 72;
    const right = 595.3 - 72 - (bar.x + bar.w);
    expect(left).toBeCloseTo(right, 0);
  });

  it('lines up an equation array at its alignment marks', async () => {
    const rows = `<m:eqArr><m:e>${mr('x&=1')}</m:e><m:e>${mr('y+z&=22')}</m:e></m:eqArr>`;
    const page = (await layout(makeDocx({ styles: ARIAL_STYLES, body: display(rows) }))).pages[0];
    const eqs = glyphOps(page.body).filter((o) => o.texts.join('') === '=');
    expect(eqs).toHaveLength(2);
    expect(eqs[0].x + eqs[0].xs[0]).toBeCloseTo(eqs[1].x + eqs[1].xs[0], 2);
    // Rows of plain text sit a line apart (1.18em).
    expect(eqs[1].y - eqs[0].y).toBeCloseTo(1.18 * 11, 1);
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
    // Webdings clipboard (drawn from the bundled Noto Color Emoji), Webdings box (U+2B1C, which Noto
    // Color Emoji has, drawn as the look-alike □ instead), Webdings "no piracy" (in no bundled font), Wingdings
    // ➢ (drawn as the older U+27A2 look-alike), Wingdings 3 rightwards arrow, and a Wingdings 2 code
    // with no Unicode equivalent.
    const body = `<w:p>${sym('Webdings', 'F0A4')}${sym('Webdings', 'F063')}${sym('Webdings', 'F023')}${sym('Wingdings', 'F0D8')}${sym('Wingdings 3', 'F067')}${sym('Wingdings 2', 'F0FF')}</w:p>`;
    const { result, warnings } = await layoutWithWarnings(makeDocx({ styles: ARIAL_STYLES, body }));
    const ops = glyphOps(result.pages[0].body);
    expect(ops.flatMap((o) => o.texts)).toEqual(['📋', '□', '🕲', '➢', '→', '']);
    expect(ops.find((o) => o.texts.includes('📋'))?.face.familyName).toBe('Noto Color Emoji');
    expect(ops.find((o) => o.texts.includes('□'))?.face.familyName).toBe('DejaVu Sans');
    expect(warnings).toEqual([
      'No font has a glyph for Webdings character U+F023 (U+1F572); it needs the Webdings font or one with U+1F572',
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
    expect(one.ops[0].face.familyName).toBe('TeX Gyre Adventor');
    expect(two.y - one.y).toBeCloseTo(((2060 + 451) / 2048) * 11, 1);
  });

  async function bundledRegistry(hints: Record<string, { altName?: string; family?: string }> = {}) {
    const fonts = new FontRegistry();
    for (const b of FULL_FONTS) fonts.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, covers: b.covers, load: async () => readFileSync(new URL(`../fonts/${b.file}`, import.meta.url)) }, 'bundled');
    for (const [name, h] of Object.entries(hints)) fonts.setHints(name, h);
    await fonts.prepare(Object.keys(hints).concat('Arial Black', 'Helvetica').map((family) => ({ family, bold: false, italic: false, codePoints: new Set([0x41, 0x414]) })));
    return fonts;
  }
  const lineEm = (f: { ascent: number; descent: number; externalLeading: number }) => f.ascent + f.descent + f.externalLeading;

  it('follows an alternate name to Segoe UI only when the document uses Segoe UI itself', async () => {
    // Word for Mac's PDFs draw such a Myriad Pro in Arial; Segoe UI comes from Office's cloud fonts.
    const without = await bundledRegistry({ 'Myriad Pro': { altName: 'Segoe UI', family: 'swiss' } });
    expect(without.resolve('Myriad Pro', false, false).face.familyName).toBe('Liberation Sans');
    const withIt = await bundledRegistry({ 'Myriad Pro': { altName: 'Segoe UI', family: 'swiss' }, 'Segoe UI': { family: 'swiss' } });
    expect(withIt.resolve('Myriad Pro', false, false).face.familyName).toBe('Selawik');
  });

  it("gives Helvetica Word for Mac's line, without its gap when reached by an alternate name", async () => {
    const fonts = await bundledRegistry({ 'Helvetica, Arial, sans-serif': { altName: 'Helvetica', family: 'roman' } });
    expect(lineEm(fonts.resolve('Helvetica', false, false))).toBeCloseTo(1.2, 3);
    expect(lineEm(fonts.resolve('Helvetica, Arial, sans-serif', false, false))).toBeCloseTo(1.0, 3);
  });

  it("sets a Cyrillic letter Arial Black's stand-in lacks on Arial Black's line", async () => {
    const fonts = await bundledRegistry();
    const de = fonts.faceForChar('Arial Black', false, false, 0x414);
    expect(fonts.resolve('Arial Black', false, false).face.font.hasGlyph(0x414)).toBe(false);
    expect(de.face.familyName).not.toBe('Archivo Black');
    expect(lineEm(de)).toBeCloseTo(1.101 + 0.31, 3);
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

  it('aligns text on a right, centre or decimal tab without its trailing spaces', async () => {
    const end = async (type: string, text: string) => {
      const body = `<w:p><w:pPr><w:tabs><w:tab w:val="${type}" w:pos="4000"/></w:tabs></w:pPr><w:r><w:t>Left</w:t><w:tab/><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
      const ops = glyphOps((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
      const nines = ops.flatMap((o) => o.texts.map((t, i) => ({ t, x: o.x + o.xs[i] }))).filter((c) => c.t === '9');
      return nines.at(-1)!.x + (1139 / 2048) * 11;
    };
    for (const type of ['right', 'center', 'decimal']) expect(await end(type, '99   ')).toBeCloseTo(await end(type, '99'), 2);
    expect(await end('right', '99')).toBeCloseTo(72 + 200, 1);
  });

  it('keeps a line of nothing but tabs whole past the margin before Word 2013, but wraps tabs that follow text', async () => {
    const tabs = (text: string) => `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="9300"/></w:tabs></w:pPr><w:r>${text ? `<w:t>${text}</w:t>` : ''}<w:tab/><w:tab/></w:r></w:p>`;
    const body = p('Before') + tabs('') + p('Middle') + tabs('Text') + p('After');
    const gaps = async (mode: number) => {
      const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, settings: compat(mode), body }))).pages[0].body);
      const y = (t: string) => ls.find((l) => l.text === t)!.y;
      return [(y('Middle') - y('Before')) / ARIAL_11, (y('After') - y('Middle')) / ARIAL_11];
    };
    // After "Text" the first tab, past the margin, starts a line that keeps the second.
    const [blank14, text14] = await gaps(14);
    expect(blank14).toBeCloseTo(2, 2);
    expect(text14).toBeCloseTo(3, 2);
    // From Word 2013 the tabs-only line wraps at its overflowing tab as well.
    const [blank15] = await gaps(15);
    expect(blank15).toBeCloseTo(3, 2);
  });

  it('shows a soft hyphen only where a line breaks at it', async () => {
    // 35pt holds "aaaa-" but not "aaaabbbb".
    const narrow = `<w:pPr><w:ind w:right="${Math.round((451.3 - 35) * 20)}"/></w:pPr>`;
    const body = `<w:p>${narrow}<w:r><w:t>aaaa</w:t><w:softHyphen/><w:t>bbbb</w:t></w:r></w:p><w:p><w:r><w:t>End</w:t><w:softHyphen/></w:r></w:p>`;
    const ls = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body);
    expect(ls.map((l) => l.text)).toEqual(['aaaa-', 'bbbb', 'End']);
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
      `<w:p><w:sdt><w:sdtPr>${pr}<w:dataBinding w:prefixMappings="${storeId === core ? "xmlns:ns1='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:ns0='http://purl.org/dc/elements/1.1/'" : "xmlns:ns0='urn:test'"}" w:xpath="${xpath}" w:storeItemID="${storeId}"/></w:sdtPr><w:sdtContent><w:r><w:t>${cached}</w:t></w:r></w:sdtContent></w:sdt></w:p>`;
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

  it("strokes synthetic bold text solid after a dashed line", async () => {
    // Noto Sans CJK JP has no bold face, so bold kanji are stroked.
    const body = p('Dashed', '', '<w:u w:val="dash"/>') + p('文字', '', '<w:rFonts w:eastAsia="MS Mincho"/><w:b/>');
    const content = pdfStreams((await convertDocx(makeDocx({ styles: ARIAL_STYLES, body }))).pdf).join('\n');
    const dashed = content.search(/\[[\d.]+ [\d.]+\] 0 d/);
    const bold = content.indexOf(' 2 Tr');
    expect(dashed).toBeGreaterThanOrEqual(0);
    expect(bold).toBeGreaterThan(dashed);
    expect(content.slice(content.lastIndexOf('BT', bold), bold)).toContain('[] 0 d');
  });
});

describe('charts', () => {
  const inlineChart = (rid: string, wPt: number, hPt: number) =>
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${wPt * 12700}" cy="${hPt * 12700}"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const pts = (cats: string[], fmt?: string) =>
    `${fmt ? `<c:formatCode>${fmt}</c:formatCode>` : ''}<c:ptCount val="${cats.length}"/>${cats.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`).join('')}`;
  const axes =
    '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/></c:catAx>' +
    '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="1"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
  const values = [4.3, 2.5, 3.5, 4.5];
  const barChart =
    '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr><c:cat><c:strLit>${pts(['North', 'South', 'East', 'West'])}</c:strLit></c:cat><c:val><c:numLit>${pts(values.map(String), 'General')}</c:numLit></c:val></c:ser>` +
    `<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart>`;

  it('falls back to a bounded set of log-axis labels and rejects an unrepresentable linear span', async () => {
    const axis = (scaling: string) => barChart.replace('<c:valAx><c:axId val="2"/><c:scaling>', '<c:valAx><c:axId val="2"/><c:scaling>' + scaling);
    const log = await chartPage(axis('<c:logBase val="1.0001"/><c:min val="0"/><c:max val="8"/>'));
    const labels = glyphOps(log).map((g) => g.texts.join(''));
    expect(labels).toContain('8');
    expect(labels.length).toBeLessThan(20);
    await expect(chartPage(axis('<c:min val="-1e308"/><c:max val="1e308"/>'))).rejects.toThrow(/axis range/);
  });

  it("draws a chart from its cached data, on Excel's automatic value axis", async () => {
    // Word stores no picture of a chart, so without drawing the data the chart would be blank.
    const r = await layout(makeDocx({ body: `<w:p>${inlineChart('rIdChart1', 360, 216)}</w:p>`, charts: { rIdChart1: barChart } }));
    const ops = r.pages[0].body;
    const texts = glyphOps(ops).map((g) => ({ text: g.texts.join(''), x: g.x, y: g.y }));
    for (const t of ['North', 'South', 'East', 'West', 'Sales']) expect(texts.map((g) => g.text)).toContain(t);
    // 4.5 at most rounds out to an axis that ends on a whole step: 0 to 5.
    const numbers = texts.filter((g) => /^\d+(\.\d+)?$/.test(g.text));
    const zero = numbers.find((g) => g.text === '0')!;
    const five = numbers.find((g) => g.text === '5')!;
    expect(zero).toBeDefined();
    expect(five).toBeDefined();
    expect(Math.max(...numbers.map((g) => Number(g.text)))).toBe(5);
    const bars = ops
      .filter((o): o is PathOp => o.kind === 'path' && (o.fill ?? '').replace('#', '').toUpperCase() === 'FF0000')
      .map((o) => {
        const ys = o.segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 1));
        return { top: Math.min(...ys), bottom: Math.max(...ys) };
      })
      // The fifth red shape is the legend key, a small square under the plot.
      .filter((b) => b.bottom - b.top > 10);
    expect(bars).toHaveLength(4);
    // Bars stand on the zero line, as tall as their values, the tallest 9/10 of the way to the 5 label.
    const base = bars[0].bottom;
    bars.forEach((b, i) => {
      expect(b.bottom).toBeCloseTo(base, 3);
      expect((b.bottom - b.top) / (bars[3].bottom - bars[3].top)).toBeCloseTo(values[i] / 4.5, 3);
    });
    expect((zero.y - (bars[3].top + (zero.y - base))) / (zero.y - five.y)).toBeCloseTo(0.9, 3);
  });

  it('formats axis and data labels with Excel number formats', () => {
    expect(formatNumber(0.15, '0.00%')).toBe('15.00%');
    expect(formatNumber(12345.6, '#,##0')).toBe('12,346');
    expect(formatNumber(-1234.5, '#,##0.0;(#,##0.0)')).toBe('(1,234.5)');
    expect(formatNumber(3.5, '"$"0.00')).toBe('$3.50');
    expect(formatNumber(0.1 + 0.2, 'General')).toBe('0.3');
    // Serial dates and times: 44409 is 1 August 2021, and 51 seconds after midnight.
    expect(formatNumber(44409, 'mmm\\-yy')).toBe('Aug-21');
    expect(formatNumber(51 / 86400, 'h:mm:ss AM/PM')).toBe('12:00:51 AM');
  });

  it('tints and shades DrawingML colours in linear light, as Word does', () => {
    // Measured from Word's PDFs: accent2 C0504D tinted 77% and shaded 76%, and black tinted 75% (chart gridlines).
    expect(applyColorMods('C0504D', [{ name: 'tint', val: 77000 }]).color).toBe('D19392');
    expect(applyColorMods('C0504D', [{ name: 'shade', val: 76000 }]).color).toBe('AA4643');
    expect(applyColorMods('000000', [{ name: 'tint', val: 75000 }]).color).toBe('898989');
  });

  it('lets satMod push a colour past full saturation, clipping at the RGB cube as Word does', () => {
    // From a Word for Mac gradient's samples: 00B050 is already fully saturated, yet satMod 115% brightens it.
    expect(applyColorMods('00B050', [{ name: 'satMod', val: 115000 }]).color).toBe('00BD4F');
    expect(applyColorMods('00B050', [{ name: 'shade', val: 30000 }, { name: 'satMod', val: 115000 }]).color).toBe('006D2A');
  });

  const strokes = (ops: DrawOpLike[], color: string) => ops.filter((o): o is PathOp => o.kind === 'path' && !o.fill && o.stroke === color);
  const ys = (o: PathOp) => o.segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 1));
  const chartWith = (plot: string) =>
    `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}${axes}</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart>`;
  const seriesWith = (extra: string, vals = values) =>
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Sales</c:v></c:tx><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr>${extra}<c:cat><c:strLit>${pts(['North', 'South', 'East', 'West'])}</c:strLit></c:cat><c:val><c:numLit>${pts(vals.map(String), 'General')}</c:numLit></c:val></c:ser>`;
  const chartPage = async (xml: string) => (await layout(makeDocx({ body: `<w:p>${inlineChart('rIdChart1', 360, 216)}</w:p>`, charts: { rIdChart1: xml } }))).pages[0].body;
  const axisTop = (ops: DrawOpLike[]) => Math.max(...glyphOps(ops).map((g) => g.texts.join('')).filter((t) => /^\d+$/.test(t)).map(Number));

  it('draws error bars and scales the value axis to take them in', async () => {
    // Without the bars 4.5 rounds out to 5; one either side of each point reaches 5.5, so the axis runs to 6.
    const bars = '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="0"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></c:spPr><c:val val="1"/></c:errBars>';
    const ops = await chartPage(chartWith(`<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesWith(bars)}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`));
    expect(axisTop(ops)).toBe(6);
    const drawn = strokes(ops, '00AA00');
    expect(drawn).toHaveLength(4);
    // Each bar spans two units: a third of the axis, with a cap across each end.
    const unit = (Math.max(...ys(drawn[0])) - Math.min(...ys(drawn[0]))) / 2;
    for (const b of drawn) {
      expect(b.segs.filter((s) => s.op === 'M')).toHaveLength(3);
      expect((Math.max(...ys(b)) - Math.min(...ys(b))) / unit).toBeCloseTo(2, 3);
    }
  });

  it('draws trendlines into the axis range and names them in the legend', async () => {
    // A linear fit of 4.3, 2.5, 3.5, 4.5 carried two periods back starts near 3.1 and ends near 4.1: inside 0..5.
    // Carried six periods forward it climbs past 5, so the axis grows to take the line in.
    const trend = (forward: number) => `<c:trendline><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln></c:spPr><c:trendlineType val="linear"/><c:forward val="${forward}"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>`;
    const plot = (forward: number) => chartWith(`<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesWith(trend(forward))}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`);
    const near = await chartPage(plot(0));
    expect(axisTop(near)).toBe(5);
    expect(strokes(near, 'AA00AA').length).toBeGreaterThanOrEqual(1);
    expect(glyphOps(near).map((g) => g.texts.join(''))).toContain('Linear (Sales)');
    expect(axisTop(await chartPage(plot(6)))).toBeGreaterThan(5);
  });

  it('draws radar charts: closed series outlines, value labels and category names', async () => {
    const radar = `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${seriesWith('<c:marker><c:symbol val="none"/></c:marker>')}<c:axId val="1"/><c:axId val="2"/></c:radarChart>`;
    const ops = await chartPage(chartWith(radar));
    const texts = glyphOps(ops).map((g) => g.texts.join(''));
    for (const t of ['North', 'South', 'East', 'West', '0']) expect(texts).toContain(t);
    const outline = strokes(ops, '0000FF').filter((o) => o.segs.length > 2);
    expect(outline).toHaveLength(1);
    expect(outline[0].segs.filter((s) => s.op === 'L')).toHaveLength(3);
    expect(outline[0].segs.at(-1)!.op).toBe('Z');
    // The first category stands straight up from the centre, the third of four (East) straight down.
    const north = outline[0].segs[0].pts;
    const east = outline[0].segs[2].pts;
    expect(north[0]).toBeCloseTo(east[0], 3);
    expect(north[1]).toBeLessThan(east[1]);
  });

  const radarWith = (vary: boolean, ser: string) => `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="${vary ? 1 : 0}"/>${ser}<c:axId val="1"/><c:axId val="2"/></c:radarChart>`;
  const lineOf = (rgb: string) => `<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill></a:ln></c:spPr>`;

  it("draws each stretch of a radar line in the colour of the point it runs into", async () => {
    const dPt = `<c:marker><c:symbol val="none"/></c:marker><c:dPt><c:idx val="2"/>${lineOf('00AA00')}</c:dPt>`;
    const ops = await chartPage(chartWith(radarWith(false, seriesWith(dPt))));
    // South to East is the point's colour; the other three stretches, closing the ring, the series'.
    // (Four categories put every stretch on a diagonal; the legend's line key is level.)
    const stretches = (color: string) => strokes(ops, color).filter((o) => o.segs.length === 2 && ys(o)[0] !== ys(o)[1]);
    const green = stretches('00AA00');
    expect(green).toHaveLength(1);
    expect(green[0].segs.map((s) => s.op)).toEqual(['M', 'L']);
    expect(ys(green[0])[1]).toBeGreaterThan(ys(green[0])[0]);
    expect(stretches('0000FF')).toHaveLength(3);
    // Varied colours give each point of the chart's lone series its own automatic colour, line and
    // all (as Excel prints them), so each stretch takes its point's colour and no ring is left.
    const varied = await chartPage(chartWith(radarWith(true, seriesWith('<c:marker><c:symbol val="none"/></c:marker>'))));
    const paths = varied.filter((o): o is PathOp => o.kind === 'path' && !o.fill && !!o.stroke);
    expect(paths.filter((o) => o.stroke === '0000FF' && o.segs.at(-1)!.op === 'Z')).toHaveLength(0);
    // Four stretches in four colours (the spokes, also diagonal, share the axis colour).
    const legs = paths.filter((o) => o.segs.length === 2 && ys(o)[0] !== ys(o)[1]);
    const count = (color: string | null | undefined) => legs.filter((o) => o.stroke === color).length;
    expect(new Set(legs.map((o) => o.stroke).filter((c) => count(c) === 1)).size).toBe(4);
  });

  it("draws a radar's spokes from its category axis line unless the axis is deleted, then from its gridlines", async () => {
    const catAx = (deleted: boolean, extra: string) => `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${deleted ? 1 : 0}"/><c:axPos val="b"/>${extra}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="CC0000"/></a:solidFill></a:ln></c:spPr><c:crossAx val="2"/><c:crosses val="autoZero"/><c:auto val="1"/></c:catAx>`;
    const grid = `<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="00CC00"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>`;
    const chart = (deleted: boolean) => chartWith(radarWith(false, seriesWith('<c:marker><c:symbol val="none"/></c:marker>'))).replace(/<c:catAx>.*?<\/c:catAx>/, catAx(deleted, grid));
    const shownAxis = await chartPage(chart(false));
    expect(strokes(shownAxis, 'CC0000')).toHaveLength(4);
    expect(strokes(shownAxis, '00CC00')).toHaveLength(0);
    const deletedAxis = await chartPage(chart(true));
    expect(strokes(deletedAxis, 'CC0000')).toHaveLength(0);
    expect(strokes(deletedAxis, '00CC00')).toHaveLength(4);
  });

  it('lists the categories of a lone radar series in varied colours, but not beside a trendline', async () => {
    const texts = (ops: DrawOpLike[]) => glyphOps(ops).map((g) => g.texts.join(''));
    const radar = texts(await chartPage(chartWith(radarWith(true, seriesWith('')))));
    expect(radar).not.toContain('Sales');
    for (const c of ['North', 'South', 'East', 'West']) expect(radar.filter((t) => t === c)).toHaveLength(2);
    const trend = '<c:trendline><c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>';
    const varied = (extra: string) => chartWith(`<c:lineChart><c:grouping val="standard"/><c:varyColors val="1"/>${seriesWith(extra)}<c:marker val="1"/><c:axId val="1"/><c:axId val="2"/></c:lineChart>`);
    expect(texts(await chartPage(varied('')))).not.toContain('Sales');
    expect(texts(await chartPage(varied(trend)))).toEqual(expect.arrayContaining(['Sales', 'Linear (Sales)']));
  });

  it('centres a side legend between 11pt below the top and 11pt above the bottom', async () => {
    const ops = await chartPage(barChart.replace('<c:legendPos val="b"/>', '<c:legendPos val="r"/>'));
    const clip = ops.find((o): o is Extract<DrawOpLike, { kind: 'clip' }> => o.kind === 'clip')!;
    const keys = ops.filter((o): o is PathOp => o.kind === 'path' && o.fill === 'FF0000');
    const key = keys.reduce((a, b) => (Math.min(...b.segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 0))) > Math.min(...a.segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 0))) ? b : a));
    // Without a title the band runs from 11pt to 11pt, so a lone entry sits at the chart's middle.
    const mid = (Math.max(...ys(key)) + Math.min(...ys(key))) / 2;
    expect(mid).toBeCloseTo(clip.y + 1 + (clip.h - 2) / 2, 3);
  });
});

describe('shape gradients', () => {
  const shape = (spPr: string) =>
    '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="1270000"/><a:graphic><a:graphicData>' +
    `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:prstGeom prst="rect"/>${spPr}</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  const stop = (pos: number, rgb: string, alpha?: number) => `<a:gs pos="${pos}"><a:srgbClr val="${rgb}">${alpha === undefined ? '' : `<a:alpha val="${alpha}"/>`}</a:srgbClr></a:gs>`;
  const grad = (stops: string, kind = '<a:lin ang="5400000" scaled="0"/>') => `<a:gradFill><a:gsLst>${stops}</a:gsLst>${kind}</a:gradFill>`;
  const blackToWhite = stop(0, '000000') + stop(100000, 'FFFFFF');
  const focused = (path: string) => `<a:path path="${path}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`;
  const drawn = async (spPr: string) => {
    const ops = (await layout(makeDocx({ body: shape(spPr) }))).pages[0].body;
    const clip = ops.find((o): o is Extract<DrawOpLike, { kind: 'clip' }> => o.kind === 'clip')!;
    return { clip, shade: ops.find((o): o is ShadeOp => o.kind === 'shade')!, bands: ops.filter((o) => o.kind === 'path').length };
  };

  it('draws a linear gradient as one shading down the shape, blended straight through sRGB', async () => {
    // A shape's gradient is a straight sRGB blend in Word for Mac's PDFs (charts and SmartArt blend in linear light).
    const { clip, shade, bands } = await drawn(grad(blackToWhite));
    expect(bands).toBe(0);
    expect(shade.colors).toEqual(['000000', 'FFFFFF']);
    const mid = clip.x + clip.w / 2;
    expect(shade.geometry).toEqual({ type: 'axial', x0: expect.closeTo(mid, 2), y0: expect.closeTo(clip.y, 2), x1: expect.closeTo(mid, 2), y1: expect.closeTo(clip.y + clip.h, 2) });
  });

  it('shades a path gradient out from its focus, in ellipses or rectangles', async () => {
    const { clip, shade } = await drawn(grad(blackToWhite, focused('circle')));
    // The last ring reaches the corners.
    expect(shade.geometry).toEqual({ type: 'radial', cx: expect.closeTo(clip.x + 50, 2), cy: expect.closeTo(clip.y + 50, 2), rx: expect.closeTo(100 * Math.SQRT1_2, 2), ry: expect.closeTo(100 * Math.SQRT1_2, 2) });
    const rect = (await drawn(grad(blackToWhite, focused('rect')))).shade.geometry;
    if (rect.type !== 'rect') throw new Error('not rectangular rings');
    rect.inner.forEach((v, i) => expect(v).toBeCloseTo([clip.x + 50, clip.y + 50][i % 2], 2));
    rect.outer.forEach((v, i) => expect(v).toBeCloseTo([clip.x, clip.y, clip.x + 100, clip.y + 100][i], 2));
  });

  it("takes a gradient's opacity once, or each stop's own where they differ", async () => {
    const even = (await drawn(grad(stop(0, '000000', 50000) + stop(100000, '000000', 50000)))).shade;
    expect(even.alpha).toBeCloseTo(0.5, 3);
    expect(even.alphas).toBeUndefined();
    const fade = (await drawn(grad(stop(0, '000000', 0) + stop(100000, '000000')))).shade;
    expect(fade.alpha).toBeUndefined();
    expect(fade.alphas![0]).toBeCloseTo(0, 3);
    expect(fade.alphas!.at(-1)).toBeCloseTo(1, 3);
  });

  it.skipIf(!ghostscript)('paints gradients as a PDF reader shows them, translucent ones evenly', async () => {
    const levels = async (fill: string, points: Array<[number, number]>) => {
      const docx = makeDocx({ body: shape(fill) });
      const { clip } = await drawn(fill);
      const { at } = ghostscriptPixels((await convertDocx(docx)).pdf, 72);
      return points.map(([fx, fy]) => at(clip.x + 100 * fx, clip.y + 100 * fy)[0]);
    };
    // Black at half opacity over the white page is mid grey from top to foot.
    for (const v of await levels(grad(stop(0, '000000', 50000) + stop(100000, '000000', 50000)), [[0.5, 0.1], [0.5, 0.5], [0.5, 0.9]])) expect(Math.abs(v - 128)).toBeLessThan(10);
    // Fading in from clear to solid black.
    const [top, middle, foot] = await levels(grad(stop(0, '000000', 0) + stop(100000, '000000')), [[0.5, 0.1], [0.5, 0.5], [0.5, 0.9]]);
    expect(top).toBeGreaterThan(215);
    expect(Math.abs(middle - 128)).toBeLessThan(15);
    expect(foot).toBeLessThan(40);
    // Black at the focus, white at the edge, out in circles and in rectangles.
    for (const path of ['circle', 'rect']) {
      const [centre, nearEdge, corner] = await levels(grad(blackToWhite, focused(path)), [[0.5, 0.5], [0.5, 0.05], [0.03, 0.03]]);
      expect(centre).toBeLessThan(20);
      expect(nearEdge).toBeGreaterThan(path === 'rect' ? 215 : 160);
      expect(corner).toBeGreaterThan(215);
    }
  });
});
