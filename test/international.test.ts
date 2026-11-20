/**
 * East Asian and right-to-left text: document grids, missing CJK fonts, punctuation compression,
 * autospacing, East Asian list numbers and bidirectional text. Expected numbers come from Word's
 * PDFs; CJK text is drawn with the bundled Noto Sans CJK.
 */
import { describe, expect, it } from 'vitest';
import type { Paragraph, TextInline } from '../src/docx/model.js';
import { formatNumber } from '../src/docx/numbering.js';
import { parseDocx } from '../src/docx/parser.js';
import { arabicForms, bidiClass, bidiLevels, visualOrder } from '../src/layout/bidi.js';
import { nodeCompression } from '../src/util/node-compression.js';
import { makeDocx, p } from './helpers/docx.js';
import { glyphOps, layout, lines } from './helpers/layout.js';
import type { DrawOp } from '../src/layout/types.js';

/** A4 with 1" margins (451.3pt of text width), optionally with a document grid. */
const sect = (grid = '', refs = '') =>
  `<w:sectPr>${refs}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>${grid}</w:sectPr>`;

/** SimSun, which is not bundled, at `sz` half-points, single-spaced with no space after. */
const simsun = (sz: number) =>
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="SimSun" w:hAnsi="SimSun" w:eastAsia="SimSun"/><w:sz w:val="${sz}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>`;

/** Each drawn character with its x, in reading order. */
function chars(ops: DrawOp[]): Array<{ t: string; x: number }> {
  const out: Array<{ t: string; x: number }> = [];
  for (const op of glyphOps(ops)) op.texts.forEach((t, i) => out.push({ t, x: op.x + op.xs[i] }));
  return out;
}

const xOf = (cs: Array<{ t: string; x: number }>, t: string) => cs.find((c) => c.t === t)!.x;

describe('East Asian lines', () => {
  it("keeps a missing Chinese font's line height on its fallback", async () => {
    // Word gives SimSun a line of 258 + 74 of 256 units, about 1.3em.
    const r = await layout(makeDocx({ styles: simsun(24), body: p('中文') + p('中文') }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo((12 * 332) / 256, 1);
  });

  it("keeps DengXian's line height on its fallback", async () => {
    // Word makes DengXian's lines 1.3 times its own 1659 + 475 of 2048 units, like SimSun's.
    const r = await layout(makeDocx({ styles: simsun(24).replace(/SimSun/g, 'DengXian'), body: p('中文') + p('中文') }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo((12 * 1.3 * (1659 + 475)) / 2048, 1);
  });

  it("lets SimSun's taller ascent absorb a Latin font's line gap", async () => {
    // Times New Roman's 0.51pt line gap sits above its own ascent, below SimSun's: the line stays
    // SimSun's 1.3em, as in Word's tables of contents.
    const mixed = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t>Abstract </w:t></w:r><w:r><w:t>中文</w:t></w:r></w:p>';
    const ys = lines((await layout(makeDocx({ styles: simsun(24), body: mixed + mixed }))).pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo((12 * 332) / 256, 1);
  });

  it("grows a line to Word's fallback height for characters the East Asian font lacks", async () => {
    // MS Mincho has no Hangul and no simplified forms like 轻, which Word draws from fallback fonts
    // in a line about 1.71em tall; its own kanji keep its 1.3em line.
    const styles = simsun(22).replace(/SimSun/g, 'MS Mincho');
    const ys = lines((await layout(makeDocx({ styles, body: p('日本') + p('한국') + p('轻量') + p('日本') }))).pages[0].body).map((l) => l.y);
    expect(ys[2] - ys[1]).toBeCloseTo(11 * (1.28 + 0.43), 1);
    expect(ys[3] - ys[2]).toBeCloseTo(11 * (0.43 + 258 / 256), 1);
    expect(ys[1] - ys[0]).toBeCloseTo(11 * (74 / 256 + 1.28), 1);
  });

  it('draws ASCII in a fixed-pitch East Asian font half an em wide', async () => {
    // SimSun's ASCII is monospaced at half an em; PMingLiU's is proportional.
    const at = async (font: string) => chars((await layout(makeDocx({ styles: simsun(24).replace(/SimSun/g, font), body: p('mil') }))).pages[0].body);
    const fixed = await at('SimSun');
    expect(xOf(fixed, 'i') - xOf(fixed, 'm')).toBeCloseTo(6, 2);
    expect(xOf(fixed, 'l') - xOf(fixed, 'i')).toBeCloseTo(6, 2);
    const proportional = await at('PMingLiU');
    expect(xOf(proportional, 'i') - xOf(proportional, 'm')).toBeGreaterThan(7);
  });

  it('draws curly quotes and ellipses in a fixed-pitch East Asian font an em wide', async () => {
    const cs = chars((await layout(makeDocx({ styles: simsun(24), body: p('“中”…') }))).pages[0].body);
    expect(xOf(cs, '中') - xOf(cs, '“')).toBeCloseTo(12, 2);
    expect(xOf(cs, '”') - xOf(cs, '中')).toBeCloseTo(12, 2);
    expect(xOf(cs, '…') - xOf(cs, '”')).toBeCloseTo(12, 2);
  });

  it('centres a line in the document grid lines it takes', async () => {
    const r = await layout(makeDocx({ styles: simsun(21), sectPr: sect('<w:docGrid w:type="lines" w:linePitch="360"/>'), body: p('中文') + p('中文') }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    // A 10.5pt SimSun line (332/256 em tall) in an 18pt grid line: the rest split above and below.
    const h = (10.5 * 332) / 256;
    expect(ys[0]).toBeCloseTo(72 + (18 - h) / 2 + h - (10.5 * 74) / 256, 2);
    expect(ys[1] - ys[0]).toBeCloseTo(18, 2);
  });

  it("takes one grid line for a line that fits it but for a Latin font's external leading", async () => {
    // Arial 12 beside SimSun 12 is 15.56pt, Arial's 0.39pt external leading falling under SimSun's
    // taller ascent: the 15.6pt grid line holds it.
    const body = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Title </w:t></w:r><w:r><w:t>中文</w:t></w:r></w:p>';
    const r = await layout(makeDocx({ styles: simsun(24), sectPr: sect('<w:docGrid w:type="lines" w:linePitch="312"/>'), body: body + body }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    expect(ys[1] - ys[0]).toBeCloseTo(15.6, 2);
  });

  it("sizes spacing given in lines by the section's line grid, or as 12pt without one", async () => {
    // w:beforeLines is in hundredths of a line; a later w:before in points replaces it.
    const styles = simsun(21) + '<w:style w:type="paragraph" w:styleId="Gap"><w:name w:val="Gap"/><w:pPr><w:spacing w:beforeLines="100"/></w:pPr></w:style>';
    const body = p('中文') + p('中文', '<w:pStyle w:val="Gap"/>') + p('中文', '<w:pStyle w:val="Gap"/><w:spacing w:before="0"/>');
    const gaps = async (grid: string) => {
      const ys = lines((await layout(makeDocx({ styles, sectPr: sect(grid), body }))).pages[0].body).map((l) => l.y);
      return (ys[1] - ys[0]) - (ys[2] - ys[1]);
    };
    expect(await gaps('<w:docGrid w:type="lines" w:linePitch="360"/>')).toBeCloseTo(18, 1);
    expect(await gaps('')).toBeCloseTo(12, 1);
  });

  it('snaps lines in table cells to the document grid only when the document asks it to', async () => {
    const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p('中文')}${p('中文')}</w:tc></w:tr></w:tbl>`;
    const pitch = async (grid: string, settings = '') => {
      const ys = lines((await layout(makeDocx({ styles: simsun(21), settings, sectPr: sect(grid), body: table }))).pages[0].body).map((l) => l.y);
      return ys[1] - ys[0];
    };
    const grid = '<w:docGrid w:type="lines" w:linePitch="360"/>';
    expect(await pitch(grid)).toBeCloseTo(await pitch(''), 2);
    expect(await pitch(grid, '<w:compat><w:adjustLineHeightInTable/></w:compat>')).toBeCloseTo(18, 2);
  });

  it('widens every character by a character grid, except in runs that leave the grid', async () => {
    const grid = '<w:docGrid w:type="linesAndChars" w:linePitch="240" w:charSpace="4096"/>';
    const body = '<w:p><w:r><w:t>一二三</w:t></w:r><w:r><w:rPr><w:snapToGrid w:val="0"/></w:rPr><w:t>四五六</w:t></w:r></w:p>';
    const cs = chars((await layout(makeDocx({ styles: simsun(20), sectPr: sect(grid), body }))).pages[0].body);
    // charSpace is in 4096ths of a point.
    expect(xOf(cs, '二') - xOf(cs, '一')).toBeCloseTo(11, 2);
    expect(xOf(cs, '四') - xOf(cs, '三')).toBeCloseTo(11, 2);
    expect(xOf(cs, '五') - xOf(cs, '四')).toBeCloseTo(10, 2);
  });

  it('squeezes full-width punctuation to keep a character on an overfull line', async () => {
    // 46 ten-point characters need 460pt; the four commas can give up 3pt each.
    const text = '一二三四五，'.repeat(4) + '六七八九十'.repeat(4) + '一二';
    const doc = (settings: string) => makeDocx({ styles: simsun(20), settings, body: p(text) });
    const plain = lines((await layout(doc(''))).pages[0].body);
    const squeezed = lines((await layout(doc('<w:characterSpacingControl w:val="compressPunctuation"/>'))).pages[0].body);
    expect(plain).toHaveLength(2);
    expect(squeezed).toHaveLength(1);
    const cs = chars(squeezed[0].ops);
    const last = cs[cs.length - 1];
    expect(last.x + 10).toBeLessThanOrEqual(72 + 451.3 + 0.01);
    // A comma gives up as much before its glyph as after it.
    const comma = cs.findIndex((c) => c.t === '，');
    const before = cs[comma].x - cs[comma - 1].x - 10;
    const after = cs[comma + 1].x - cs[comma].x - 10;
    expect(before).toBeLessThan(0);
    expect(after).toBeCloseTo(before, 2);
  });

  it('lets a closing punctuation mark hang past the right margin at its natural spacing', async () => {
    // 45 ten-point characters fill the 451.3pt line; the full stop starts inside it.
    const text = '一二三四五六七八九十'.repeat(4) + '一二三四五。六七';
    const doc = (pPr: string) => makeDocx({ styles: simsun(20), body: p(text, `<w:jc w:val="both"/>${pPr}`) });
    const hung = lines((await layout(doc(''))).pages[0].body);
    expect(hung).toHaveLength(2);
    const cs = chars(hung[0].ops);
    expect(cs[cs.length - 1]).toEqual({ t: '。', x: expect.closeTo(72 + 450, 2) });
    expect(xOf(cs, '二') - xOf(cs, '一')).toBeCloseTo(10, 2);
    // Without it the stop wraps, and takes the character before it along.
    const wrapped = lines((await layout(doc('<w:overflowPunct w:val="0"/>'))).pages[0].body);
    expect(chars(wrapped[1].ops).map((c) => c.t).join('')).toBe('五。六七');
  });

  it("measures character-unit indents in the paragraph mark's font size", async () => {
    const style = '<w:style w:type="paragraph" w:styleId="Ind"><w:name w:val="Ind"/><w:pPr><w:ind w:leftChars="200" w:left="420"/></w:pPr></w:style>';
    const x = async (sz: number, pPr: string) => {
      const body = p('Text', `<w:pStyle w:val="Ind"/>${pPr}`);
      return glyphOps((await layout(makeDocx({ styles: simsun(sz) + style, body }))).pages[0].body)[0].x;
    };
    expect(await x(24, '')).toBeCloseTo(72 + 24, 2);
    expect(await x(20, '')).toBeCloseTo(72 + 20, 2);
    expect(await x(20, '<w:rPr><w:sz w:val="28"/></w:rPr>')).toBeCloseTo(72 + 28, 2);
    // A twip indent set later replaces the inherited character one.
    expect(await x(24, '<w:ind w:left="720"/>')).toBeCloseTo(72 + 36, 2);
  });

  it('spaces East Asian text a quarter em from digits and Latin letters, but not from full-width brackets', async () => {
    const body = (pPr: string) => p('第5屆A（5）', pPr);
    const on = chars((await layout(makeDocx({ styles: simsun(20), body: body('') }))).pages[0].body);
    const off = chars((await layout(makeDocx({ styles: simsun(20), body: body('<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>') }))).pages[0].body);
    const gap = (cs: typeof on, a: number) => cs[a + 1].x - cs[a].x;
    expect(gap(on, 0) - gap(off, 0)).toBeCloseTo(2.5, 2); // 第5
    expect(gap(on, 1) - gap(off, 1)).toBeCloseTo(2.5, 2); // 5屆
    expect(gap(on, 2) - gap(off, 2)).toBeCloseTo(2.5, 2); // 屆A
    expect(gap(on, 4) - gap(off, 4)).toBeCloseTo(0, 2); // （5
    expect(gap(on, 5) - gap(off, 5)).toBeCloseTo(0, 2); // 5）
  });

  it("takes an empty theme East Asian font from the theme's font for the document's language", () => {
    const theme =
      '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hans" typeface="等线 Light"/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/></a:minorFont></a:fontScheme>';
    const styles = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault></w:docDefaults>';
    const font = (lang: string) => {
      const doc = parseDocx(makeDocx({ styles, theme, settings: `<w:themeFontLang w:val="en-US" w:eastAsia="${lang}"/>`, body: p('文字') }), nodeCompression.inflateRaw);
      return ((doc.body[0] as Paragraph).content[0] as TextInline).props.fontEastAsia;
    };
    expect(font('ja-JP')).toBe('游明朝');
    expect(font('zh-CN')).toBe('等线');
    expect(font('zh-TW')).toBe('新細明體');
  });
});

describe('East Asian list numbers', () => {
  it('counts in Chinese and Japanese numerals', () => {
    expect([1, 10, 11, 20, 101, 110, 1234].map((n) => formatNumber(n, 'chineseCounting'))).toEqual(['一', '十', '十一', '二十', '一百零一', '一百一十', '一千二百三十四']);
    // Japanese counting drops the one before 十, 百 and 千 and writes no zeros.
    expect([11, 101, 1234].map((n) => formatNumber(n, 'japaneseCounting'))).toEqual(['十一', '百一', '千二百三十四']);
    expect([10, 1234].map((n) => formatNumber(n, 'ideographLegalTraditional'))).toEqual(['壹拾', '壹仟貳佰參拾肆']);
    expect(formatNumber(1234, 'chineseLegalSimplified')).toBe('壹仟贰佰叁拾肆');
  });

  it('writes digits, full-width digits and cyclic signs', () => {
    expect(formatNumber(120, 'ideographDigital')).toBe('一二〇');
    expect(formatNumber(12, 'decimalFullWidth')).toBe('１２');
    expect(formatNumber(3, 'ideographEnclosedCircle')).toBe('㈢');
    expect([1, 10, 11].map((n) => formatNumber(n, 'ideographTraditional'))).toEqual(['甲', '癸', '甲']);
    expect([1, 12, 13].map((n) => formatNumber(n, 'ideographZodiac'))).toEqual(['子', '亥', '子']);
  });
});

describe('right-to-left text', () => {
  const cps = (s: string) => [...s].map((c) => c.codePointAt(0)!);

  it('joins Arabic letters, with lam and alef as one ligature', () => {
    // سلام: seen initial, lam-alef final (the alef folds into it), meem isolated after the alef.
    expect(arabicForms(cps('سلام'))).toEqual([0xfeb3, 0xfefc, 0, 0xfee1]);
  });

  it('orders a right-to-left run after left-to-right text, and keeps numbers left to right', () => {
    const ltr = cps('abc אבג');
    const levels = bidiLevels(ltr.map(bidiClass), ltr.map(() => false), 0);
    expect(levels).toEqual([0, 0, 0, 0, 1, 1, 1]);
    expect(visualOrder(levels)).toEqual([0, 1, 2, 3, 6, 5, 4]);

    const rtl = cps('ابج 123');
    const rl = bidiLevels(rtl.map(bidiClass), rtl.map(() => false), 1);
    expect(visualOrder(rl)).toEqual([4, 5, 6, 3, 2, 1, 0]);
  });

  it('starts the lines of a right-to-left paragraph at the right margin', async () => {
    const r = await layout(makeDocx({ body: p('שלום עולם', '<w:bidi/>') }));
    const cs = chars(r.pages[0].body);
    // The first letter is the rightmost one.
    expect(Math.max(...cs.map((c) => c.x))).toBe(cs[0].x);
    expect(cs[cs.length - 1].x).toBeLessThan(cs[0].x);
  });
});

describe('notes, frames and substitutes', () => {
  it('stacks endnotes below one another at the end of the document', async () => {
    const ref = (id: number) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
    const note = (id: number, text: string) => `<w:endnote w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:endnote>`;
    const endnotes =
      '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
      note(1, 'First note') +
      note(2, 'Second note');
    const r = await layout(makeDocx({ endnotes, body: `<w:p><w:r><w:t>Text</w:t></w:r>${ref(1)}${ref(2)}</w:p>` }));
    const ls = lines(r.pages[0].body);
    const first = ls.find((l) => l.text.includes('First note'))!;
    const second = ls.find((l) => l.text.includes('Second note'))!;
    expect(first.y).toBeGreaterThan(ls[0].y);
    expect(second.y - first.y).toBeGreaterThan(8);
  });

  it('gives an empty page-anchored framed paragraph no room in a footer', async () => {
    const refs = '<w:footerReference w:type="default" r:id="ftr1"/>';
    const framed = '<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="5000" w:y="15800"/></w:pPr></w:p>';
    const footerY = async (footer: string) => lines((await layout(makeDocx({ sectPr: sect('', refs), footers: { ftr1: footer }, body: p('Body') }))).pages[0].headerFooter)[0].y;
    // The footer sits on its last line, so a line taken by the frame would lift the text.
    expect(await footerY(p('Footer') + framed)).toBeCloseTo(await footerY(p('Footer')), 2);
  });

  it('draws David with Liberation Serif at its own line height', async () => {
    const rPr = '<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/><w:sz w:val="24"/>';
    const styles = '<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';
    const r = await layout(makeDocx({ styles, body: p('One', '', rPr) + p('Two', '', rPr) }));
    const ls = lines(r.pages[0].body);
    expect(ls[1].y - ls[0].y).toBeCloseTo(12 * (0.735 + 0.249), 2);
    expect(ls[0].ops[0].face.font.names.family).toBe('Liberation Serif');
  });
});
