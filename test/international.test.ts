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
  it('switches the primary face across script slots within one long run', async () => {
    const text = 'Aé中A'.repeat(12);
    const props = '<w:rFonts w:ascii="Carlito" w:hAnsi="Caladea" w:eastAsia="Noto Sans CJK JP"/>';
    const doc = makeDocx({ body: p(text, '', props) });
    const { pages } = await layout(doc);
    const glyphs = glyphOps(pages[0].body).flatMap((op) => op.texts.map((char) => [char, op.face.familyName]));
    expect(glyphs.map(([char]) => char).join('')).toBe(text);
    const families = new Map([['A', 'Carlito'], ['é', 'Caladea'], ['中', 'Noto Sans CJK JP']]);
    for (const [char, family] of glyphs) expect(family).toBe(families.get(char));
  });

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

  it('gives a line a second grid line once it is 0.5% taller than one', async () => {
    // As in Word's PDFs: 14pt SimHei (18.16pt) takes two 18pt grid lines, a picture 18.003pt tall
    // (228637 EMU) one.
    const title = '<w:p><w:r><w:rPr><w:rFonts w:ascii="SimHei" w:hAnsi="SimHei" w:eastAsia="SimHei"/><w:sz w:val="28"/></w:rPr><w:t>标题</w:t></w:r></w:p>';
    const picture = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="228637"/><a:graphic><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:prstGeom prst="rect"/>' +
      '<a:solidFill><a:srgbClr val="0077AA"/></a:solidFill></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const r = await layout(makeDocx({ styles: simsun(21), sectPr: sect('<w:docGrid w:type="lines" w:linePitch="360"/>'), body: title + p('中文') + picture + p('中文') }));
    const ys = lines(r.pages[0].body).map((l) => l.y);
    const h = (10.5 * 332) / 256;
    expect(ys[1]).toBeCloseTo(72 + 36 + (18 - h) / 2 + h - (10.5 * 74) / 256, 2);
    expect(ys[2] - ys[1]).toBeCloseTo(36, 2);
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

describe('vertical text', () => {
  it('sets a vertical section down the page from the right margin, its kana upright', async () => {
    // Landscape A4 with 1" margins, its text running top to bottom in lines that move leftwards.
    const sectPr = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:textDirection w:val="tbRl"/></w:sectPr>';
    const page = (await layout(makeDocx({ sectPr, body: p('あいう') }))).pages[0];
    expect(page.width).toBeCloseTo(841.9, 1);
    expect(page.body[0]).toEqual({ kind: 'transform', matrix: [0, 1, -1, 0, page.width, 0] });
    // Each kana is turned back about the centre of its em box; where those centres land on the page:
    const centres = page.body.flatMap((o) => (o.kind === 'transform' && o.matrix[1] === -1 ? [{ x: page.width - (o.matrix[5] - o.matrix[4]) / 2, y: (o.matrix[4] + o.matrix[5]) / 2 }] : []));
    expect(centres).toHaveLength(3);
    // The line runs down from the top margin, just inside the right margin.
    const size = page.body.find((o) => o.kind === 'glyphs')?.size ?? 0;
    expect(centres[0].y).toBeCloseTo(72 + size / 2, 1);
    expect(centres[1].y - centres[0].y).toBeCloseTo(size, 1);
    expect(centres[0].x).toBeLessThan(page.width - 72);
    expect(centres[0].x).toBeGreaterThan(page.width - 72 - 20);
    expect(centres[1].x).toBeCloseTo(centres[0].x, 3);
  });
});

describe('right-to-left text', () => {
  const cps = (s: string) => [...s].map((c) => c.codePointAt(0)!);

  // Liberation Sans, Arial's stand-in, has Hebrew: a Hebrew run is shaped as one.
  const arial = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr></w:rPrDefault></w:docDefaults>';

  it('underlines and links a right-to-left run where it is drawn', async () => {
    const body = '<w:p><w:pPr><w:bidi/></w:pPr><w:hyperlink w:anchor="here"><w:r><w:rPr><w:rtl/><w:u w:val="single"/></w:rPr><w:t>שמש דגל</w:t></w:r></w:hyperlink>' +
      '<w:bookmarkStart w:id="0" w:name="here"/><w:bookmarkEnd w:id="0"/></w:p>';
    const ops = (await layout(makeDocx({ styles: arial, body }))).pages[0].body;
    expect(glyphOps(ops)).toHaveLength(1);
    const xs = chars(ops).map((c) => c.x);
    const rules = ops.filter((o) => o.kind === 'line');
    const links = ops.filter((o) => o.kind === 'link');
    expect(rules).toHaveLength(1);
    expect(links).toHaveLength(1);
    const rule = rules[0].kind === 'line' ? rules[0] : null;
    expect(rule!.x1).toBeCloseTo(Math.min(...xs), 1);
    expect(rule!.x2).toBeGreaterThan(Math.max(...xs));
    expect(links[0].kind === 'link' && links[0].w).toBeCloseTo(rule!.x2 - rule!.x1, 1);
  });

  it('underlines the pieces of a run that right-to-left text sets apart, not the text between them', async () => {
    // The run "abc שמש" is followed by " דגל": once the Hebrew is reordered the second run's word
    // sits between the first run's Latin and Hebrew.
    const body = '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">abc שמש</w:t></w:r><w:r><w:t xml:space="preserve"> דגל</w:t></w:r></w:p>';
    const ops = (await layout(makeDocx({ styles: arial, body }))).pages[0].body;
    const cs = chars(ops);
    const rules = ops.flatMap((o) => (o.kind === 'line' ? [o] : []));
    const covered = (t: string) => cs.filter((c) => c.t === t).every((c) => rules.some((r) => c.x > r.x1 - 0.5 && c.x < r.x2 - 0.5));
    expect(rules).toHaveLength(2);
    for (const t of 'abcשמ') expect(covered(t), t).toBe(true);
    for (const t of 'דגל') expect(cs.filter((c) => c.t === t).some((c) => rules.some((r) => c.x > r.x1 + 0.5 && c.x < r.x2 - 0.5)), t).toBe(false);
  });

  it('joins Arabic letters, with lam and alef as one ligature', () => {
    // سلام: seen initial, lam-alef final (the alef folds into it), meem isolated after the alef.
    expect(arabicForms(cps('سلام'))).toEqual([0xfeb3, 0xfefc, 0, 0xfee1]);
  });

  it('joins Persian letters too', () => {
    // تکمیل: teh initial, keheh and meem medial, farsi yeh medial, lam final.
    expect(arabicForms(cps('تکمیل'))).toEqual([0xfe97, 0xfb91, 0xfee4, 0xfbff, 0xfede]);
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

  it('keeps a decimal number whole in right-to-left text, with its brackets mirrored', async () => {
    // A lone separator between digits joins them (rule W4); the brackets read right to left and so
    // face the other way, as Word draws "(0.25)" in a right-to-left table cell.
    const r = await layout(makeDocx({ body: p('(0.25)', '<w:bidi/>') }));
    const visual = chars(r.pages[0].body)
      .sort((a, b) => a.x - b.x)
      .map((c) => c.t)
      .join('');
    expect(visual).toBe('(0.25)');
  });

  it('draws Arabic that Arial has in a fallback at Arial\'s width, on Arial\'s lines', async () => {
    const ar = (t: string) => p(t, '<w:bidi/>', '<w:rFonts w:cs="Arial"/><w:rtl/>');
    const latin = (t: string) => p(t, '', '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>');
    const r = await layout(makeDocx({ body: ar('الرقم الاستدلالي') + ar('للأثمان') + latin('One') + latin('Two') }));
    const ls = lines(r.pages[0].body);
    const arabic = ls[0].ops.find((o) => o.texts.some((t) => /[\u0600-\u06ff]/.test(t)))!;
    expect(arabic.face.familyName).toBe('DejaVu Sans');
    // DejaVu Sans's Arabic is about 1.3 times as wide as Arial's, so it is drawn narrower.
    expect(arabic.scaleX).toBeCloseTo(0.766, 2);
    expect(ls[1].y - ls[0].y).toBeCloseTo(ls[3].y - ls[2].y, 2);
  });

  it('lays out a right-to-left table from the right margin, its first column on the right', async () => {
    const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
    const table = `<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>`;
    const cs = chars((await layout(makeDocx({ body: table + p(''), sectPr: sect() }))).pages[0].body);
    // A4's text ends at 523.3pt; the two 100pt columns stand against it, A's on the right.
    expect(xOf(cs, 'A')).toBeGreaterThan(523.3 - 100 - 0.01);
    expect(xOf(cs, 'B')).toBeGreaterThan(523.3 - 200 - 0.01);
    expect(xOf(cs, 'B')).toBeLessThan(523.3 - 100);
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

  it("sets a proportional Japanese font's kana and punctuation in its stand-in's proportional forms", async () => {
    const span = async (font: string) => {
      const rPr = `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}"/><w:sz w:val="20"/>`;
      const cs = chars((await layout(makeDocx({ body: p('あ、いあ', '', rPr) }))).pages[0].body);
      return cs[3].x - cs[0].x;
    };
    // MS Gothic sets each an em wide; MS PGothic's comma takes about half an em.
    expect(await span('MS Gothic')).toBeCloseTo(30, 2);
    expect(await span('MS PGothic')).toBeLessThan(26);
  });

  it('draws an HG font named for a bold weight bold', async () => {
    const op = async (font: string) => {
      const rPr = `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}"/>`;
      return glyphOps((await layout(makeDocx({ body: p('文字', '', rPr) }))).pages[0].body)[0];
    };
    expect((await op('HG創英角ｺﾞｼｯｸUB')).syntheticBold).toBe(true);
    expect((await op('HGｺﾞｼｯｸE')).syntheticBold).toBe(true);
    expect((await op('HGｺﾞｼｯｸM')).syntheticBold).toBe(false);
  });

  it("draws Century with Gelasio set to Century's wider widths", async () => {
    const op = glyphOps((await layout(makeDocx({ body: p('Text', '', '<w:rFonts w:ascii="Century" w:hAnsi="Century"/>') }))).pages[0].body)[0];
    expect(op.face.familyName).toBe('Gelasio');
    expect(op.scaleX).toBeCloseTo(1.048, 3);
  });

  it("draws Consolas with Liberation Mono narrowed to Consolas's cells", async () => {
    const op = glyphOps((await layout(makeDocx({ body: p('Code', '', '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>') }))).pages[0].body)[0];
    expect(op.face.familyName).toBe('Liberation Mono');
    expect(op.scaleX).toBeCloseTo(1126 / 1229, 4);
  });
});
