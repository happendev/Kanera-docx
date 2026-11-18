/**
 * Layout rules that Word applies and that page breaks depend on. Expected numbers come from
 * Word's behaviour with the metric-compatible bundled fonts (Liberation Sans for Arial:
 * winAscent 1854, winDescent 434, lineGap 67 per 2048 units).
 */
import { describe, expect, it } from 'vitest';
import { makeDocx, p } from './helpers/docx.js';
import { glyphOps, layout, lines, pageText } from './helpers/layout.js';

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

  it('scales the line by a "multiple" spacing and adds paragraph spacing', async () => {
    const body = p('One', '<w:spacing w:line="360" w:lineRule="auto" w:after="200"/>') + p('Two') + p('Three', '<w:spacing w:before="100" w:line="480" w:lineRule="exact"/>');
    const ys = lines((await layout(makeDocx({ styles: ARIAL_STYLES, body }))).pages[0].body).map((l) => l.y);
    const descent = (434 / 2048) * 11;
    // Text sits at the bottom of the line box, so the extra half line goes above "One".
    expect(ys[0]).toBeCloseTo(72 + ARIAL_11 * 1.5 - descent, 1);
    // 10pt after, then a single line.
    expect(ys[1] - ys[0]).toBeCloseTo(descent + 10 + ARIAL_11 - descent, 1);
    // 5pt before, then an exact 24pt line.
    expect(ys[2] - ys[1]).toBeCloseTo(descent + 5 + 24 - descent, 1);
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
