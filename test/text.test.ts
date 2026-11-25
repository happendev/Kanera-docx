/**
 * Text and Markdown extraction: what a reader sees in Word, in reading order, with Word's structure
 * mapped onto Markdown's.
 */
import { describe, expect, it } from 'vitest';
import { docxToMarkdown as markdownCore, extractText as extractCore } from '../src/index.js';
import { docxToMarkdown as markdownInBrowser } from '../src/browser.js';
import { docxToMarkdown, docxToText } from '../src/node.js';
import { makeDocx, p } from './helpers/docx.js';
import { inlineImage, redPng, sampleDocx } from './helpers/sample.js';

const M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const V_NS = 'xmlns:v="urn:schemas-microsoft-com:vml"';

const run = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (inner: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inner}</w:p>`;
const cell = (inner: string, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${inner}</w:tc>`;
const table = (rows: string[][], grid = rows[0].length) =>
  `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="2000"/>'.repeat(grid)}</w:tblGrid>${rows.map((r) => `<w:tr>${r.join('')}</w:tr>`).join('')}</w:tbl>`;

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Kop3"><w:name w:val="heading 3"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/></w:style>';

const NUMBERING =
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>' +
  '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/><w:suff w:val="space"/></w:lvl>' +
  '</w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';

const numPr = (numId: number, ilvl = 0) => `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

describe('plain text', () => {
  it('writes each paragraph on a line, with tabs, breaks and Word’s special characters', async () => {
    const body =
      p('First paragraph') +
      para(run('a') + '<w:r><w:tab/></w:r>' + run('b') + '<w:r><w:br/></w:r>' + run('c')) +
      para(run('co') + '<w:r><w:softHyphen/></w:r>' + run('operate') + ' ' + '<w:r><w:noBreakHyphen/></w:r>' + run('x')) +
      p('') +
      p('') +
      p('Last');
    expect(await docxToText(makeDocx({ body }))).toBe('First paragraph\na\tb\nc\ncooperate-x\n\nLast\n');
  });

  it('shows tracked insertions and drops deletions, hidden text and field codes', async () => {
    const body = para(
      run('Kept ') +
        '<w:ins w:id="1" w:author="A"><w:r><w:t xml:space="preserve">inserted </w:t></w:r></w:ins>' +
        '<w:del w:id="2" w:author="A"><w:r><w:delText>deleted </w:delText></w:r></w:del>' +
        run('secret ', '<w:vanish/>') +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE \\@ "d MMMM yyyy" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        run('3 March 2024') +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    );
    expect(await docxToText(makeDocx({ body }))).toBe('Kept inserted 3 March 2024\n');
  });

  it('gives list items Word’s numbers and bullets, with symbol-font bullets as Unicode', async () => {
    const body =
      p('One', numPr(1)) + p('Sub', numPr(1, 1)) + p('Deep', numPr(1, 2)) + p('Two', numPr(1)) + p('Bullet', numPr(2)) + p('Nested', numPr(2, 1));
    expect(await docxToText(makeDocx({ body, numbering: NUMBERING }))).toBe('1.\tOne\na)\tSub\n1.a.1 Deep\n2.\tTwo\n•\tBullet\no\tNested\n');
  });

  it('maps symbol-font characters (w:sym and Wingdings text) to Unicode', async () => {
    const body = para('<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>' + run(' done ') + run('l', '<w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/>'));
    expect(await docxToText(makeDocx({ body }))).toBe('✓ done ⚫\n');
  });

  it('separates table cells with tabs and rows with lines', async () => {
    const body = table([
      [cell(p('Name')), cell(p('Qty')), cell(p('Note'))],
      [cell(p('Apples')), cell(p('3')), cell(p('Red') + p('Crisp'))],
      [cell(p('Merged'), '<w:gridSpan w:val="2"/>'), cell(p('x'))],
    ]);
    expect(await docxToText(makeDocx({ body }))).toBe('Name\tQty\tNote\nApples\t3\tRed Crisp\nMerged\t\tx\n');
  });

  it('numbers footnotes and endnotes in reference order and lists them after the body', async () => {
    const body = para(run('A') + '<w:r><w:footnoteReference w:id="7"/></w:r>' + run(' B') + '<w:r><w:endnoteReference w:id="3"/></w:r>' + run(' C') + '<w:r><w:footnoteReference w:id="2"/></w:r>');
    const note = (tag: string, id: number, text: string) => `<w:${tag} w:id="${id}"><w:p><w:r><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${tag}>`;
    const docx = makeDocx({ body, footnotes: note('footnote', 2, 'Second.') + note('footnote', 7, 'First.'), endnotes: note('endnote', 3, 'End.') });
    expect(await docxToText(docx)).toBe('A[1] B[i] C[2]\n\n[1] First.\n[2] Second.\n[i] End.\n');
    expect(await docxToText(docx, { includeNotes: false })).toBe('A[1] B[i] C[2]\n');
    expect(await docxToMarkdown(docx)).toBe('A[^1] B[^i] C[^2]\n\n[^1]: First.\n\n[^2]: Second.\n\n[^i]: End.\n');
  });

  it('includes headers and footers only when asked', async () => {
    const docx = sampleDocx();
    const plain = await docxToText(docx);
    expect(plain).not.toContain('ACME');
    const withHf = await docxToText(docx, { includeHeadersFooters: true });
    expect(withHf.startsWith('ACME Corp — Confidential\nQuarterly report\n')).toBe(true);
    expect(withHf).toContain('Second page\nPage 1\n');
  });

  it('writes linear equations and text box content after its paragraph', async () => {
    const math =
      `<m:oMath ${M_NS}><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>+</m:t></m:r>` +
      '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b+c</m:t></m:r></m:den></m:f></m:oMath>';
    const textbox = `<w:r><w:pict ${V_NS}><v:shape style="width:100pt;height:40pt"><v:textbox><w:txbxContent>${p('Boxed')}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
    const body = para(run('Area: ') + math) + para(run('Before') + textbox + run(' after'));
    expect(await docxToText(makeDocx({ body }))).toBe('Area: x^2+a/(b+c)\nBefore after\nBoxed\n');
  });

  it('joins a dropped capital to its paragraph', async () => {
    const body = p('O', '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>') + p('nce upon a time');
    expect(await docxToText(makeDocx({ body }))).toBe('Once upon a time\n');
  });
});

describe('Markdown', () => {
  it('maps heading styles, titles and outline levels to headings', async () => {
    const body =
      p('Report', '<w:pStyle w:val="Title"/>') +
      p('Intro', '<w:pStyle w:val="Heading1"/>') +
      p('Scope', '<w:pStyle w:val="Heading2"/>') +
      p('Localised id', '<w:pStyle w:val="Kop3"/>') +
      p('Outline only', '<w:outlineLvl w:val="3"/>') +
      p('Body text.');
    expect(await docxToMarkdown(makeDocx({ body, styles: STYLES }))).toBe('# Report\n\n# Intro\n\n## Scope\n\n### Localised id\n\n#### Outline only\n\nBody text.\n');
  });

  it('nests emphasis properly and keeps markers off whitespace', async () => {
    const body = para(run('Plain ') + run('bold ', '<w:b/>') + run('both', '<w:b/><w:i/>') + run(' bold', '<w:b/>') + run(' and ') + run(' italic ', '<w:i/>') + run('x') + run('gone', '<w:strike/>') + run('2', '<w:vertAlign w:val="superscript"/>'));
    expect(await docxToMarkdown(makeDocx({ body }))).toBe('Plain **bold *both* bold** and  *italic* x~~gone~~<sup>2</sup>\n');
  });

  it('keeps punctuation and spaces at the edge of emphasis where Markdown can close it', async () => {
    const body =
      para(run('Cinema') + run('.', '<w:b/>')) +
      para(run('word,', '<w:b/>') + run('next')) +
      para(run('bold ', '<w:b/>') + run('x')) +
      para(run('Visit ', '<w:b/>') + '<w:hyperlink r:id="L1">' + run('https://example.com/a', '<w:b/>') + '</w:hyperlink>' + run(' today', '<w:b/>'));
    expect(await docxToMarkdown(makeDocx({ body, links: { L1: 'https://example.com/a' } }))).toBe(
      'Cinema.\n\n**word**,next\n\n**bold** x\n\n**Visit <https://example.com/a> today**\n',
    );
  });

  it('takes emphasis and monospace from the face of the text’s script', async () => {
    const body =
      para(run('مرحبا', '<w:b/><w:rtl/>')) +
      para(run('مرحبا', '<w:bCs/><w:rtl/>')) +
      para(run('عنوان', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Arial"/><w:rtl/>'));
    expect(await docxToMarkdown(makeDocx({ body }))).toBe('**مرحبا**\n\n**مرحبا**\n\nعنوان\n');
  });

  it('writes the Plain Text style as code only in a monospaced font', async () => {
    const styles = '<w:style w:type="paragraph" w:styleId="PlainText"><w:name w:val="Plain Text"/></w:style>';
    const body = p('Title text', '<w:pStyle w:val="PlainText"/>') + p('x = 1', '<w:pStyle w:val="PlainText"/>', '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>');
    expect(await docxToMarkdown(makeDocx({ body, styles }))).toBe('Title text\n\n```\nx = 1\n```\n');
  });

  it('keeps ruby readings', async () => {
    const ruby = '<w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t>かん</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢</w:t></w:r></w:rubyBase></w:ruby></w:r>';
    const docx = makeDocx({ body: para(ruby + run('字')) });
    expect(await docxToMarkdown(docx)).toBe('<ruby>漢<rp>（</rp><rt>かん</rt><rp>）</rp></ruby>字\n');
    expect(await docxToText(docx)).toBe('漢（かん）字\n');
  });

  it('indents nested list items under the text of their parent, whatever its label', async () => {
    const body = p('One', numPr(1)) + p('Sub', numPr(1, 1)) + p('Deep', numPr(1, 2)) + p('Two', numPr(1));
    expect(await docxToMarkdown(makeDocx({ body, numbering: NUMBERING }))).toBe('1. One\n   - a) Sub\n     - 1.a.1 Deep\n2. Two\n');
  });

  it('writes lists with their numbers, nesting and bullets', async () => {
    const body = p('One', numPr(1)) + p('Sub', numPr(1, 1)) + p('Two', numPr(1)) + p('After') + p('Bullet', numPr(2)) + p('Nested', numPr(2, 1));
    expect(await docxToMarkdown(makeDocx({ body, numbering: NUMBERING }))).toBe('1. One\n   - a) Sub\n2. Two\n\nAfter\n\n- Bullet\n  - Nested\n');
  });

  it('writes tables as GitHub tables, unwraps single-cell boxes and falls back to HTML for nested tables', async () => {
    const gfm = table([
      [cell(p('Name')), cell(p('A|B'))],
      [cell(p('Line 1') + p('Line 2')), cell(para(run('bold', '<w:b/>')))],
    ]);
    const box = table([[cell(p('Note: boxed text'))]]);
    const nested = table([[cell(p('Outer')), cell(table([[cell(p('In 1')), cell(p('In 2'))]]))], [cell(p('x')), cell(p('y'))]]);
    expect(await docxToMarkdown(makeDocx({ body: gfm + box + nested }))).toBe(
      '| Name | A\\|B |\n| --- | --- |\n| Line 1<br>Line 2 | **bold** |\n\nNote: boxed text\n\n' +
        '<table>\n<tr><td>Outer</td><td><table>\n<tr><td>In 1</td><td>In 2</td></tr>\n</table></td></tr>\n<tr><td>x</td><td>y</td></tr>\n</table>\n',
    );
  });

  it('keeps vertically merged and spanned cells in their columns', async () => {
    const body = table([
      [cell(p('A'), '<w:vMerge w:val="restart"/>'), cell(p('B'))],
      [cell(p(''), '<w:vMerge/>'), cell(p('C'))],
      [cell(p('Wide'), '<w:gridSpan w:val="2"/>')],
    ]);
    expect(await docxToMarkdown(makeDocx({ body }))).toBe('| A | B |\n| --- | --- |\n|  | C |\n| Wide |  |\n'.replace(/ {2}\|/g, ' |'));
  });

  it('writes links, and leaves internal links as text', async () => {
    const body =
      para('<w:hyperlink r:id="L1">' + run('the site', '<w:b/>') + '</w:hyperlink>' + run(' and ') + '<w:hyperlink w:anchor="_Toc1">' + run('a section') + '</w:hyperlink>') +
      para('<w:hyperlink r:id="L2">' + run('https://example.com/a') + '</w:hyperlink>') +
      para('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK "https://example.com/x y" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run('field link') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>');
    const docx = makeDocx({ body, links: { L1: 'https://example.com/(site)', L2: 'https://example.com/a' } });
    expect(await docxToMarkdown(docx)).toBe('[**the site**](https://example.com/%28site%29) and a section\n\n<https://example.com/a>\n\n[field link](https://example.com/x%20y)\n');
  });

  it('escapes text that Markdown would read as formatting', async () => {
    const body = p('# not a heading') + p('1. not a list') + p('- nor this') + p('a *b* [c](d) _e_ snake_case <tag> \\') + p('---') + p('Name: ____ Yes____No');
    expect(await docxToMarkdown(makeDocx({ body }))).toBe(
      '\\# not a heading\n\n1\\. not a list\n\n\\- nor this\n\na \\*b\\* [c\\](d) \\_e\\_ snake_case \\<tag> \\\\\n\n\\---\n\nName: \\_\\_\\_\\_ Yes\\_\\_\\_\\_No\n',
    );
  });

  it('writes equations as LaTeX', async () => {
    const inline = `<m:oMath ${M_NS}><m:r><m:t>α≤</m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>`;
    const display =
      `<m:oMathPara ${M_NS}><m:oMath><m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup>` +
      '<m:e><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:den></m:f></m:e></m:nary></m:oMath></m:oMathPara>';
    const body = para(run('Then ') + inline + run(' holds.')) + para(display);
    expect(await docxToMarkdown(makeDocx({ body }))).toBe('Then $\\alpha \\le \\sqrt{x}$ holds.\n\n$$\n\\sum_{i=1}^n \\frac{1}{x_i}\n$$\n');
  });

  it('shows pictures with their alternative text', async () => {
    const image = inlineImage('img1', 20, 10).replace('<wp:docPr id="1" name="Picture 1"/>', '<wp:docPr id="1" name="Picture 1" descr="A red [box]"/>');
    const docx = makeDocx({ body: para(run('See ') + image), media: { img1: { name: 'red.png', data: redPng() } } });
    expect(await docxToMarkdown(docx)).toBe('See ![A red \\[box\\]](word/media/red.png)\n');
    expect(await docxToMarkdown(docx, { includeImages: false })).toBe('See\n');
  });

  it('writes block quotes, tables of contents and monospaced paragraphs as code', async () => {
    const mono = '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>';
    const body =
      p('Wise words', '<w:pStyle w:val="Quote"/>') +
      p('Intro\t1', '<w:pStyle w:val="TOC1"/>') +
      p('Detail\t2', '<w:pStyle w:val="TOC2"/>') +
      p('const a = 1;', '', mono) +
      p('  return a * 2;', '', mono) +
      para(run('Call ') + run('run()', mono) + run(' now.'));
    expect(await docxToMarkdown(makeDocx({ body, styles: STYLES }))).toBe('> Wise words\n\n- Intro 1\n  - Detail 2\n\n```\nconst a = 1;\n  return a * 2;\n```\n\nCall `run()` now.\n');
  });

  it('converts the sample document through every entry point alike', async () => {
    const expected =
      '# Quarterly report\n\nPlain, **bold,** *italic,* underlined red\n\nGrüße — “quoted” € 42 ½ Ωmega\n\n- First bullet\n- Second bullet\n\n' +
      '| Region | Revenue |\n| --- | --- |\n| North | 1,204 |\n\n![](word/media/red.png)\n\n[Visit example.com](https://example.com/) and a note[^1]\n\nSecond page\n\n[^1]: Footnote text.\n';
    const docx = sampleDocx();
    expect(await docxToMarkdown(docx)).toBe(expected);
    expect(await markdownCore(docx)).toBe(expected);
    expect(await markdownInBrowser(new Blob([docx as Uint8Array<ArrayBuffer>]))).toBe(expected);
    expect((await extractCore(docx, { format: 'markdown' })).text).toBe(expected);
  });
});

it('returns an empty string for an empty document', async () => {
  expect(await docxToText(makeDocx({ body: '' }))).toBe('');
  expect(await docxToMarkdown(makeDocx({ body: p('') }))).toBe('');
});
