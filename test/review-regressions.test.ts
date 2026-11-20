import { nodeCompression } from '../src/util/node-compression.js';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { convertDocx } from '../src/node.js';
import { convertDocx as browserConvert } from '../src/browser.js';
import { convertDocx as coreConvert } from '../src/index.js';
import { parseDocx } from '../src/docx/parser.js';
import { deobfuscateFont } from '../src/fonts/odttf.js';
import { FontRegistry } from '../src/fonts/registry.js';
import { SfntFont } from '../src/fonts/sfnt.js';
import { cjkFallbacks } from '../src/fonts/substitutes.js';
import { ZipArchive } from '../src/util/zip.js';
import { makeDocx, p } from './helpers/docx.js';
import { layout, lines, pageText } from './helpers/layout.js';
import { ghostscript, pdfPagesText, pdfStreams, pdfText } from './helpers/pdf.js';

const font = (file: string) => readFileSync(new URL('../fonts/' + file, import.meta.url));
const arial = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>';
const field = (name: string) => '<w:fldSimple w:instr="' + name + '"><w:r><w:t>1</w:t></w:r></w:fldSimple>';
const row = (body: string, pr = '') => '<w:tr><w:trPr>' + pr + '</w:trPr><w:tc>' + body + '</w:tc></w:tr>';
const table = (rows: string) => '<w:tbl><w:tblGrid><w:gridCol w:w="5400"/></w:tblGrid>' + rows + '</w:tbl>';

describe('document content and pagination', () => {
  it('uses defaults for invalid Word property values', () => {
    const body = p('Text', '<w:tabs><w:tab w:val="bogus" w:leader="bogus" w:pos="720"/>' +
      '<w:tab w:val="end" w:leader="dot" w:pos="1440"/></w:tabs>');
    const sectPr = '<w:sectPr><w:headerReference w:type="bogus" r:id="header"/>' +
      '<w:type w:val="bogus"/><w:vAlign w:val="bogus"/>' +
      '<w:footnotePr><w:numRestart w:val="bogus"/></w:footnotePr>' +
      '<w:pgBorders w:display="bogus"><w:top w:val="single" w:sz="8"/></w:pgBorders></w:sectPr>';
    const docx = makeDocx({ body, sectPr, headers: { header: p('Header') } });
    const doc = parseDocx(docx, nodeCompression.inflateRaw);
    const para = doc.body[0];
    expect(para?.type === 'paragraph' && para.props.tabs).toEqual([
      { type: 'left', leader: 'none', pos: 36 },
      { type: 'right', leader: 'dot', pos: 72 },
    ]);
    expect(doc.finalSection).toMatchObject({ type: 'nextPage', vAlign: 'top', footnoteRestart: 'continuous', pageBorders: { display: 'allPages' } });
    expect(doc.finalSection.headers).toEqual({});
  });

  it('imports HTML text in order and reports unsupported imported parts', async () => {
    const doc = makeDocx({
      body: p('Before') + '<w:altChunk r:id="html"/><w:altChunk r:id="rtf"/>' + p('After'),
      imports: {
        html: { name: 'chunk1.dat', contentType: 'text/html', text: '<html><head><title>Hidden</title></head><body><p>Payment is <b>due</b> in 30 days.</p><p>A &amp; B</p></body></html>' },
        rtf: { name: 'chunk2.rtf', contentType: 'application/rtf', text: 'unsupported' },
      },
    });
    expect(pageText((await layout(doc)).pages[0])).toBe('Before\nPayment is due in 30 days.\nA & B\nAfter');
    const { warnings } = await convertDocx(doc);
    expect(warnings).toContain('Skipped unsupported or missing altChunk import: word/chunk2.rtf');
    expect(warnings).toContain('Imported HTML altChunk as plain text: word/chunk1.dat');
  });

  it('reports unsupported blocks containing document text', async () => {
    const result = await convertDocx(makeDocx({ body: '<ext:wrapper xmlns:ext="urn:example">' + p('Imported paragraph') + '</ext:wrapper>' }));
    expect(result.warnings.some((w) => w.includes('Skipped unsupported content block'))).toBe(true);
  });

  it('omits a tracked-deleted cell from the final document', async () => {
    const cells = '<w:tc>' + p('Current') + '</w:tc><w:tc><w:tcPr><w:cellDel w:id="1" w:author="Editor"/></w:tcPr>' + p('Withdrawn') + '</w:tc>';
    const doc = makeDocx({ body: table('<w:tr>' + cells + '</w:tr>') });
    expect(pageText((await layout(doc)).pages[0])).toBe('Current');
    if (ghostscript) expect(pdfPagesText((await convertDocx(doc)).pdf)[0].trim()).toBe('Current');
  });

  it('resolves bound values by namespace URI and preserves cached text for unknown prefixes', async () => {
    const store = '{A1B2C3D4-0000-0000-0000-000000000001}';
    const control = (prefix: string) => '<w:p><w:sdt><w:sdtPr><w:dataBinding w:prefixMappings="xmlns:b=\'urn:billing\'" w:xpath="/root[1]/' +
      prefix + ':name[1]" w:storeItemID="' + store + '"/></w:sdtPr><w:sdtContent><w:r><w:t>Saved</w:t></w:r></w:sdtContent></w:sdt></w:p>';
    const doc = makeDocx({ body: control('b') + control('unknown'), customXml: {
      [store]: '<root xmlns:a="urn:shipping"><a:name>Shipping</a:name><name xmlns="urn:billing">Billing</name></root>',
    } });
    expect(pageText((await layout(doc)).pages[0])).toBe('Billing\nSaved');
  });

  it('continues independent lists that share an abstract numbering definition', async () => {
    const instance = (id: number, start: number) => '<w:num w:numId="' + id + '"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="' + start + '"/></w:lvlOverride></w:num>';
    const para = (id: number) => p('Item', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + id + '"/></w:numPr>');
    const numbering = '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
      instance(1, 1) + instance(2, 10);
    const doc = makeDocx({ body: [1, 2, 1, 2].map(para).join(''), numbering });
    expect(pageText((await layout(doc)).pages[0]).match(/\d+\./g)).toEqual(['1.', '10.', '2.', '11.']);
  });

  it('rewraps just the remaining cell content as columns narrow and widen', async () => {
    const words = Array.from({ length: 120 }, (_, i) => 'Word' + String(i + 1).padStart(3, '0'));
    const sectPr = '<w:sectPr><w:pgSz w:w="9000" w:h="4300"/><w:pgMar w:top="900" w:bottom="900" w:left="600" w:right="600"/>' +
      '<w:cols w:num="2" w:equalWidth="0"><w:col w:w="1800" w:space="300"/><w:col w:w="5700"/></w:cols></w:sectPr>';
    const content = p(words.slice(0, 65).join(' ')) + p(words.slice(65).join(' '));
    const body = table(row(content)).replace('<w:tbl>', '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>');
    const { pages } = await layout(makeDocx({ body, sectPr }));
    expect(pages.length).toBeGreaterThan(1);
    const glyphs = pages.flatMap((page) => page.body.filter((op) => op.kind === 'glyphs'));
    expect(glyphs.map((op) => op.texts.join('')).join('').match(/Word\d{3}/g)).toEqual(words);
    expect(glyphs.some((op) => op.x >= 135 && (op.texts.join('').match(/Word/g)?.length ?? 0) > 2)).toBe(true);
  });

  it('restarts footnotes at the section start and inherits page restarts from document settings', async () => {
    const mark = (id: number) => '<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="' + id + '"/></w:r></w:p>';
    const note = (id: number) => '<w:footnote w:id="' + id + '"><w:p><w:r><w:footnoteRef/><w:t>Note</w:t></w:r></w:p></w:footnote>';
    const sect = (props: string) => '<w:sectPr><w:footnotePr>' + props + '</w:footnotePr></w:sectPr>';
    const first = sect('<w:numStart w:val="3"/><w:numRestart w:val="eachSect"/>');
    const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const doc = makeDocx({ body: mark(1) + p('', first) + mark(2) + mark(3) + pageBreak + mark(4),
      sectPr: sect(''), settings: '<w:footnotePr><w:numStart w:val="5"/><w:numRestart w:val="eachPage"/></w:footnotePr>',
      footnotes: [1, 2, 3, 4].map(note).join('') });
    const { pages, warnings } = await layout(doc);
    expect(warnings).toEqual([]);
    expect(pages.map((pg) => pageText(pg).match(/Body\d+|\d+Note/g))).toEqual([
      ['Body3', '3Note'], ['Body5', 'Body6', '5Note', '6Note'], ['Body5', '5Note'],
    ]);
    const continuous = await layout(makeDocx({
      body: mark(1) + p('', '<w:sectPr/>') + mark(2), sectPr: '<w:sectPr><w:type w:val="continuous"/></w:sectPr>',
      settings: '<w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr>', footnotes: note(1) + note(2),
    }));
    expect(continuous.pages.map(pageText)).toEqual(['Body1\nBody2\n1Note\n2Note']);
  });

  it('draws the prohibition symbol with two cutouts and a diagonal bar', async () => {
    const body = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><a:graphic><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr>' +
      '<a:prstGeom prst="noSmoking"/><a:solidFill><a:srgbClr val="0077AA"/></a:solidFill>' +
      '</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const result = await layout(makeDocx({ body }));
    const path = result.pages[0].body.find((op) => op.kind === 'path' && op.fill === '0077AA');
    expect(path?.kind === 'path' && path.segs.filter((s) => s.op === 'Z')).toHaveLength(3);
  });

  it('leaves unfilled custom subpaths out of the fill and unstroked ones out of the outline', async () => {
    const square = (attrs: string) => `<a:path w="10" h="10"${attrs}><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="0"/></a:lnTo>` +
      '<a:lnTo><a:pt x="10" y="10"/></a:lnTo><a:close/></a:path>';
    const body = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="1270000"/><a:graphic><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr>' +
      `<a:custGeom><a:pathLst>${square('')}${square(' fill="none"')}${square(' stroke="0"')}</a:pathLst></a:custGeom>` +
      '<a:solidFill><a:srgbClr val="0077AA"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></a:ln>' +
      '</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const ops = (await layout(makeDocx({ body }))).pages[0].body;
    const subpaths = (fill: string | null, stroke: string | null) => {
      const op = ops.find((o) => o.kind === 'path' && o.fill === fill && o.stroke === stroke);
      return op?.kind === 'path' ? op.segs.filter((s) => s.op === 'M').length : 0;
    };
    expect(subpaths('0077AA', null)).toBe(2);
    expect(subpaths(null, 'AA0000')).toBe(2);
  });

  it('excludes deleted rows before determining table layout', async () => {
    const doc = makeDocx({ body: table(row(p('Current')) + row(p('Withdrawn'), '<w:del w:id="1" w:author="Editor"/>')) });
    expect(pageText((await layout(doc)).pages[0])).toBe('Current');
  });

  it('resolves body fields after the break inside their paragraph', async () => {
    const body = '<w:p>' + field('NUMPAGES') + '</w:p><w:p><w:r><w:br w:type="page"/></w:r>' + field('PAGE') + '</w:p><w:p>' + field('SECTIONPAGES') + '</w:p>';
    expect((await layout(makeDocx({ body }))).pages.map(pageText)).toEqual(['2', '2\n2']);
  });

  it('uses Normal spacing when a stylesheet only declares run defaults', async () => {
    const normal = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
    const compact = '<w:spacing w:line="240" w:lineRule="auto" w:after="0"/>';
    const body = p('One') + p('Two') + p('Three', compact) + p('Four', compact);
    const result = await layout(makeDocx({ body, styles: arial + normal }));
    const [one, two, three, four] = lines(result.pages[0].body);
    expect(two.y - one.y).toBeCloseTo((four.y - three.y) * 1.15 + 8, 1);
  });

  it('updates PAGE fields separately in each repeated table header', async () => {
    const sectPr = '<w:sectPr><w:pgSz w:w="8000" w:h="3200"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr>';
    const header = row('<w:p>' + field('PAGE') + '</w:p>', '<w:tblHeader/>');
    const body = table(header + Array.from({ length: 24 }, (_, i) => row(p('Entry ' + i))).join(''));
    const result = await layout(makeDocx({ body, sectPr }));
    expect(result.pages.length).toBeGreaterThan(2);
    expect(result.pages.map((page) => pageText(page).split('\n')[0])).toEqual(result.pages.map((_, i) => String(i + 1)));
    expect(result.warnings).not.toContain('Page fields did not stabilize after repeated pagination');
  });

  it('moves the first unsplittable row to a fresh page', async () => {
    const sectPr = '<w:sectPr><w:pgSz w:w="8000" w:h="4600"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr>';
    const body = p('Before', '<w:spacing w:line="2880" w:lineRule="exact"/>') + table(row(p('First') + p('Second') + p('Third'), '<w:cantSplit/>'));
    expect((await layout(makeDocx({ styles: arial, body, sectPr }))).pages.map(pageText)).toEqual(['Before', 'First\nSecond\nThird']);
  });

  it('keeps every line of a multi-page footnote visible exactly once', async () => {
    const evidence = Array.from({ length: 150 }, (_, i) => 'Evidence line ' + (i + 1));
    const doc = makeDocx({ styles: arial, body: '<w:p><w:r><w:t>Reference</w:t><w:footnoteReference w:id="1"/></w:r></w:p>',
      footnotes: '<w:footnote w:id="1">' + evidence.map((s) => p(s)).join('') + '</w:footnote>' });
    const { pages } = await layout(doc);
    expect(pages.length).toBeGreaterThan(2);
    expect(pages.flatMap((page) => lines(page.body).map((line) => line.text)).filter((s) => s.startsWith('Evidence'))).toEqual(evidence);
    for (const page of pages) {
      for (const line of lines(page.body)) {
        expect(line.y).toBeGreaterThan(72);
        expect(line.y).toBeLessThan(page.height - 72);
      }
    }
  });
});

describe('PDF fonts, links and public inputs', () => {
  it('warns and falls back when a user font loader rejects with null', async () => {
    const family = 'Unavailable Font';
    const body = p('Fallback text', '', `<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>`);
    const { pdf, warnings } = await convertDocx(makeDocx({ body }), {
      fonts: [{ family, load: () => Promise.reject(null) }],
    });
    expect(warnings).toContain('Could not load font unavailable font: null');
    if (ghostscript) expect(pdfPagesText(pdf)[0]).toContain('Fallback text');
  });

  it('preserves caller-owned Node buffers on repeated conversions', async () => {
    const data = font('Caladea-Regular.ttf');
    const original = Buffer.from(data);
    const doc = makeDocx({ body: p('Identical outlines', '', '<w:rFonts w:ascii="Caladea"/>') });
    const options = { fonts: [{ data }] };
    const first = await coreConvert(doc, options);
    expect(data.equals(original)).toBe(true);
    const second = await coreConvert(doc, options);
    expect(Buffer.from(second.pdf).equals(first.pdf)).toBe(true);
    if (ghostscript) {
      const coverage = execFileSync('gs', ['-q', '-dBATCH', '-dNOPAUSE', '-dSAFER', '-sDEVICE=inkcov', '-o', '-', '-'], {
        input: second.pdf, encoding: 'utf8',
      });
      const channels = coverage.trim().split(/\s+/).slice(0, 4).map(Number);
      expect(channels.some((value) => value > 0)).toBe(true);
    }
  });

  it('deobfuscates an embedded font without writing through its Buffer view', () => {
    const source = font('Caladea-Regular.ttf');
    const original = Buffer.from(source);
    const key = '{00112233-4455-6677-8899-AABBCCDDEEFF}';
    const encoded = deobfuscateFont(source, key);
    expect(source.equals(original)).toBe(true);
    expect(original.equals(encoded)).toBe(false);
    expect(original.equals(deobfuscateFont(encoded, key))).toBe(true);
  });

  it.each([2, 0x200])('substitutes a font whose outline embedding flag is %i', async (flags) => {
    const data = font('Caladea-Regular.ttf');
    const table = new SfntFont(data).table('OS/2')!;
    data.writeUInt16BE(flags, table.offset + 8);
    const result = await convertDocx(makeDocx({ body: p('Text', '', '<w:rFonts w:ascii="Review font"/>') }), { fonts: [{ family: 'Review font', data }] });
    expect(result.warnings.some((w) => w.includes('disallows outline embedding'))).toBe(true);
    expect(pdfText(result.pdf)).not.toContain('/BaseFont /AAAAAB+Caladea');
  });

  it('embeds the original font when subsetting is disallowed', async () => {
    const data = font('Caladea-Regular.ttf');
    data.writeUInt16BE(0x100, new SfntFont(data).table('OS/2')!.offset + 8);
    const result = await coreConvert(makeDocx({ body: p('Full font') }), { fonts: [{ data }], compress: false });
    expect(Buffer.from(result.pdf).includes(data)).toBe(true);
  });

  it('compacts the embedded glyph IDs and remaps composite components', async () => {
    const data = font('DejaVuSans.ttf');
    const original = new SfntFont(data);
    const text = 'Bé';
    const doc = makeDocx({ body: p(text, '', '<w:rFonts w:ascii="Subset test" w:hAnsi="Subset test"/>') });
    const result = await coreConvert(doc, { fonts: [{ family: 'Subset test', data }] });
    const streams = pdfStreams(result.pdf);
    const embedded = streams.find((stream) => stream.startsWith('\x00\x01\x00\x00'));
    expect(embedded).toBeDefined();
    const subset = new SfntFont(Buffer.from(embedded!, 'latin1'));
    const plain = original.glyphId('B'.codePointAt(0)!);
    const composite = original.glyphId('é'.codePointAt(0)!);
    const children = original.compositeComponents(composite);
    expect(children).toHaveLength(2);
    expect(subset.numGlyphs).toBe(5);
    expect(subset.numberOfHMetrics).toBe(subset.numGlyphs);
    expect(subset.table('loca')?.length).toBe((subset.numGlyphs + 1) * 4);
    expect(subset.advance(1)).toBe(original.advance(plain));
    expect(subset.advance(2)).toBe(original.advance(composite));
    expect(subset.compositeComponents(2)).toEqual(children.map(({ offset }, index) => ({ offset, gid: index + 3 })));
    expect(pdfText(result.pdf)).toMatch(/\/CIDToGIDMap \d+ 0 R/);
    const mapping = streams.find((stream) => stream.length === 2 * (composite + 1));
    expect(mapping).toBeDefined();
    expect(mapping!.charCodeAt(plain * 2 + 1)).toBe(1);
    expect(mapping!.charCodeAt(composite * 2 + 1)).toBe(2);
    if (ghostscript) expect(pdfPagesText(result.pdf)[0]).toContain(text);
  });

  it('keeps external links external when a local bookmark has the same name', async () => {
    const body = '<w:p><w:bookmarkStart w:id="0" w:name="Intro"/><w:r><w:t>Local</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
      '<w:p><w:hyperlink r:id="external" w:anchor="Intro"><w:r><w:t>Remote intro</w:t></w:r></w:hyperlink></w:p>';
    const { pdf } = await convertDocx(makeDocx({ body, links: { external: 'https://example.org/manual' } }));
    expect(pdfText(pdf)).toContain('/URI (https://example.org/manual#Intro)');
  });

  it('treats font URLs as directories and shares downloads only within one fetch provider', async () => {
    const doc = makeDocx({ body: p('Test'), styles: arial });
    const urls: string[] = [];
    const provider = (): typeof fetch => async (input) => {
      urls.push(String(input));
      return new Response(new Uint8Array(font('LiberationSans-Regular.ttf')));
    };
    const fontBaseUrl = 'https://font-test.invalid/assets/fonts';
    const first = provider();
    await browserConvert(doc, { fontBaseUrl, fetch: first });
    await browserConvert(doc, { fontBaseUrl: fontBaseUrl + '/', fetch: first });
    await browserConvert(doc, { fontBaseUrl, fetch: provider() });
    expect(urls).toEqual(Array(2).fill(fontBaseUrl + '/LiberationSans-Regular.ttf'));
  });

  it('encodes supplementary Unicode characters in hyperlinks', async () => {
    const doc = makeDocx({ body: '<w:p><w:hyperlink r:id="h"><w:r><w:t>Report</w:t></w:r></w:hyperlink></w:p>', links: { h: 'https://example.org/📄' } });
    expect(pdfText((await convertDocx(doc)).pdf)).toContain('/URI (https://example.org/%F0%9F%93%84)');
  });

  it('preserves distinct Unicode text for characters sharing the missing glyph', async () => {
    const doc = makeDocx({ body: p('漢字日本語') });
    const r = await coreConvert(doc, { fonts: [{ data: font('LiberationSans-Regular.ttf') }] });
    const streams = pdfStreams(r.pdf).join('\n');
    for (const ch of '漢字日本語') expect(streams).toContain('<' + ch.charCodeAt(0).toString(16).toUpperCase() + '>');
    if (ghostscript) expect(pdfPagesText(r.pdf)[0]).toContain('漢字日本語');
  });

  it('fetches a single regular face for Latin and loads CJK on demand by default', async () => {
    const files: string[] = [];
    const fetchFont = async (input: RequestInfo | URL) => {
      const name = new URL(String(input)).pathname.split('/').pop()!;
      files.push(name);
      return new Response(new Uint8Array(font(name)));
    };
    const options = { fontBaseUrl: 'https://font-test.invalid/review/', fetch: fetchFont };
    await browserConvert(makeDocx({ styles: arial, body: p('Hello') + p('日本語', '', '<w:vanish/>') }), options);
    expect(files).toEqual(['LiberationSans-Regular.ttf']);
    const doc = makeDocx({ styles: arial, body: p('漢字日本語 한글') });
    const common = await browserConvert(doc, { ...options, fontPack: 'common' });
    expect(files.some((name) => name.includes('CJK'))).toBe(false);
    expect(common.warnings.join('\n')).toContain('No font has a glyph');
    const cjk = await browserConvert(doc, options);
    // The text has no East Asian font, so it takes Times New Roman's: a serif, like Word's fallback.
    expect(files.filter((name) => name.includes('CJK'))).toEqual(['NotoSerifCJKsc-Regular.ttf']);
    // Looking for the CJK glyphs never fetches the other styles of the Latin fallbacks.
    expect(files.filter((name) => /Bold|Italic/.test(name))).toEqual([]);
    expect(cjk.warnings.filter((w) => w.includes('No font has a glyph'))).toEqual([]);
    if (ghostscript) expect(pdfPagesText(cjk.pdf)[0]).toContain('漢字日本語');
  });

  const baseFonts = (pdf: Uint8Array) => [...new Set([...Buffer.from(pdf).toString('latin1').matchAll(/\/BaseFont \/[A-Z]{6}\+(\S+)/g)].map((m) => m[1]))];

  it('lets a font passed under a family name replace the bundled face of that name', async () => {
    const doc = makeDocx({ body: p('Alias', '', '<w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>') + p('Bold alias', '', '<w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:b/>') });
    const result = await convertDocx(doc, { fonts: [{ family: 'Liberation Sans', data: font('LiberationSerif-Regular.ttf') }] });
    expect(baseFonts(result.pdf)).toEqual(['LiberationSerif']);
  });

  it('keeps loading lazy faces of a fallback family until one has the character', async () => {
    const doc = makeDocx({ body: p('Go ➤', '', '<w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/>') });
    const result = await coreConvert(doc, {
      fonts: [
        { data: font('LiberationSans-Regular.ttf') },
        // Same family and style; only the second face has U+27A4.
        { family: 'DejaVu Sans', load: async () => font('Carlito-Regular.ttf') },
        { family: 'DejaVu Sans', load: async () => font('DejaVuSans.ttf') },
      ],
    });
    expect(result.warnings.filter((w) => w.includes('No font has a glyph'))).toEqual([]);
    expect(baseFonts(result.pdf)).toContain('DejaVuSans');
  });

  it('picks East Asian fallbacks by the font’s region and style, then the text’s language', async () => {
    const registry = new FontRegistry();
    // Stand-in faces registered under the regional Noto CJK names: resolution goes by family name.
    const regional = { 'Noto Serif CJK SC': 'LiberationSerif-Regular.ttf', 'Noto Sans CJK SC': 'LiberationSans-Regular.ttf', 'Noto Serif CJK JP': 'Caladea-Regular.ttf', 'Noto Sans CJK KR': 'Carlito-Regular.ttf' };
    for (const [family, file] of Object.entries(regional)) registry.add({ family, load: async () => font(file) }, 'user');
    registry.add({ family: 'Noto Sans CJK JP', load: async () => font('NotoSansCJKjp-Regular.ttf') }, 'bundled');
    // The font's region wins over the text's language: SimSun is Simplified Chinese whatever the text says.
    const cases = [['SimSun', 'ja-JP'], ['MS Mincho', 'zh-CN'], ['Malgun Gothic', undefined], ['ＭＳ ゴシック', 'ko-KR']] as const;
    await registry.prepare(cases.map(([family, lang]) => ({ family, bold: false, italic: false, codePoints: new Set([0x6f22]), lang })));
    const picked = cases.map(([family, lang]) => registry.resolve(family, false, false, lang).face.families.find((f) => f.startsWith('noto')));
    expect(picked).toEqual(['noto serif cjk sc', 'noto serif cjk jp', 'noto sans cjk kr', 'noto sans cjk jp']);
    // Without a regional font name, the text's language decides and the Latin font's class the style.
    expect(cjkFallbacks('Calibri', 'ko-KR')[0]).toBe('noto sans cjk kr');
    expect(cjkFallbacks('Times New Roman', 'zh-TW')[0]).toBe('noto serif cjk tc');
    expect(cjkFallbacks('PingFang HK', 'zh-CN')[0]).toBe('noto sans cjk hk');
    expect(cjkFallbacks('Songti TC', 'ja-JP')[0]).toBe('noto serif cjk tc');
  });
});

describe('decoded byte budgets', () => {
  it('enforces limits on ZIP parts and on the total of distinct parts read', async () => {
    const doc = makeDocx({ body: p('Budget') });
    const zip = new ZipArchive(doc, nodeCompression.inflateRaw, { maxTotalBytes: new ZipArchive(doc, nodeCompression.inflateRaw).entries.get('word/document.xml')!.size });
    expect(zip.read('word/document.xml')).toBe(zip.read('word/document.xml'));
    expect(() => zip.read('_rels/.rels')).toThrow(/limit/);
    await expect(convertDocx(doc, { archiveLimits: { maxPartBytes: 64 } })).rejects.toThrow(/limit/);
  });

  it('rejects a truncated stream and enforces the actual output limit during decoding', () => {
    const text = new TextEncoder().encode('Quarterly report '.repeat(8));
    const compressed = deflateRawSync(text);
    expect(() => nodeCompression.inflateRaw(compressed.subarray(0, compressed.length - 1), text.length)).toThrow(/truncated/);
    expect(() => nodeCompression.inflateRaw(compressed, text.length - 1)).toThrow(/limit/);
    expect(nodeCompression.inflateRaw(compressed, text.length)).toEqual(text);
  });
});
