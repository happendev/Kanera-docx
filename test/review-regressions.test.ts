import { nodeCompression } from '../src/util/node-compression.js';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { convertDocx } from '../src/node.js';
import { convertDocx as browserConvert } from '../src/browser.js';
import { convertDocx as coreConvert } from '../src/index.js';
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

  it('draws the prohibition symbol with two cutouts and a diagonal bar', async () => {
    const body = '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1270000" cy="635000"/><a:graphic><a:graphicData>' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr>' +
      '<a:prstGeom prst="noSmoking"/><a:solidFill><a:srgbClr val="0077AA"/></a:solidFill>' +
      '</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const result = await layout(makeDocx({ body }));
    const path = result.pages[0].body.find((op) => op.kind === 'path' && op.fill === '0077AA');
    expect(path?.kind === 'path' && path.segs.filter((s) => s.op === 'Z')).toHaveLength(3);
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
    expect(files.filter((name) => name.includes('CJK'))).toEqual(['NotoSansCJKjp-Regular.ttf']);
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
