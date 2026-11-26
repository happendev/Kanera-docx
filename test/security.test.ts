import { nodeCompression } from '../src/util/node-compression.js';
import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { convertDocx, docxToMarkdown, docxToText } from '../src/node.js';
import { ChartPointBudget, parseChart } from '../src/ooxml/chart.js';
import { DEFAULT_THEME } from '../src/ooxml/theme.js';
import { reserveChartTableCells } from '../src/text/extract.js';
import { formatNumber } from '../src/docx/numbering.js';
import { parseWidth } from '../src/ooxml/props.js';
import { applyWordTintShade, measure } from '../src/ooxml/values.js';
import { ZipArchive } from '../src/util/zip.js';
import { singleAdjustments, singleSubstitutions } from '../src/fonts/gsub.js';
import { parseXml, decodeEntities, resolveAlternateContent } from '../src/util/xml.js';
import { parseDocx } from '../src/docx/parser.js';
import { collectionSize, FontBudget, MAX_FONT_FACES, SfntFont } from '../src/fonts/sfnt.js';
import { ImageBudget } from '../src/image/budget.js';
import { toPdfImage } from '../src/image/image.js';
import { encodeJpeg } from '../src/image/jpeg-encoder.js';
import { makeDocx, p, zipStore } from './helpers/docx.js';
import { inlineImage, redPng, solidPng } from './helpers/sample.js';
import { pdfText } from './helpers/pdf.js';

describe('conversion limits', () => {
  it('validates chart cache lengths and charges sparse positions to the document budget', () => {
    const budget = new ChartPointBudget();
    const parse = (count: number, index: number) => {
      const series = '<c:ser><c:val><c:numLit><c:ptCount val="' + count + '"/><c:pt idx="' + index + '"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>';
      const chart = parseXml('<c:chartSpace><c:chart><c:plotArea><c:barChart>' + series + '</c:barChart></c:plotArea></c:chart></c:chartSpace>');
      return parseChart(chart, DEFAULT_THEME, () => undefined, budget);
    };
    expect(() => parse(-1, 0)).toThrow(/Invalid chart point/);
    expect(() => parse(1, 1.5)).toThrow(/Invalid chart point/);
    expect(() => parse(32769, 0)).toThrow(/Chart point limit/);
    expect(() => parse(1, 32768)).toThrow(/Chart point limit/);
    // Leave room for three slots across the remaining charts.
    for (let i = 0; i < 7; i++) budget.reserve(32768);
    budget.reserve(32765);
    expect(parse(1, 2)?.plots[0].series[0].values).toEqual([null, null, 1]);
    expect(() => parse(1, 0)).toThrow(/Chart point limit/);
  });

  it('counts source image pixels once per part and resets the budget for each document', async () => {
    // Two different pictures: the same picture stored under two part names is embedded and counted once.
    const media = { a: { name: 'a.png', data: redPng() }, b: { name: 'b.png', data: solidPng([0, 0, 255]) } };
    const doc = (ids: string[]) => makeDocx({ body: '<w:p>' + ids.map((id) => inlineImage(id, 8, 8)).join('') + '</w:p>', media });
    const options = { imageLimits: { maxImagePixels: 8, maxTotalPixels: 8 } };
    await expect(convertDocx(doc(['a', 'a']), options)).resolves.toHaveProperty('pageCount', 1);
    await expect(convertDocx(doc(['a', 'b']), options)).rejects.toThrow(/pixel limit/);
    await expect(convertDocx(doc(['b']), options)).resolves.toHaveProperty('pageCount', 1);
    await expect(convertDocx(doc(['a']), { imageLimits: { maxImagePixels: 7 } })).rejects.toThrow(/pixel limit/);
  });

  it('embeds JPEGs beyond the pixel budget unchanged instead of decoding them', async () => {
    const width = 64, height = 48;
    const pixels = new Uint8Array(width * height * 4).map((_, i) => (i % 4 === 3 ? 255 : (i * 37) % 251));
    const jpeg = encodeJpeg(width, height, pixels, 100, { channels: 4 });
    const budget = new ImageBudget({ maxImagePixels: width * height - 1 });
    const image = await toPdfImage(jpeg, nodeCompression, budget);
    expect(image).toMatchObject({ width, height, filter: 'DCTDecode' });
    expect(Buffer.from(image!.data).equals(Buffer.from(jpeg))).toBe(true);
    // Recompression still spends the budget when the image fits.
    const fitting = new ImageBudget({ maxImagePixels: width * height, maxTotalPixels: width * height });
    await toPdfImage(jpeg, nodeCompression, fitting);
    expect(fitting.tryReserve(1, 1)).toBe(false);
  });

  it('caps automatic JPEG selection at four million pixels and honors explicit quality above it', async () => {
    // Change only the SOF dimensions. A declining budget lets us observe selection without
    // allocating a four-megapixel image or passing the synthetic frame to the decoder.
    const jpeg = Buffer.from(encodeJpeg(8, 8, new Uint8Array(256), 100, { channels: 4 }));
    const frame = jpeg.indexOf(Buffer.from([0xff, 0xc0]));
    expect(frame).toBeGreaterThan(0);
    jpeg.writeUInt16BE(2000, frame + 5);
    for (const [width, quality, selected] of [
      [2000, undefined, true], [2001, undefined, false], [2001, 80, true], [2001, false, false],
    ] as const) {
      jpeg.writeUInt16BE(width, frame + 7);
      const budget = new ImageBudget({ maxTotalPixels: 0 });
      const reserve = vi.spyOn(budget, 'tryReserve');
      const image = await toPdfImage(jpeg, nodeCompression, budget, quality);
      expect(reserve).toHaveBeenCalledTimes(selected ? 1 : 0);
      expect(image?.data).toBe(jpeg);
    }
  });

  it('charges JPEG pixels only when 8-bit table zero meets the automatic cutoff', async () => {
    const jpeg = Buffer.from(encodeJpeg(8, 8, new Uint8Array(8 * 8 * 4).fill(85), 100, { channels: 4 }));
    const start = jpeg.indexOf(Buffer.from([0xff, 0xdb]));
    const end = start + 2 + jpeg.readUInt16BE(start + 2);
    const table = (sum: number) => [0, ...Array.from({ length: 64 }, (_, i) => Math.floor(sum / 64) + Number(i < sum % 64))];
    const other = [1, ...Array(64).fill(1)];
    const segment = (...tables: number[][]) => {
      const payload = Buffer.from(tables.flat());
      const header = Buffer.from([0xff, 0xdb, 0, 0]);
      header.writeUInt16BE(payload.length + 2, 2);
      return Buffer.concat([header, payload]);
    };
    for (const [replacement, selected] of [
      [segment(other, table(750)), true],
      [segment(table(751), other), false],
      [segment(other), false],
      [segment(table(750), [0x10, ...Array(128).fill(1)]), false],
      [Buffer.concat([segment(table(750)), segment(table(751), other)]), false],
    ] as const) {
      const bytes = Buffer.concat([jpeg.subarray(0, start), replacement, jpeg.subarray(end)]);
      const budget = new ImageBudget({ maxTotalPixels: 64 });
      const image = await toPdfImage(bytes, nodeCompression, budget);
      expect(budget.tryReserve(8, 8)).toBe(!selected);
      if (!selected) expect(image?.data).toBe(bytes);
    }
  });

  it('applies the raster budget to both GIF and BMP dimensions', async () => {
    const gif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAkQBADs=', 'base64');
    const bmp = Buffer.alloc(58);
    bmp.write('BM');
    bmp.writeUInt32LE(54, 10);
    bmp.writeUInt32LE(40, 14);
    bmp.writeInt32LE(1, 18);
    bmp.writeInt32LE(1, 22);
    bmp.writeUInt16LE(1, 26);
    bmp.writeUInt16LE(24, 28);
    for (const data of [gif, bmp]) {
      expect(await toPdfImage(data, nodeCompression)).toMatchObject({ width: 1, height: 1 });
      await expect(toPdfImage(data, nodeCompression, new ImageBudget({ maxImagePixels: 0 }))).rejects.toThrow(/pixel limit/);
    }
    const frame = Buffer.from(gif);
    frame.writeUInt16LE(2, frame.indexOf(0x2c) + 5);
    await expect(toPdfImage(frame, nodeCompression)).rejects.toThrow(/frame dimensions/);
    await expect(toPdfImage(bmp.subarray(0, -1), nodeCompression)).rejects.toThrow(/truncated/);
  });

  it('counts padded TIFF tiles before allocating their decoded buffers', async () => {
    // One grayscale pixel in a two-pixel tile.
    const entries = [[256, 1], [257, 1], [258, 8], [322, 2], [323, 1], [324, 98], [325, 2]];
    const tiff = Buffer.alloc(100);
    tiff.write('II');
    tiff.writeUInt16LE(42, 2);
    tiff.writeUInt32LE(8, 4);
    tiff.writeUInt16LE(entries.length, 8);
    entries.forEach(([tag, value], i) => {
      const offset = 10 + i * 12;
      tiff.writeUInt16LE(tag, offset);
      tiff.writeUInt16LE(4, offset + 2);
      tiff.writeUInt32LE(1, offset + 4);
      tiff.writeUInt32LE(value, offset + 8);
    });
    expect(await toPdfImage(tiff, nodeCompression)).toMatchObject({ width: 1, height: 1 });
    await expect(toPdfImage(tiff, nodeCompression, new ImageBudget({ maxImagePixels: 1 }))).rejects.toThrow(/pixel limit/);
  });

  it('keeps font hints and skips the binary part when embedded fonts are disabled', async () => {
    const xml = (s: string) => new TextEncoder().encode(s);
    const bytes = zipStore([
      ['word/document.xml', xml('<w:document><w:body>' + p('Hello') + '</w:body></w:document>')],
      ['word/_rels/document.xml.rels', xml('<Relationships><Relationship Id="ft" Type="/fontTable" Target="fonts.xml"/></Relationships>')],
      ['word/fonts.xml', xml('<w:fonts><w:font w:name="Example"><w:altName w:val="Arial"/><w:embedRegular r:id="f"/></w:font></w:fonts>')],
      ['word/_rels/fonts.xml.rels', xml('<Relationships><Relationship Id="f" Type="/font" Target="face.odttf"/></Relationships>')],
      ['word/face.odttf', new Uint8Array(1025)],
    ]);
    const archiveLimits = { maxPartBytes: 1024 };
    await expect(convertDocx(bytes, { archiveLimits })).rejects.toThrow(/face.odttf/);
    await expect(convertDocx(bytes, { archiveLimits, useEmbeddedFonts: false })).resolves.toHaveProperty('pageCount', 1);
    expect(parseDocx(bytes, nodeCompression.inflateRaw, archiveLimits, false).fontHints.get('Example')?.altName).toBe('Arial');
  });

  it('checks collection offsets and reserves faces and mappings before expansion', () => {
    const collection = Buffer.alloc(16);
    collection.write('ttcf');
    collection.writeUInt32BE(2, 8);
    expect(() => collectionSize(collection)).toThrow(/Truncated/);
    const budget = new FontBudget();
    budget.reserveFaces(MAX_FONT_FACES);
    expect(() => budget.reserveFaces(1)).toThrow(/face limit/);
    const data = readFileSync(new URL('../fonts/Carlito-Regular.ttf', import.meta.url));
    budget.reserveMappings(1_048_576);
    expect(() => new SfntFont(data, 0, budget)).toThrow(/mapping limit/);
    const cmap = new SfntFont(data).table('cmap')!;
    data.writeUInt16BE(Math.floor(cmap.length / 8) + 1, cmap.offset + 2);
    expect(() => new SfntFont(data)).toThrow(/Truncated/);
  });

  it('handles namespace shadowing and enforces the XML depth boundary', () => {
    expect(decodeEntities('&lt;&constructor;&toString;')).toBe('<&constructor;&toString;');
    const alternate = parseXml('<mc:AlternateContent xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<mc:Choice xmlns:wps="unknown" Requires="wps"><first/></mc:Choice><mc:Choice Requires="wps"><second/></mc:Choice></mc:AlternateContent>');
    expect(resolveAlternateContent(alternate).map((el) => el.name)).toEqual(['second']);
    const tree = (n: number) => '<e>'.repeat(n) + '</e>'.repeat(n);
    expect(parseXml(tree(128)).name).toBe('e');
    expect(() => parseXml(tree(129))).toThrow(/nesting limit/);
  });

  it('counts rendered chart cells across text tables', () => {
    for (const [markdown, cells] of [[false, 9], [true, 12]] as const) {
      const remaining = reserveChartTableCells(2 * cells, 2, 2, markdown);
      expect(remaining).toBe(cells);
      expect(reserveChartTableCells(remaining, 2, 2, markdown)).toBe(0);
      expect(() => reserveChartTableCells(remaining - 1, 2, 2, markdown)).toThrow(/table cell limit/);
    }
  });

  it('charges every Word text import before decoding it', async () => {
    const doc = (copies: number) => makeDocx({
      body: '<w:altChunk r:id="c"/>'.repeat(copies),
      imports: { c: { name: 'note.txt', contentType: 'text/plain', text: 'A'.repeat(1024) } },
    });
    const options = { archiveLimits: { maxTotalBytes: 8192 } };
    await expect(docxToText(doc(8), options)).resolves.toHaveLength(8 * 1025);
    await expect(docxToText(doc(9), options)).rejects.toThrow(/Word import byte limit/);
  });

  it('reads quoted HTML tags and handles unfinished altChunk markup', async () => {
    const text = '<head>open <style>hidden</style>text<p>a<b title="1 > 0">b</b></td>c</p>one < two <a';
    const doc = makeDocx({ body: '<w:altChunk r:id="c"/>', imports: { c: { name: 'frag.html', contentType: 'text/html', text } } });
    await expect(docxToText(doc)).resolves.toBe('open text\nab c\none < two <a\n');
  });

  it('validates XML structure and counts namespace declarations in the per-element attribute limit', async () => {
    const attrs = Array.from({ length: 255 }, (_, i) => ' a' + i + '="v"').join('');
    const doc = (xml: string) => zipStore([['word/document.xml', new TextEncoder().encode(xml)]]);
    const start = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' + attrs;
    await expect(docxToText(doc(start + '><w:body/></w:document>'))).resolves.toBe('');
    await expect(docxToText(doc(start + ' b="v"><w:body/></w:document>'))).rejects.toThrow(/XML attribute limit/);
    for (const xml of ['<other><w:body/></other>', '<w:document><w:body/>', '<w:document><w:body/></other>', '<w:document><w:body/></w:document><extra/>']) {
      await expect(docxToText(doc(xml))).rejects.toThrow(/invalid main document part/);
    }
  });

  it('bounds Word field nesting across runs and accepts the boundary', async () => {
    const fld = (type: string) => '<w:r><w:fldChar w:fldCharType="' + type + '"/></w:r>';
    const doc = (n: number) => makeDocx({ body: '<w:p>' + fld('begin').repeat(n) + fld('end').repeat(n) + '</w:p>' + p('Done') });
    await expect(docxToText(doc(128))).resolves.toBe('Done\n');
    await expect(docxToText(doc(129))).rejects.toThrow(/field nesting limit/);
  });

  it('treats out-of-range outline and list levels as body text', async () => {
    for (const level of [-2, 1.5, 9]) {
      const numbering = '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="' + level + '"><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
      const props = '<w:outlineLvl w:val="' + level + '"/><w:numPr><w:ilvl w:val="' + level + '"/><w:numId w:val="1"/></w:numPr>';
      await expect(docxToMarkdown(makeDocx({ body: p('Item', props), numbering }))).resolves.toBe('Item\n');
    }
  });

  it('keeps measurement, colour and numbering fallbacks finite and bounded', () => {
    expect(measure('.5in')).toBe(36);
    for (const value of ['.', '1.2.3', '9'.repeat(309) + 'in']) expect(measure(value)).toBeUndefined();
    expect(parseWidth(parseXml('<w:tcW w:type="pct" w:w="%"/>'))?.value).toBe(0);
    expect(applyWordTintShade('ABCDEF', 'GG', '2z')).toBe('ABCDEF');
    expect(formatNumber(5, 'chicago')).toBe('**');
    expect(formatNumber(0, 'chicago')).toBe('0');
    for (const format of ['upperRoman', 'lowerLetter', 'chicago']) expect(formatNumber(1_000_000, format)).toBe('1000000');
  });

  it('reads legacy ZIP names and a comment containing a signature', () => {
    const zip = Buffer.from(zipStore([['dir\\x.txt', Uint8Array.of(7)]]));
    const central = zip.readUInt32LE(zip.length - 6);
    zip[30 + 4] = zip[central + 46 + 4] = 0x82; // é in CP437
    const comment = Buffer.alloc(32);
    comment.writeUInt32LE(0x06054b50);
    zip.writeUInt16LE(comment.length, zip.length - 2);
    const archive = new ZipArchive(Buffer.concat([zip, comment]), nodeCompression.inflateRaw);
    expect(Array.from(archive.read('dir/é.txt')!)).toEqual([7]);
    // Bytes after the end record and its comment, such as a newline added in transit, are ignored.
    const padded = new ZipArchive(Buffer.concat([zip, comment, Buffer.from('\n')]), nodeCompression.inflateRaw);
    expect(Array.from(padded.read('dir/é.txt')!)).toEqual([7]);
  });

  it('applies the feature budget to substitution and adjustment tables', () => {
    const font = new SfntFont(readFileSync(new URL('../fonts/NotoSansCJKjp-Regular.ttf', import.meta.url)));
    const budget = new FontBudget();
    budget.reserveFeatureWork(1_048_576);
    expect(() => singleSubstitutions(font.tableBytes('GSUB')!, 'vert', budget)).toThrow(/feature limit/);
    expect(() => singleAdjustments(font.tableBytes('GPOS')!, 'palt', budget)).toThrow(/feature limit/);
  });

  it('checks glyph locations against the declared table length', () => {
    const font = new SfntFont(readFileSync(new URL('../fonts/Carlito-Regular.ttf', import.meta.url)));
    expect(font.glyphRange(0)).toBeTruthy();
    font.table('loca')!.length = 2;
    expect(() => font.glyphRange(0)).toThrow(/Truncated font table/);
  });

  it('checks the complete table row and both forms of section columns', () => {
    const table = (span: number) => makeDocx({ body: '<w:tbl><w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>' +
      '<w:tc><w:tcPr><w:gridSpan w:val="' + span + '"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' });
    expect(() => parseDocx(table(254), nodeCompression.inflateRaw)).not.toThrow();
    expect(() => parseDocx(table(255), nodeCompression.inflateRaw)).toThrow(/Table column limit/);
    for (const cols of ['<w:cols w:num="65"/>', '<w:cols w:equalWidth="0">' + '<w:col/>'.repeat(65) + '</w:cols>']) {
      expect(() => parseDocx(makeDocx({ body: p('Body'), sectPr: '<w:sectPr>' + cols + '</w:sectPr>' }), nodeCompression.inflateRaw)).toThrow(/Section column limit/);
    }
  });

  it('writes only permitted link schemes and resolves internal bookmarks', async () => {
    const links = { web: 'HtTpS://example.org', http: 'http://example.org', email: 'mailto:user@example.org', phone: 'tel:123', relative: 'local.docx', bookmark: '#here' };
    const doc = makeDocx({ links, body: '<w:p><w:bookmarkStart w:id="1" w:name="here"/>' +
      Object.keys(links).map((id) => '<w:hyperlink r:id="' + id + '"><w:r><w:t>' + id + '</w:t></w:r></w:hyperlink>').join('') +
      '<w:hyperlink w:anchor="here"><w:r><w:t>Bookmark</w:t></w:r></w:hyperlink></w:p>' });
    const pdf = pdfText((await convertDocx(doc)).pdf);
    for (const url of [links.web, links.http, links.email]) expect(pdf).toContain('/URI (' + url + ')');
    expect(pdf.match(/\/Dest \[/g)).toHaveLength(2);
    expect(pdf).not.toContain('/URI (tel:');
    expect(pdf).not.toContain('/URI (local.docx)');
    expect(pdfText((await convertDocx(doc, { additionalLinkSchemes: ['TEL'] })).pdf)).toContain('/URI (tel:123)');
  });
});
