import { nodeCompression } from '../src/util/node-compression.js';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { convertDocx } from '../src/node.js';
import { parseDocx } from '../src/docx/parser.js';
import { collectionSize, FontBudget, MAX_FONT_FACES, SfntFont } from '../src/fonts/sfnt.js';
import { ImageBudget } from '../src/image/budget.js';
import { toPdfImage } from '../src/image/image.js';
import { makeDocx, p, zipStore } from './helpers/docx.js';
import { inlineImage, redPng } from './helpers/sample.js';
import { pdfText } from './helpers/pdf.js';

describe('conversion limits', () => {
  it('counts source image pixels once per part and resets the budget for each document', async () => {
    const media = { a: { name: 'a.png', data: redPng() }, b: { name: 'b.png', data: redPng() } };
    const doc = (ids: string[]) => makeDocx({ body: '<w:p>' + ids.map((id) => inlineImage(id, 8, 8)).join('') + '</w:p>', media });
    const options = { imageLimits: { maxImagePixels: 8, maxTotalPixels: 8 } };
    await expect(convertDocx(doc(['a', 'a']), options)).resolves.toHaveProperty('pageCount', 1);
    await expect(convertDocx(doc(['a', 'b']), options)).rejects.toThrow(/pixel limit/);
    await expect(convertDocx(doc(['b']), options)).resolves.toHaveProperty('pageCount', 1);
    await expect(convertDocx(doc(['a']), { imageLimits: { maxImagePixels: 7 } })).rejects.toThrow(/pixel limit/);
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
    await expect(convertDocx(bytes, { archiveLimits, embeddedFonts: false })).resolves.toHaveProperty('pageCount', 1);
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
