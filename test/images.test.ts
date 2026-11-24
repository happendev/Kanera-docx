import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { convertDocxToImages, docxToImages } from '../src/node.js';
import { makeDocx, p } from './helpers/docx.js';
import { inlineImage, redPng } from './helpers/sample.js';

/** Decode an 8-bit RGB PNG (as the renderer writes them) to its pixels. */
function decodePng(png: Uint8Array): { width: number; height: number; rgb: Uint8Array } {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  expect(png[24]).toBe(8);
  expect(png[25]).toBe(2);
  const idat: Uint8Array[] = [];
  for (let at = 8; at < png.length; ) {
    const length = view.getUint32(at);
    const type = String.fromCharCode(...png.subarray(at + 4, at + 8));
    if (type === 'IDAT') idat.push(png.subarray(at + 8, at + 8 + length));
    at += 12 + length;
  }
  const raw = inflateSync(Buffer.concat(idat));
  const stride = width * 3;
  const rgb = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = raw[y * (stride + 1)];
    // Check once per row: an expect() per byte costs seconds on a full page.
    expect(filter).toBeLessThanOrEqual(2);
    for (let x = 0; x < stride; x++) {
      const v = raw[y * (stride + 1) + 1 + x];
      const left = x >= 3 ? rgb[y * stride + x - 3] : 0;
      const up = y ? rgb[(y - 1) * stride + x] : 0;
      rgb[y * stride + x] = (v + (filter === 0 ? 0 : filter === 1 ? left : up)) & 255;
    }
  }
  return { width, height, rgb };
}

/** Pixels in the region whose colour matches `test`. */
function count(img: { width: number; rgb: Uint8Array }, test: (r: number, g: number, b: number) => boolean, x0 = 0, y0 = 0, x1 = img.width, y1 = img.rgb.length / img.width / 3): number {
  let n = 0;
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const o = (y * img.width + x) * 3;
      if (test(img.rgb[o], img.rgb[o + 1], img.rgb[o + 2])) n++;
    }
  }
  return n;
}

const dark = (r: number, g: number, b: number) => r + g + b < 200 * 3;
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

describe('page images', () => {
  it('renders one PNG per page at the requested resolution', async () => {
    const body = Array.from({ length: 3 }, (_, i) => p(`Page ${i + 1}`)).join(pageBreak);
    const result = await convertDocxToImages(makeDocx({ body }), { dpi: 36 });
    expect(result.pageCount).toBe(3);
    expect(result.images.map((i) => i.page)).toEqual([1, 2, 3]);
    for (const image of result.images) {
      expect([...image.data.subarray(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
      // A4 (595.3 × 841.9 pt) at 36 dpi.
      expect([image.width, image.height]).toEqual([298, 421]);
      const png = decodePng(image.data);
      expect([png.width, png.height]).toEqual([298, 421]);
    }
  });

  it('draws text where the PDF does, and leaves the rest of the page white', async () => {
    const [image] = await docxToImages(makeDocx({ body: p('Hello world') }), { dpi: 72 });
    const png = decodePng(image.data);
    // The text starts at the 1" margin, on the first line.
    const ink = count(png, dark, 72, 72, 160, 90);
    expect(ink).toBeGreaterThan(40);
    expect(count(png, dark, 0, 0, png.width, 70)).toBe(0);
    expect(count(png, dark, 0, 100, png.width, png.height)).toBe(0);
  });

  it('renders only the chosen pages, and rejects pages that do not exist', async () => {
    const docx = makeDocx({ body: [p('One'), p('Two'), p('Three')].join(pageBreak) });
    const images = await docxToImages(docx, { dpi: 20, pages: [3, 1] });
    expect(images.map((i) => i.page)).toEqual([3, 1]);
    await expect(docxToImages(docx, { pages: [4] })).rejects.toThrow(RangeError);
    await expect(docxToImages(docx, { dpi: 0 })).rejects.toThrow(RangeError);
    await expect(docxToImages(docx, { format: 'jpeg', quality: 101 })).rejects.toThrow(RangeError);
  });

  it('encodes JPEG pages', async () => {
    const [image] = await docxToImages(makeDocx({ body: p('Hello') }), { format: 'jpeg', dpi: 50, quality: 80 });
    expect([image.data[0], image.data[1], image.data[2]]).toEqual([0xff, 0xd8, 0xff]);
    expect([image.data.at(-2), image.data.at(-1)]).toEqual([0xff, 0xd9]);
    expect([image.width, image.height]).toEqual([414, 585]);
  });

  it('draws pictures, shading and borders', async () => {
    const shaded = p('Shaded', '<w:shd w:val="clear" w:color="auto" w:fill="0000FF"/>');
    const bordered = p('Bordered', '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="00FF00"/></w:pBdr>');
    const picture = `<w:p>${inlineImage('rIdImg', 80, 40)}</w:p>`;
    const docx = makeDocx({ body: shaded + bordered + picture, media: { rIdImg: { name: 'red.png', data: redPng() } } });
    const [image] = await docxToImages(docx, { dpi: 72 });
    const png = decodePng(image.data);
    expect(count(png, (r, g, b) => b > 200 && r < 60 && g < 60)).toBeGreaterThan(400);
    expect(count(png, (r, g, b) => g > 200 && r < 60 && b < 60)).toBeGreaterThan(200);
    // The 80 × 40 pt picture, upscaled from 4 × 2 pixels.
    const red = count(png, (r, g, b) => r > 200 && g < 60 && b < 60);
    expect(red).toBeGreaterThan(80 * 40 * 0.8);
    expect(red).toBeLessThan(80 * 40 * 1.2);
  });

  it('draws CFF and CJK fonts', async () => {
    const run = (font: string, text: string) => `<w:p><w:r><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
    const [image] = await docxToImages(makeDocx({ body: run('TeX Gyre Adventor', 'Adventor') + p('日本語のテキスト') }), { dpi: 72 });
    const png = decodePng(image.data);
    expect(count(png, dark, 72, 70, 200, 88)).toBeGreaterThan(40);
    expect(count(png, dark, 72, 88, 200, 110)).toBeGreaterThan(40);
  });
});
