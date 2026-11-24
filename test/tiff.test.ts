import { inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { ImageBudget } from '../src/image/budget.js';
import { decodeFax } from '../src/image/ccitt.js';
import { toPdfImage } from '../src/image/image.js';
import { decodeJpegPixels } from '../src/image/jpeg-decoder.js';
import { encodeJpeg } from '../src/image/jpeg-encoder.js';
import { decodePixels } from '../src/image/pixels.js';
import { decodeTiff, tiffJpeg } from '../src/image/tiff.js';
import { nodeCompression } from '../src/util/node-compression.js';

/** A little-endian TIFF holding `chunks` as its strips (or tiles), with the other tags in `fields`. */
function tiff(fields: Record<number, number | number[] | Uint8Array>, chunks: Uint8Array[], tiled = false): Uint8Array {
  const tags = new Map<number, number[] | Uint8Array>(Object.entries(fields).map(([k, v]) => [Number(k), typeof v === 'number' ? [v] : v]));
  let at = 8;
  tags.set(tiled ? 324 : 273, chunks.map((c) => ((at += c.length), at - c.length)));
  tags.set(tiled ? 325 : 279, chunks.map((c) => c.length));
  const extra: number[] = [];
  const entries = [...tags.keys()].sort((a, b) => a - b).map((tag) => {
    const v = tags.get(tag)!;
    const bytes = v instanceof Uint8Array ? v : new Uint8Array(new Uint32Array(v).buffer);
    const entry = { tag, type: v instanceof Uint8Array ? 7 : 4, count: v.length, bytes, offset: 0 };
    if (bytes.length > 4) {
      entry.offset = at + extra.length;
      extra.push(...bytes);
    }
    return entry;
  });
  if (extra.length & 1) extra.push(0);
  const ifd = at + extra.length;
  const out = new Uint8Array(ifd + 2 + entries.length * 12 + 4);
  const dv = new DataView(out.buffer);
  out.set([0x49, 0x49, 42, 0]);
  dv.setUint32(4, ifd, true);
  let o = 8;
  for (const c of chunks) out.set(c, (o += c.length) - c.length);
  out.set(extra, o);
  dv.setUint16(ifd, entries.length, true);
  entries.forEach((e, i) => {
    const p = ifd + 2 + i * 12;
    dv.setUint16(p, e.tag, true);
    dv.setUint16(p + 2, e.type, true);
    dv.setUint32(p + 4, e.count, true);
    if (e.bytes.length > 4) dv.setUint32(p + 8, e.offset, true);
    else out.set(e.bytes, p + 8);
  });
  return out;
}

/** Bits written as a string of 0s and 1s, padded with zeros to whole bytes. */
function bits(s: string): Uint8Array {
  const out = new Uint8Array(Math.ceil(s.length / 8));
  for (let i = 0; i < s.length; i++) if (s[i] === '1') out[i >> 3] |= 0x80 >> (i & 7);
  return out;
}

// Two rows of 16 pixels: 4 white, 8 black, 4 white.
const ROWS = [0x0f, 0xf0, 0x0f, 0xf0];
const EOL = '000000000001';
// One-dimensional row: white 4, black 8, white 4.
const ROW_1D = '1011' + '000101' + '1011';
// T.6: the first row horizontal then V0 against a white row; the second V0 three times; then EOFB.
const G4 = bits('001' + '1011' + '000101' + '1' + '111' + EOL + EOL);
const MH = bits(ROW_1D.padEnd(16, '0') + ROW_1D.padEnd(16, '0'));
const reverse = (d: Uint8Array) => d.map((b) => parseInt(b.toString(2).padStart(8, '0').split('').reverse().join(''), 2));

function rgbaImage(width: number, height: number): { width: number; height: number; data: Uint8Array } {
  const data = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) data.set([x * 16, y * 8, 128 + ((x * y) & 63), 255], (y * width + x) * 4);
  }
  return { width, height, data };
}

const jpeg = (width: number, height: number, rgba: Uint8Array) => encodeJpeg(width, height, rgba, 90, { channels: 4 });
const PICTURE = rgbaImage(16, 32);

/** JPEG files of `rows`-row bands of a 16 × 32 picture, coded in 16 × 16 MCUs (4:2:0). */
function jpegBands(rows: number): Uint8Array[] {
  const bands: Uint8Array[] = [];
  for (let y = 0; y < 32; y += rows) bands.push(jpeg(16, Math.min(rows, 32 - y), PICTURE.data.subarray(y * 64, (y + rows) * 64)));
  return bands;
}

/** A JPEG's marker segments of the given kinds, and the JPEG without them. */
function splitSegments(jpeg: Uint8Array, markers: number[]): { taken: Uint8Array; rest: Uint8Array } {
  const taken: number[] = [];
  const rest: number[] = [0xff, 0xd8];
  let p = 2;
  while (jpeg[p + 1] !== 0xda) {
    const end = p + 2 + ((jpeg[p + 2] << 8) | jpeg[p + 3]);
    (markers.includes(jpeg[p + 1]) ? taken : rest).push(...jpeg.subarray(p, end));
    p = end;
  }
  rest.push(...jpeg.subarray(p));
  return { taken: new Uint8Array(taken), rest: new Uint8Array(rest) };
}

const JPEG_FIELDS = { 256: 16, 257: 32, 258: [8, 8, 8], 259: 7, 262: 6, 277: 3 };

describe('CCITT fax', () => {
  it('decodes T.6, T.4 and Modified Huffman rows', () => {
    expect(Array.from(decodeFax(G4, 16, 2, { coding: 't6' }))).toEqual(ROWS);
    expect(Array.from(decodeFax(MH, 16, 2, { coding: 'mh' }))).toEqual(ROWS);
    // T.4 with an EOL before each row, the first padded with fill bits.
    expect(Array.from(decodeFax(bits('0000' + EOL + ROW_1D + EOL + ROW_1D), 16, 2, { coding: 't4' }))).toEqual(ROWS);
    // Two-dimensional T.4: a tag bit after each EOL, 1 for a one-dimensional row, 0 for a two-dimensional one.
    expect(Array.from(decodeFax(bits(EOL + '1' + ROW_1D + EOL + '0' + '111'), 16, 2, { coding: 't4', twoDimensional: true }))).toEqual(ROWS);
    expect(() => decodeFax(bits('0000001' + '0'.repeat(20)), 16, 1, { coding: 't6' })).toThrow(/CCITT/);
  });

  it('embeds a fax TIFF of one strip as CCITT data, in its bit order and colours', async () => {
    // BlackIsZero, with the bits of each byte stored in reverse order.
    const g4 = tiff({ 256: 16, 257: 2, 259: 4, 262: 1, 266: 2, 278: 2 }, [reverse(G4)]);
    const img = await toPdfImage(g4, nodeCompression);
    expect(img).toMatchObject({ width: 16, height: 2, colorSpace: '/DeviceGray', bitsPerComponent: 1, filter: 'CCITTFaxDecode' });
    expect(img!.decodeParms).toBe('<< /K -1 /Columns 16 /Rows 2 /BlackIs1 true >>');
    expect(img!.data).toEqual(G4);

    const mh = await toPdfImage(tiff({ 256: 16, 257: 2, 259: 2, 262: 0 }, [MH]), nodeCompression);
    expect(mh!.decodeParms).toBe('<< /K 0 /Columns 16 /Rows 2 /EncodedByteAlign true >>');
    const t4 = await toPdfImage(tiff({ 256: 16, 257: 2, 259: 3, 262: 0, 292: 1 }, [bits(EOL + '1' + ROW_1D + EOL + '0' + '111')]), nodeCompression);
    expect(t4!.decodeParms).toBe('<< /K 1 /Columns 16 /Rows 2 /EndOfLine true >>');
  });

  it('decodes fax TIFFs that PDF readers could not read as one stream', async () => {
    // Two strips of one row, and a T.4 row without its EOL.
    const strips = tiff({ 256: 16, 257: 2, 259: 2, 262: 0, 278: 1 }, [MH.subarray(0, 2), MH.subarray(2)]);
    const noEol = tiff({ 256: 16, 257: 2, 259: 3, 262: 1 }, [bits(ROW_1D + EOL + ROW_1D)]);
    const row = [0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0];
    for (const [file, photometric] of [[strips, 0], [noEol, 1]] as const) {
      const img = await toPdfImage(file, nodeCompression);
      expect(img).toMatchObject({ filter: 'FlateDecode', bitsPerComponent: 8 });
      expect(img!.colorSpace).toBe(photometric === 0 ? '[/Indexed /DeviceRGB 1 <ffffff000000>]' : '[/Indexed /DeviceRGB 1 <000000ffffff>]');
      expect(Array.from(inflateSync(img!.data))).toEqual([...row, ...row]);
      const px = await decodePixels(file, nodeCompression, new ImageBudget());
      const grey = Array.from(px!.data.filter((_, i) => i % 4 === 0));
      // A bit of 1 is black in WhiteIsZero, white in BlackIsZero.
      expect(grey).toEqual([...row, ...row].map((v) => ((v === 1) === (photometric === 0) ? 0 : 255)));
    }
  });
});

describe('JPEG-compressed TIFF', () => {
  it('joins JPEG strips into one JPEG, restart markers between them, without recompressing', async () => {
    // The joined strips decode as the whole picture coded at once: the same blocks, and chroma
    // upsampled across the seams.
    const bands = jpegBands(16);
    const expected = decodeJpegPixels(jpeg(16, 32, PICTURE.data)).data;
    // Tables in each strip, or shared in JPEGTables.
    const shared = bands.map((b) => splitSegments(b, [0xdb, 0xc4]));
    const tables = new Uint8Array([0xff, 0xd8, ...shared[0].taken, 0xff, 0xd9]);
    for (const file of [tiff({ ...JPEG_FIELDS, 278: 16 }, bands), tiff({ ...JPEG_FIELDS, 278: 16, 347: tables }, shared.map((s) => s.rest))]) {
      const jpeg = tiffJpeg(file)!;
      const pixels = decodeJpegPixels(jpeg);
      expect([pixels.width, pixels.height]).toEqual([16, 32]);
      expect(pixels.data).toEqual(expected);
      expect(jpeg.findIndex((b, i) => b === 0xff && jpeg[i + 1] === 0xd0)).toBeGreaterThan(0);
      const img = await toPdfImage(file, nodeCompression, undefined, false);
      expect(img).toMatchObject({ width: 16, height: 32, filter: 'DCTDecode', colorSpace: '/DeviceRGB' });
      expect(img!.data).toEqual(jpeg);
      expect((await decodePixels(file, nodeCompression, new ImageBudget()))!.data).toEqual(expected);
    }
  });

  it('marks how the JPEG components are coded, and passes old-style JPEG files through', async () => {
    const [whole] = jpegBands(32);
    const adobe = (d: Uint8Array) => {
      const at = d.findIndex((b, i) => b === 0xff && d[i + 1] === 0xee);
      return at < 0 ? null : d[at + 15];
    };
    expect(adobe(tiffJpeg(tiff(JPEG_FIELDS, [whole]))!)).toBe(1);
    expect(adobe(tiffJpeg(tiff({ ...JPEG_FIELDS, 262: 2 }, [whole]))!)).toBe(0);
    expect(tiffJpeg(tiff({ ...JPEG_FIELDS, 262: 0 }, [whole]))).toBeNull();
    // Recompression decodes RGB-coded components as RGB.
    const rgb = tiff({ ...JPEG_FIELDS, 262: 2 }, [whole]);
    const smaller = await toPdfImage(rgb, nodeCompression, undefined, 50);
    expect(smaller!.data).not.toEqual(tiffJpeg(rgb));
    const before = decodeJpegPixels(tiffJpeg(rgb)!).data;
    const after = decodeJpegPixels(smaller!.data).data;
    expect(Math.max(...before.map((v, i) => Math.abs(v - after[i])))).toBeLessThan(48);
    const old = tiff({ 256: 16, 257: 32, 259: 6, 262: 6, 277: 3, 513: 8, 514: whole.length }, [whole]);
    expect(tiffJpeg(old)).toEqual(whole);
    expect((await toPdfImage(old, nodeCompression, undefined, false))!.data).toEqual(whole);
  });

  it('decodes JPEG tiles and strips that cannot join', () => {
    // Strips of 8 rows split the 16-row MCUs; tiles are side by side.
    const bands = jpegBands(8);
    const expected = new Uint8Array(bands.flatMap((b) => Array.from(decodeJpegPixels(b).data)));
    const strips = tiff({ ...JPEG_FIELDS, 278: 8 }, bands);
    expect(tiffJpeg(strips)).toBeNull();
    expect((decodeTiff(strips, nodeCompression.inflateZlib) as { data: Uint8Array }).data).toEqual(expected);

    const tile = (x: number) => {
      const data = new Uint8Array(8 * 32 * 4);
      for (let y = 0; y < 32; y++) data.set(PICTURE.data.subarray((y * 16 + x) * 4, (y * 16 + x + 8) * 4), y * 32);
      return jpeg(8, 32, data);
    };
    const tiles = [tile(0), tile(8)];
    const tiled = tiff({ ...JPEG_FIELDS, 322: 8, 323: 32 }, tiles, true);
    expect(tiffJpeg(tiled)).toBeNull();
    const out = (decodeTiff(tiled, nodeCompression.inflateZlib) as { data: Uint8Array }).data;
    const [left, right] = tiles.map((t) => decodeJpegPixels(t).data);
    for (let y = 0; y < 32; y++) {
      expect(out.subarray(y * 64, y * 64 + 32)).toEqual(left.subarray(y * 32, y * 32 + 32));
      expect(out.subarray(y * 64 + 32, y * 64 + 64)).toEqual(right.subarray(y * 32, y * 32 + 32));
    }
  });
});
