import { deflateRawSync, deflateSync, inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { imageSize, toPdfImage } from '../src/image/image.js';
import { decodeTiff } from '../src/image/tiff.js';
import { nodeCompression } from '../src/util/node-compression.js';
import { pakoCompression } from '../src/util/pako.js';
import { child, children, parseXml, resolveAlternateContent, textContent } from '../src/util/xml.js';
import { crc32 } from './helpers/docx.js';

function samples(): Uint8Array[] {
  let seed = 7;
  const rand = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const random = new Uint8Array(70_000).map(() => Math.floor(rand() * 256));
  const text = new TextEncoder().encode('<w:p><w:r><w:t>The quick brown fox</w:t></w:r></w:p>'.repeat(3000));
  const runs = new Uint8Array(100_000).map((_, i) => (i >> 10) & 3);
  return [new Uint8Array(0), new Uint8Array([42]), random, text, runs];
}

const codecs = [
  { name: 'node:zlib', codec: nodeCompression },
  { name: 'pako', codec: pakoCompression },
];

describe.each(codecs)('$name compression', ({ codec }) => {
  it('interoperates with raw and wrapped zlib, including empty output', () => {
    for (const data of samples()) {
      expect(new Uint8Array(inflateSync(codec.deflateZlib(data)))).toEqual(data);
      for (const level of [0, 1, 9]) {
        expect(new Uint8Array(codec.inflateRaw(deflateRawSync(data, { level }), data.length))).toEqual(data);
        expect(new Uint8Array(codec.inflateZlib(deflateSync(data, { level }), data.length))).toEqual(data);
      }
    }
  });

  it('enforces output limits, including zero, and rejects incomplete streams', () => {
    const input = Uint8Array.of(42, 33);
    for (const [decode, compressed] of [
      [codec.inflateRaw, deflateRawSync(input)],
      [codec.inflateZlib, deflateSync(input)],
    ] as const) {
      for (const limit of [0, 1]) expect(() => decode(compressed, limit)).toThrow(/limit/);
      expect(() => decode(compressed.subarray(0, -1), input.length)).toThrow(/truncated/);
    }
  });

  it('streams multiple chunks and stops before forwarding excess bytes', async () => {
    const input = samples()[3];
    const data = deflateSync(input);
    const chunks: Uint8Array[] = [];
    await codec.inflateZlibStream(data, input.length, (chunk) => chunks.push(chunk.slice()));
    expect(chunks.length).toBeGreaterThan(1);
    expect(Buffer.concat(chunks).equals(input)).toBe(true);
    let received = 0;
    const run = async (source: Uint8Array, limit: number) =>
      codec.inflateZlibStream(source, limit, (chunk) => { received += chunk.length; });
    await expect(run(data, input.length - 1)).rejects.toThrow(/limit/);
    expect(received).toBeLessThan(input.length);
    await expect(run(data.subarray(0, -1), input.length)).rejects.toThrow();
  });
});

describe('XML parser', () => {
  it('normalises namespace prefixes by URI', () => {
    const doc = parseXml('<x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><x:body><x:p/></x:body></x:document>');
    expect(doc.name).toBe('w:document');
    expect(children(child(doc, 'w:body'), 'w:p')).toHaveLength(1);
  });

  it('reads Strict OOXML namespaces like Transitional ones', () => {
    const doc = parseXml('<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body/></w:document>');
    expect(child(doc, 'w:body')).toBeDefined();
  });

  it('decodes entities, character references and CDATA', () => {
    const doc = parseXml(
      '<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" a="&quot;x&quot;">A &amp; B &lt;&#169;&#x1F600;&gt;<![CDATA[<raw>]]></w:t>',
    );
    const t = doc;
    expect(t.attrs.a).toBe('"x"');
    expect(textContent(t)).toBe('A & B <©😀><raw>');
  });

  it('keeps whitespace-only text in w:t but drops it between elements', () => {
    const doc = parseXml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n  <w:t xml:space="preserve">   </w:t>\n</w:r>');
    const r = doc;
    expect(r.children).toHaveLength(1);
    expect(textContent(r)).toBe('   ');
  });

  it('picks a supported mc:Choice and falls back otherwise', () => {
    const ns =
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:zz="urn:unknown"';
    const ok = parseXml(`<mc:AlternateContent ${ns}><mc:Choice Requires="wps"><a/></mc:Choice><mc:Fallback><b/></mc:Fallback></mc:AlternateContent>`);
    expect(resolveAlternateContent(ok).map((e) => e.name)).toEqual(['a']);
    const unknown = parseXml(`<mc:AlternateContent ${ns}><mc:Choice Requires="zz"><a/></mc:Choice><mc:Fallback><b/></mc:Fallback></mc:AlternateContent>`);
    expect(resolveAlternateContent(unknown).map((e) => e.name)).toEqual(['b']);
  });
});

/** Encode rows with filter 0, or accept scanline bytes that already include the filters. */
function png(width: number, height: number, colorType: number, rows: number[][] | Uint8Array, extra: Array<[string, number[]]> = [], bitDepth = 8): Uint8Array {
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length);
    const dv = new DataView(out.buffer);
    dv.setUint32(0, data.length);
    const typed = new Uint8Array([...type].map((c) => c.charCodeAt(0)));
    out.set(typed, 4);
    out.set(data, 8);
    const crcInput = new Uint8Array(4 + data.length);
    crcInput.set(typed);
    crcInput.set(data, 4);
    dv.setUint32(8 + data.length, crc32(crcInput));
    return out;
  };
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, width);
  dv.setUint32(4, height);
  ihdr[8] = bitDepth;
  ihdr[9] = colorType;
  const raw = rows instanceof Uint8Array ? rows : new Uint8Array(rows.flatMap((r) => [0, ...r]));
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    ...extra.map(([t, d]) => chunk(t, new Uint8Array(d))),
    chunk('IDAT', new Uint8Array(deflateSync(raw))),
    chunk('IEND', new Uint8Array(0)),
  ];
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

describe.each(codecs)('$name PNG decoding', ({ codec }) => {
  it('passes opaque RGB scanlines through with a PNG predictor', async () => {
    const img = (await toPdfImage(png(2, 1, 2, [[255, 0, 0, 0, 0, 255]]), codec))!;
    expect(img).toMatchObject({ width: 2, height: 1, colorSpace: '/DeviceRGB', bitsPerComponent: 8, filter: 'FlateDecode' });
    expect(img.decodeParms).toContain('/Predictor 15');
    expect(img.smask).toBeUndefined();
  });

  it('splits RGBA into colour and a soft mask', async () => {
    const img = (await toPdfImage(png(2, 1, 6, [[10, 20, 30, 255, 40, 50, 60, 0]]), codec))!;
    expect(Array.from(inflateSync(img.data))).toEqual([10, 20, 30, 40, 50, 60]);
    expect(Array.from(inflateSync(img.smask!))).toEqual([255, 0]);
  });

  it('compresses decoded colour with a lossless predictor and keeps alpha intact', async () => {
    const width = 256, height = 48;
    const rows = Array.from({ length: height }, (_, y) => Array.from({ length: width }, (_, x) =>
      [(x + y * 2) & 255, (x * 2 + y * 3) & 255, (x * 3 + y * 5) & 255, 160 + y],
    ));
    const image = (await toPdfImage(png(width, height, 6, rows.map((row) => row.flat())), codec))!;
    expect(image.decodeParms).toContain('/Predictor 15');
    const filtered = inflateSync(image.data);
    const rgb = Buffer.alloc(width * height * 3);
    const stride = width * 3;
    for (let y = 0; y < height; y++) {
      const filter = filtered[y * (stride + 1)];
      for (let x = 0; x < stride; x++) {
        const pos = y * stride + x;
        const prev = filter === 1 ? (x < 3 ? 0 : rgb[pos - 3]) : (y === 0 ? 0 : rgb[pos - stride]);
        rgb[pos] = filtered[y * (stride + 1) + x + 1] + prev;
      }
    }
    expect(rgb.equals(Buffer.from(rows.flatMap((row) => row.flatMap((pixel) => pixel.slice(0, 3)))))).toBe(true);
    expect(inflateSync(image.smask!).equals(Buffer.from(rows.flatMap((row) => row.map((pixel) => pixel[3]))))).toBe(true);
  });

  it('applies a palette transparency chunk', async () => {
    const img = (await toPdfImage(png(3, 1, 3, [[0, 1, 1]], [['PLTE', [255, 0, 0, 0, 255, 0]], ['tRNS', [255, 0]]]), codec))!;
    expect(Array.from(inflateSync(img.data))).toEqual([255, 0, 0, 0, 255, 0, 0, 255, 0]);
    expect(Array.from(inflateSync(img.smask!))).toEqual([255, 0, 0]);
  });

  it.each([2, 6])('requires exact scanline lengths for PNG colour type %i', async (colorType) => {
    const channels = colorType === 2 ? 3 : 4;
    await expect(toPdfImage(png(1, 2, colorType, [Array(channels).fill(0)]), codec)).rejects.toThrow(/truncated scanlines/);
    await expect(toPdfImage(png(1, 1, colorType, [Array(channels + 1).fill(0)]), codec)).rejects.toThrow(/limit/);
  });

  it('preserves source resolution and individual mask pixels in a large streamed PNG', async () => {
    const width = 4096;
    const height = 2049;
    const scanlines = new Uint8Array((1 + width / 8) * height);
    scanlines[scanlines.length - 1] = 0x55;
    const input = png(width, height, 0, scanlines, [['tRNS', [0, 0]]], 1);
    let chunks = 0;
    const image = await toPdfImage(input, {
      ...codec,
      inflateZlib() { throw new Error('Unexpected whole-image inflate'); },
      inflateZlibStream(data, limit, sink) {
        return codec.inflateZlibStream(data, limit, (chunk) => {
          chunks++;
          sink(chunk);
        });
      },
    });
    expect(chunks).toBeGreaterThan(1);
    expect(image).toMatchObject({ width, height });
    const alpha = inflateSync(image!.smask!);
    expect(alpha.length).toBe(width * height);
    expect([...alpha.subarray(-5)]).toEqual([255, 0, 255, 0, 255]);
  });

  it('checks the length of each nonempty Adam7 pass', async () => {
    // A 2-by-2 grayscale image has three nonempty passes with 1, 1, and 2 samples.
    const image = png(2, 2, 0, [[10], [20], [30, 40]]);
    image[28] = 1;
    new DataView(image.buffer).setUint32(29, crc32(image.subarray(12, 29)));
    const decoded = (await toPdfImage(image, codec))!;
    expect(Array.from(inflateSync(decoded.data))).toEqual([10, 10, 10, 20, 20, 20, 30, 30, 30, 40, 40, 40]);
    await expect(toPdfImage(image.subarray(0, -1), codec)).rejects.toThrow(/truncated|missing end/);
  });

  it('rejects formats it cannot draw', async () => {
    expect(await toPdfImage(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"></svg>'), codec)).toBeNull();
  });
});

/** A one-row BMP: `header` is the info header's size; `masks` follow a 40-byte header or fill a larger one's. */
function bmp(bpp: number, compression: number, pixels: number[], masks: number[] = [], header = 40): Uint8Array {
  const width = pixels.length / (bpp / 8);
  const offset = 14 + header + (header === 40 ? masks.length * 4 : 0);
  const out = new Uint8Array(offset + Math.ceil(pixels.length / 4) * 4);
  const dv = new DataView(out.buffer);
  out.set([0x42, 0x4d]);
  dv.setUint32(2, out.length, true);
  dv.setUint32(10, offset, true);
  dv.setUint32(14, header, true);
  dv.setInt32(18, width, true);
  dv.setInt32(22, 1, true);
  dv.setUint16(26, 1, true);
  dv.setUint16(28, bpp, true);
  dv.setUint32(30, compression, true);
  masks.forEach((m, i) => dv.setUint32(54 + i * 4, m, true));
  out.set(pixels, offset);
  return out;
}

describe('raster decoding', () => {
  const rgb = (img: { data: Uint8Array } | null) => Array.from(inflateSync(img!.data));

  it('reads BMP pixels through their channel masks', async () => {
    // 32-bit BI_BITFIELDS with the bytes in R, G, B, unused order.
    expect(rgb(await toPdfImage(bmp(32, 3, [10, 20, 30, 0], [0xff, 0xff00, 0xff0000]), nodeCompression))).toEqual([10, 20, 30]);
    // 16-bit: 5-6-5 from masks, and the default 5-5-5 when uncompressed.
    const rgb565 = await toPdfImage(bmp(16, 3, [0x00, 0xf8, 0xe0, 0x07, 0x1f, 0x00], [0xf800, 0x07e0, 0x001f]), nodeCompression);
    expect(rgb(rgb565)).toEqual([255, 0, 0, 0, 255, 0, 0, 0, 255]);
    expect(rgb(await toPdfImage(bmp(16, 0, [0x00, 0x7c, 0x10, 0x42]), nodeCompression))).toEqual([255, 0, 0, 132, 132, 132]);
  });

  it('takes BMP alpha from an alpha mask, treating an all-zero channel as opaque', async () => {
    const bgra = [0xff0000, 0xff00, 0xff, 0xff000000];
    const pixels = [30, 20, 10, 128, 60, 50, 40, 255];
    for (const image of [bmp(32, 3, pixels, bgra, 108), bmp(32, 6, pixels, bgra)]) {
      const img = await toPdfImage(image, nodeCompression);
      expect(rgb(img)).toEqual([10, 20, 30, 40, 50, 60]);
      expect(Array.from(inflateSync(img!.smask!))).toEqual([128, 255]);
    }
    // Without an alpha mask the fourth byte is unused, and an alpha channel that is all zero is ignored.
    expect((await toPdfImage(bmp(32, 3, pixels, bgra.slice(0, 3)), nodeCompression))?.smask).toBeUndefined();
    expect((await toPdfImage(bmp(32, 3, [30, 20, 10, 0, 60, 50, 40, 0], bgra, 108), nodeCompression))?.smask).toBeUndefined();
  });

  it('undoes TIFF horizontal prediction without writing to the source bytes', () => {
    // Three grey pixels, uncompressed, stored as differences: 10, +5, +5.
    const entries = [[256, 3], [257, 1], [258, 8], [259, 1], [262, 1], [273, 134], [277, 1], [278, 1], [279, 3], [317, 2]];
    const tiff = new Uint8Array(137);
    const dv = new DataView(tiff.buffer);
    tiff.set([0x49, 0x49, 42, 0]);
    dv.setUint32(4, 8, true);
    dv.setUint16(8, entries.length, true);
    entries.forEach(([tag, value], i) => {
      dv.setUint16(10 + i * 12, tag, true);
      dv.setUint16(12 + i * 12, 4, true);
      dv.setUint32(14 + i * 12, 1, true);
      dv.setUint32(18 + i * 12, value, true);
    });
    tiff.set([10, 5, 5], 134);
    const source = tiff.slice();
    const grey = () => Array.from(decodeTiff(tiff, nodeCompression.inflateZlib)!.data.filter((_, i) => i % 4 === 0));
    expect(grey()).toEqual([10, 15, 20]);
    expect(grey()).toEqual([10, 15, 20]);
    expect(tiff).toEqual(source);
  });

  it('skips a JPEG whose frame header has no width or height', async () => {
    // SOI, a baseline frame header for three 8-bit components, then the start of scan.
    const jpeg = (w: number, h: number) => new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, h >> 8, h & 255, w >> 8, w & 255, 3, 1, 0x11, 0, 2, 0x11, 1, 3, 0x11, 1, 0xff, 0xda, 0, 2]);
    expect(await toPdfImage(jpeg(2, 3), nodeCompression, undefined, false)).toMatchObject({ width: 2, height: 3, filter: 'DCTDecode' });
    expect(await toPdfImage(jpeg(0, 3), nodeCompression, undefined, false)).toBeNull();
    expect(await toPdfImage(jpeg(2, 0), nodeCompression, undefined, false)).toBeNull();
  });

  it('reports no size for truncated PNG, GIF and BMP headers', () => {
    const cases: Array<[Uint8Array, number]> = [
      [png(2, 3, 0, [[0, 0], [0, 0], [0, 0]]), 24],
      [new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 2, 0, 3, 0, 0, 0, 0]), 10],
      [bmp(24, 0, [0, 0, 0, 0, 0, 0]).map((b, i) => (i === 22 ? 3 : b)), 26],
    ];
    for (const [image, needed] of cases) {
      expect(imageSize(image)).toEqual({ width: 2, height: 3 });
      expect(imageSize(image.subarray(0, needed - 1))).toBeNull();
    }
  });
});
