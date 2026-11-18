import { deflateRawSync, deflateSync, inflateRawSync, inflateSync } from 'node:zlib';
import { describe, expect, it } from 'vitest';
import { toPdfImage } from '../src/image/image.js';
import { deflateRaw, deflateZlib } from '../src/util/deflate.js';
import { inflateRaw, inflateZlib } from '../src/util/inflate.js';
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

describe('deflate / inflate', () => {
  it('round-trips through our own codec', () => {
    for (const s of samples()) {
      expect(inflateRaw(deflateRaw(s))).toEqual(s);
      expect(inflateZlib(deflateZlib(s))).toEqual(s);
    }
  });

  it('interoperates with zlib in both directions', () => {
    for (const s of samples()) {
      expect(new Uint8Array(inflateRawSync(deflateRaw(s)))).toEqual(s);
      expect(new Uint8Array(inflateSync(deflateZlib(s)))).toEqual(s);
      for (const level of [0, 1, 9]) {
        expect(inflateRaw(new Uint8Array(deflateRawSync(s, { level })))).toEqual(s);
        expect(inflateZlib(new Uint8Array(deflateSync(s, { level })))).toEqual(s);
      }
    }
  });

  it('actually compresses repetitive input', () => {
    const text = samples()[3];
    expect(deflateZlib(text).length).toBeLessThan(text.length / 20);
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

/** Encode a PNG from raw scanlines (filter byte 0 on every row). */
function png(width: number, height: number, colorType: number, rows: number[][], extra: Array<[string, number[]]> = [], bitDepth = 8): Uint8Array {
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
  const raw = new Uint8Array(rows.flatMap((r) => [0, ...r]));
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

describe('PNG decoding', () => {
  it('passes opaque RGB scanlines through with a PNG predictor', () => {
    const img = toPdfImage(png(2, 1, 2, [[255, 0, 0, 0, 0, 255]]))!;
    expect(img).toMatchObject({ width: 2, height: 1, colorSpace: '/DeviceRGB', bitsPerComponent: 8, filter: 'FlateDecode' });
    expect(img.decodeParms).toContain('/Predictor 15');
    expect(img.smask).toBeUndefined();
  });

  it('splits RGBA into colour and a soft mask', () => {
    const img = toPdfImage(png(2, 1, 6, [[10, 20, 30, 255, 40, 50, 60, 0]]))!;
    expect(Array.from(inflateZlib(img.data))).toEqual([10, 20, 30, 40, 50, 60]);
    expect(Array.from(inflateZlib(img.smask!))).toEqual([255, 0]);
  });

  it('applies a palette transparency chunk', () => {
    const img = toPdfImage(png(3, 1, 3, [[0, 1, 1]], [['PLTE', [255, 0, 0, 0, 255, 0]], ['tRNS', [255, 0]]]))!;
    expect(Array.from(inflateZlib(img.data))).toEqual([255, 0, 0, 0, 255, 0, 0, 255, 0]);
    expect(Array.from(inflateZlib(img.smask!))).toEqual([255, 0, 0]);
  });

  it('rejects formats it cannot draw', () => {
    expect(toPdfImage(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"></svg>'))).toBeNull();
  });
});
