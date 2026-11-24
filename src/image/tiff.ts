import { ImageBudget } from './budget.js';
import { decodeFax, type FaxCoding } from './ccitt.js';
import { decodeJpegPixels } from './jpeg-decoder.js';
import type { InflateBytes } from '../util/compression.js';

/**
 * Baseline TIFF decoding to RGBA for the first image in the file: uncompressed, LZW, Deflate,
 * PackBits, CCITT fax and JPEG strips or tiles; bilevel, grey, palette, RGB and CMYK pixels of 1 to
 * 16 bits; horizontal prediction; associated or unassociated alpha. {@link tiffJpeg} and
 * {@link tiffFax} give JPEG and fax data that can be embedded without decoding.
 */
/** Decoded pixels: RGBA, or one palette index per pixel with the palette's RGB triples. */
export type TiffPixels = { width: number; height: number; data: Uint8Array } | { width: number; height: number; indices: Uint8Array; palette: Uint8Array };

/** CCITT fax data for a PDF CCITTFaxDecode filter, with its decode parameters. */
export interface TiffFax {
  width: number;
  height: number;
  data: Uint8Array;
  /** PDF's K: negative for T.6, 0 for one-dimensional T.4, positive for mixed T.4. */
  k: number;
  /** Rows start on a byte boundary (TIFF's Modified Huffman). */
  byteAlign: boolean;
  /** Every row starts with an EOL code (T.4). */
  endOfLine: boolean;
  /** Pixels coded black are 0 bits displayed white (BlackIsZero). */
  blackIs1: boolean;
}

interface Ifd {
  le: boolean;
  tags: Map<number, number[]>;
  jpegTables?: Uint8Array;
}

const TAGS = new Set([256, 257, 258, 259, 262, 266, 273, 277, 278, 279, 284, 292, 293, 317, 320, 322, 323, 324, 325, 338, 347, 513, 514]);
const FAX: Record<number, FaxCoding> = { 2: 'mh', 3: 't4', 4: 't6' };

function readIfd(d: Uint8Array): Ifd | null {
  if (d.length < 8) return null;
  const le = d[0] === 0x49;
  const u16 = (o: number) => (le ? d[o] | (d[o + 1] << 8) : (d[o] << 8) | d[o + 1]);
  const u32 = (o: number) => (le ? (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0 : ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0);
  const ifd = u32(4);
  if (ifd + 2 > d.length) return null;
  const tags = new Map<number, number[]>();
  let jpegTables: Uint8Array | undefined;
  const SIZES: Record<number, number> = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 16: 8 };
  let remainingTagValues = 1 << 20;
  for (let i = 0, n = u16(ifd); i < n; i++) {
    const e = ifd + 2 + i * 12;
    if (e + 12 > d.length) return null;
    const tag = u16(e);
    if (!TAGS.has(tag)) continue;
    const type = u16(e + 2);
    const count = u32(e + 4);
    const size = (SIZES[type] ?? 1) * count;
    const at = size <= 4 ? e + 8 : u32(e + 8);
    if (at + size > d.length || !SIZES[type]) return null;
    if (tag === 347) {
      // JPEGTables: an abbreviated JPEG stream of the tables the strips or tiles share.
      jpegTables = d.subarray(at, at + size);
      continue;
    }
    const valueCount = [273, 279, 320, 324, 325].includes(tag) ? count : Math.min(1, count);
    if (valueCount > remainingTagValues) throw new Error('TIFF tag limit exceeded');
    remainingTagValues -= valueCount;
    const vals: number[] = [];
    for (let k = 0; k < valueCount; k++) {
      if (type === 3 || type === 8) vals.push(u16(at + k * 2));
      else if (type === 4 || type === 9 || type === 16) vals.push(u32(at + k * (type === 16 ? 8 : 4)));
      else if (type === 5 || type === 10) vals.push(u32(at + k * 8) / (u32(at + k * 8 + 4) || 1));
      else vals.push(d[at + k]);
    }
    tags.set(tag, vals);
  }
  return { le, tags, jpegTables };
}

/** The strips or tiles of an image: their layout, and each one's bytes. */
interface Chunks {
  width: number;
  height: number;
  tiled: boolean;
  tileW: number;
  tileH: number;
  across: number;
  data: Uint8Array[];
}

function chunks(d: Uint8Array, { tags }: Ifd, planes: number): Chunks | null {
  const one = (tag: number, def: number) => tags.get(tag)?.[0] ?? def;
  const width = one(256, 0);
  const height = one(257, 0);
  const tiled = tags.has(322);
  const tileW = tiled ? one(322, width) : width;
  const tileH = tiled ? one(323, height) : Math.min(one(278, height), height);
  const offsets = tags.get(tiled ? 324 : 273) ?? [];
  const counts = tags.get(tiled ? 325 : 279) ?? [];
  if (![width, height, tileW, tileH].every((v) => Number.isSafeInteger(v) && v > 0)) return null;
  const across = Math.ceil(width / tileW);
  const down = Math.ceil(height / tileH);
  if (offsets.length !== across * down * planes || (counts.length && counts.length !== offsets.length)) return null;
  const data: Uint8Array[] = [];
  for (let c = 0; c < offsets.length; c++) {
    const start = offsets[c];
    const end = counts.length ? start + counts[c] : d.length;
    if (start > end || end > d.length) return null;
    data.push(d.subarray(start, end));
  }
  return { width, height, tiled, tileW, tileH, across, data };
}

/**
 * A JPEG-compressed TIFF (or an old-style one that holds a whole JPEG file) as one JPEG file, for
 * embedding or decoding as a JPEG. Strips that each hold a band of rows join into one picture,
 * a restart marker between each. Null for other TIFFs and for tiles or strips that cannot join.
 */
export function tiffJpeg(d: Uint8Array): Uint8Array | null {
  const ifd = readIfd(d);
  if (!ifd) return null;
  const one = (tag: number, def: number) => ifd.tags.get(tag)?.[0] ?? def;
  const compression = one(259, 1);
  if (compression === 6) {
    // Old-style JPEG: usable when JPEGInterchangeFormat points to a complete JPEG file.
    const at = one(513, 0);
    const len = one(514, 0);
    const jpeg = at && len && at + len <= d.length ? d.subarray(at, at + len) : null;
    return jpeg && jpeg[0] === 0xff && jpeg[1] === 0xd8 ? jpeg : null;
  }
  if (compression !== 7 || one(284, 1) !== 1) return null;
  const adobe = jpegAdobeTransform(one(262, 2), one(277, 1));
  if (adobe === undefined) return null;
  const c = chunks(d, ifd, 1);
  if (!c || c.data.length === 0) return null;
  if (c.data.length === 1 && !c.tiled) return jpegWithTables(c.data[0], ifd.jpegTables, adobe);
  if (c.tiled) return null;
  return joinJpegStrips(c, ifd.jpegTables, adobe);
}

/**
 * The Adobe colour transform that tells JPEG decoders how a TIFF codes its JPEG components (0 for
 * RGB, 1 for YCbCr), null to leave them to their defaults, undefined when not supported.
 */
function jpegAdobeTransform(photometric: number, spp: number): number | null | undefined {
  if (photometric === 1 && spp === 1) return null;
  if (photometric === 2 && spp === 3) return 0;
  if (photometric === 6 && spp === 3) return 1;
  if (photometric === 5 && spp === 4) return null;
  return undefined;
}

/** Abbreviated JPEG `chunk` with the shared `tables` and an Adobe marker merged in. */
function jpegWithTables(chunk: Uint8Array, tables: Uint8Array | undefined, adobe: number | null, body = chunk.subarray(2)): Uint8Array | null {
  if (chunk[0] !== 0xff || chunk[1] !== 0xd8) return null;
  const parts: Uint8Array[] = [new Uint8Array([0xff, 0xd8])];
  if (adobe !== null && !hasAdobeMarker(chunk) && !(tables && hasAdobeMarker(tables))) {
    parts.push(new Uint8Array([0xff, 0xee, 0, 14, 0x41, 0x64, 0x6f, 0x62, 0x65, 0, 100, 0, 0, 0, 0, adobe]));
  }
  if (tables && tables[0] === 0xff && tables[1] === 0xd8) {
    const end = tables[tables.length - 2] === 0xff && tables[tables.length - 1] === 0xd9 ? tables.length - 2 : tables.length;
    parts.push(tables.subarray(2, end));
  }
  parts.push(body);
  return concat(parts);
}

function hasAdobeMarker(d: Uint8Array): boolean {
  for (const seg of segments(d)) {
    if (seg.marker === 0xee && seg.end - seg.start >= 16 && String.fromCharCode(...d.subarray(seg.start + 4, seg.start + 9)) === 'Adobe') return true;
  }
  return false;
}

/** The marker segments of a JPEG stream from after SOI up to and including SOS. */
function* segments(d: Uint8Array): Generator<{ marker: number; start: number; end: number }> {
  let p = 2;
  while (p + 4 <= d.length) {
    if (d[p] !== 0xff) return;
    const marker = d[p + 1];
    if (marker === 0xff) {
      p++;
      continue;
    }
    if (marker === 0xd8 || marker === 0xd9) return;
    const end = p + 2 + ((d[p + 2] << 8) | d[p + 3]);
    if (end > d.length) return;
    yield { marker, start: p, end };
    if (marker === 0xda) return;
    p = end;
  }
}

/**
 * Join strips that are each a sequential JPEG of the same width, components and tables into one
 * JPEG: a strip's entropy-coded data is a restart interval, so the strips' data in a row, restart
 * markers between, decodes as one picture without recompression.
 */
function joinJpegStrips(c: Chunks, tables: Uint8Array | undefined, adobe: number | null): Uint8Array | null {
  let header: Uint8Array | null = null;
  let sofAt = 0;
  let interval = 0;
  const data: Uint8Array[] = [];
  let rows = 0;
  for (let i = 0; i < c.data.length; i++) {
    const strip = c.data[i];
    if (strip[0] !== 0xff || strip[1] !== 0xd8) return null;
    let sof = -1;
    let sos = -1;
    for (const seg of segments(strip)) {
      if (seg.marker === 0xdd || (seg.marker >= 0xc2 && seg.marker <= 0xcf && seg.marker !== 0xc4 && seg.marker !== 0xcc)) return null;
      if (seg.marker === 0xc0 || seg.marker === 0xc1) sof = seg.start;
      if (seg.marker === 0xda) sos = seg.end;
    }
    if (sof < 0 || sos < 0) return null;
    // One scan: the data runs to the EOI, with no other marker in it.
    let end = strip.length;
    if (strip[end - 2] === 0xff && strip[end - 1] === 0xd9) end -= 2;
    for (let p = sos; p + 1 < end; p++) if (strip[p] === 0xff && strip[p + 1] !== 0 && strip[p + 1] !== 0xff) return null;
    const stripRows = (strip[sof + 5] << 8) | strip[sof + 6];
    const width = (strip[sof + 7] << 8) | strip[sof + 8];
    const last = i === c.data.length - 1;
    if (width !== c.width || stripRows === 0 || (last ? stripRows !== c.height - rows : stripRows !== c.tileH)) return null;
    rows += stripRows;
    if (!header) {
      // Every strip but the last holds whole rows of MCUs.
      const n = strip[sof + 9];
      let hmax = 1;
      let vmax = 1;
      for (let k = 0; k < n; k++) {
        hmax = Math.max(hmax, strip[sof + 11 + k * 3] >> 4);
        vmax = Math.max(vmax, strip[sof + 11 + k * 3] & 15);
      }
      if (n === 1) hmax = vmax = 1;
      if (c.tileH % (8 * vmax)) return null;
      interval = Math.ceil(width / (8 * hmax)) * (c.tileH / (8 * vmax));
      if (interval > 0xffff) return null;
      header = strip.subarray(0, sos);
      sofAt = sof;
    } else {
      // The same headers apart from the strip's height.
      const h = strip.subarray(0, sos);
      if (h.length !== header.length || sof !== sofAt) return null;
      for (let k = 0; k < h.length; k++) if (h[k] !== header[k] && k !== sof + 5 && k !== sof + 6) return null;
    }
    if (i > 0) data.push(new Uint8Array([0xff, 0xd0 + ((i - 1) & 7)]));
    data.push(strip.subarray(sos, end));
  }
  if (!header) return null;
  const head = header.slice();
  head[sofAt + 5] = c.height >> 8;
  head[sofAt + 6] = c.height & 255;
  // DRI goes just before the SOS segment.
  let sosStart = 0;
  for (const seg of segments(head)) if (seg.marker === 0xda) sosStart = seg.start;
  const body = concat([head.subarray(2, sosStart), new Uint8Array([0xff, 0xdd, 0, 4, interval >> 8, interval & 255]), head.subarray(sosStart), ...data, new Uint8Array([0xff, 0xd9])]);
  return jpegWithTables(head, tables, adobe, body);
}

/** A fax-coded TIFF of one strip, for embedding with PDF's CCITTFaxDecode filter. */
export function tiffFax(d: Uint8Array): TiffFax | null {
  const ifd = readIfd(d);
  if (!ifd) return null;
  const one = (tag: number, def: number) => ifd.tags.get(tag)?.[0] ?? def;
  const coding = FAX[one(259, 1)];
  const photometric = one(262, 0);
  const t4 = one(292, 0);
  if (!coding || one(277, 1) !== 1 || one(258, 1) !== 1 || (photometric !== 0 && photometric !== 1)) return null;
  // Uncompressed mode, which few PDF readers support.
  if ((coding === 't4' && t4 & 2) || (coding === 't6' && one(293, 0) & 2)) return null;
  const c = chunks(d, ifd, 1);
  if (!c || c.tiled || c.data.length !== 1) return null;
  const data = one(266, 1) === 2 ? reverseBits(c.data[0]) : c.data[0];
  // PDF readers find each T.4 row by its EOL, so the first row must have one.
  if (coding === 't4' && !startsWithEol(data)) return null;
  return {
    width: c.width,
    height: c.height,
    data,
    k: coding === 't6' ? -1 : coding === 't4' && t4 & 1 ? 1 : 0,
    byteAlign: coding === 'mh',
    endOfLine: coding === 't4',
    blackIs1: photometric === 1,
  };
}

function startsWithEol(d: Uint8Array): boolean {
  let bit = 0;
  const at = (i: number) => ((d[i >> 3] ?? 0) >> (7 - (i & 7))) & 1;
  while (bit < d.length * 8 && at(bit) === 0) bit++;
  return bit >= 11 && bit < d.length * 8;
}

/** Bytes with their bit order reversed, for TIFF's FillOrder 2. */
function reverseBits(d: Uint8Array): Uint8Array {
  const out = new Uint8Array(d.length);
  for (let i = 0; i < d.length; i++) {
    let b = d[i];
    b = ((b & 0xf0) >> 4) | ((b & 0x0f) << 4);
    b = ((b & 0xcc) >> 2) | ((b & 0x33) << 2);
    out[i] = ((b & 0xaa) >> 1) | ((b & 0x55) << 1);
  }
  return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

export function decodeTiff(d: Uint8Array, inflateZlib: InflateBytes, budget = new ImageBudget()): TiffPixels | null {
  const ifd = readIfd(d);
  if (!ifd) return null;
  const { le, tags } = ifd;
  const one = (tag: number, def: number) => tags.get(tag)?.[0] ?? def;
  const compression = one(259, 1);
  const fax = FAX[compression];
  // Fax images are usually WhiteIsZero, the photometric tiffFax also assumes when the tag is missing.
  const photometric = one(262, fax ? 0 : 1);
  const spp = one(277, 1);
  const bps = one(258, 1);
  const planar = one(284, 1);
  const predictor = one(317, 1);
  const extra = one(338, 0);
  if (![1, 5, 7, 8, 32946, 32773].includes(compression) && !fax) return null;
  if (![1, 2, 4, 8, 16].includes(bps) || !Number.isInteger(spp) || spp < 1 || spp > 5) return null;
  if (fax && (bps !== 1 || spp !== 1)) return null;

  const planes = planar === 2 ? spp : 1;
  const c = chunks(d, ifd, planes);
  if (!c) return null;
  const { width, height, tiled, tileW, tileH, across } = c;
  // Edge tiles include padding that the decoder still has to expand.
  budget.reserve(tiled ? across * tileW : width, tiled ? Math.ceil(height / tileH) * tileH : height);
  if (compression === 7) return planes === 1 ? decodeJpegChunks(c, ifd) : null;
  const chunkSpp = planar === 2 ? 1 : spp;
  const rowBytes = Math.ceil((tileW * chunkSpp * bps) / 8);
  const reversed = one(266, 1) === 2;
  // Sample values per pixel, scaled to 0..255, gathered from every strip or tile.
  const samples = new Uint8Array(width * height * spp);
  for (let k = 0; k < c.data.length; k++) {
    const plane = Math.floor((k * planes) / c.data.length);
    const index = k % Math.ceil(c.data.length / planes);
    const x0 = tiled ? (index % across) * tileW : 0;
    const y0 = tiled ? Math.floor(index / across) * tileH : index * tileH;
    if (y0 >= height) continue;
    const raw = c.data[k];
    const rows = tiled ? tileH : Math.min(tileH, height - y0);
    let buf: Uint8Array | null;
    try {
      // Prediction is undone in place: uncompressed rows are copied, not written into the caller's bytes.
      if (fax) buf = decodeFax(reversed ? reverseBits(raw) : raw, tileW, rows, { coding: fax, twoDimensional: (one(292, 0) & 1) === 1 });
      else buf = compression === 1 ? (predictor === 2 ? raw.slice() : raw) : compression === 5 ? lzw(raw, rowBytes * rows) : compression === 32773 ? packBits(raw, rowBytes * rows) : inflateZlib(raw, rowBytes * rows);
    } catch {
      buf = null;
    }
    if (!buf || buf.length < rowBytes * rows) return null;
    for (let r = 0; r < rows; r++) {
      const y = y0 + r;
      if (y >= height) break;
      const row = buf.subarray(r * rowBytes, (r + 1) * rowBytes);
      if (predictor === 2 && bps === 8) for (let i = chunkSpp; i < row.length; i++) row[i] = (row[i] + row[i - chunkSpp]) & 255;
      if (bps === 8) {
        // Whole bytes: copy the row's samples instead of unpacking them one by one.
        const n = Math.min(tileW, width - x0) * chunkSpp;
        if (planar !== 2) samples.set(row.subarray(0, n), (y * width + x0) * spp);
        else for (let x = 0; x < n; x++) samples[(y * width + x0 + x) * spp + plane] = row[x];
        continue;
      }
      let bit = 0;
      const acc16 = new Uint16Array(chunkSpp);
      for (let x = 0; x < tileW; x++) {
        for (let s = 0; s < chunkSpp; s++) {
          let v: number;
          if (bps === 8) v = row[x * chunkSpp + s];
          else if (bps === 16) {
            const o = (x * chunkSpp + s) * 2;
            let w = le ? row[o] | (row[o + 1] << 8) : (row[o] << 8) | row[o + 1];
            if (predictor === 2) w = acc16[s] = x === 0 ? w : (w + acc16[s]) & 0xffff;
            v = w >> 8;
          } else {
            const byte = row[bit >> 3] ?? 0;
            const raw = (byte >> (8 - bps - (bit & 7))) & ((1 << bps) - 1);
            bit += bps;
            v = photometric === 3 ? raw : Math.round((raw * 255) / ((1 << bps) - 1));
          }
          const X = x0 + x;
          if (X < width) samples[(y * width + X) * spp + (planar === 2 ? plane : s)] = v;
        }
      }
    }
  }

  const map = tags.get(320);
  // Palette images stay indexed: a third of the data of RGB to compress, and no conversion.
  if (photometric === 3 && map && spp === 1 && bps <= 8 && map.length === 3 << bps) {
    const n = 1 << bps;
    const palette = new Uint8Array(n * 3);
    for (let i = 0; i < n; i++) {
      palette[i * 3] = map[i] >> 8;
      palette[i * 3 + 1] = map[n + i] >> 8;
      palette[i * 3 + 2] = map[2 * n + i] >> 8;
    }
    return { width, height, indices: samples, palette };
  }
  // Bilevel images too: two colours, one index per pixel.
  if ((photometric === 0 || photometric === 1) && spp === 1 && bps === 1) {
    const indices = new Uint8Array(width * height);
    for (let i = 0; i < indices.length; i++) indices[i] = samples[i] ? 1 : 0;
    const palette = photometric === 0 ? new Uint8Array([255, 255, 255, 0, 0, 0]) : new Uint8Array([0, 0, 0, 255, 255, 255]);
    return { width, height, indices, palette };
  }
  const out = new Uint8Array(width * height * 4);
  const premultiplied = extra === 1;
  for (let i = 0; i < width * height; i++) {
    const p = i * spp;
    const first = samples[p];
    let r: number, g: number, b: number;
    let a = 255;
    if (photometric === 2) {
      r = first;
      g = spp > 1 ? samples[p + 1] : first;
      b = spp > 2 ? samples[p + 2] : first;
      if (spp > 3) a = samples[p + 3];
    } else if (photometric === 3 && map) {
      const n = map.length / 3;
      r = map[first] >> 8;
      g = map[n + first] >> 8;
      b = map[2 * n + first] >> 8;
    } else if (photometric === 5) {
      if (spp < 4) return null;
      const k = 255 - samples[p + 3];
      r = ((255 - first) * k) / 255;
      g = ((255 - samples[p + 1]) * k) / 255;
      b = ((255 - samples[p + 2]) * k) / 255;
      if (spp > 4) a = samples[p + 4];
    } else {
      r = g = b = photometric === 0 ? 255 - first : first;
      if (spp > 1) a = samples[p + 1];
    }
    if (premultiplied && a > 0 && a < 255) {
      r = (r * 255) / a;
      g = (g * 255) / a;
      b = (b * 255) / a;
    }
    out[i * 4] = Math.min(255, r);
    out[i * 4 + 1] = Math.min(255, g);
    out[i * 4 + 2] = Math.min(255, b);
    out[i * 4 + 3] = a;
  }
  return { width, height, data: out };
}

/** JPEG strips or tiles decoded one by one into one RGBA image. */
function decodeJpegChunks(c: Chunks, ifd: Ifd): TiffPixels | null {
  const one = (tag: number, def: number) => ifd.tags.get(tag)?.[0] ?? def;
  const adobe = jpegAdobeTransform(one(262, 2), one(277, 1));
  if (adobe === undefined) return null;
  const { width, height, tileW, tileH, across } = c;
  const out = new Uint8Array(width * height * 4);
  for (let k = 0; k < c.data.length; k++) {
    const x0 = (k % across) * tileW;
    const y0 = Math.floor(k / across) * tileH;
    if (y0 >= height) continue;
    const jpeg = jpegWithTables(c.data[k], ifd.jpegTables, adobe);
    if (!jpeg) return null;
    let px: { width: number; height: number; data: Uint8Array };
    try {
      px = decodeJpegPixels(jpeg);
    } catch {
      return null;
    }
    const w = Math.min(px.width, width - x0);
    for (let y = 0; y < px.height && y0 + y < height; y++) {
      out.set(px.data.subarray(y * px.width * 4, (y * px.width + w) * 4), ((y0 + y) * width + x0) * 4);
    }
  }
  return { width, height, data: out };
}

/** TIFF LZW: MSB-first codes of 9 to 12 bits that widen one code early. */
function lzw(src: Uint8Array, expected: number): Uint8Array | null {
  const out = new Uint8Array(expected);
  let o = 0;
  const prefix = new Int32Array(4096);
  const suffix = new Uint8Array(4096);
  const length = new Int32Array(4096);
  // First byte of each code's string, so a new entry needs no walk along its prefixes.
  const head = new Uint8Array(4096);
  for (let i = 0; i < 256; i++) {
    suffix[i] = i;
    head[i] = i;
    length[i] = 1;
    prefix[i] = -1;
  }
  let next = 258;
  let width = 9;
  let prev = -1;
  let bitPos = 0;
  const total = src.length * 8;
  const emit = (code: number) => {
    const len = length[code];
    if (len > expected - o) throw new Error('TIFF: excess LZW data');
    let c = code;
    for (let i = len - 1; i >= 0; i--) {
      out[o + i] = suffix[c];
      c = prefix[c];
    }
    o += len;
  };
  while (bitPos + width <= total && o < expected) {
    // A code of at most 12 bits starting anywhere in a byte lies within three bytes.
    const at = bitPos >> 3;
    const window = (src[at] << 16) | ((src[at + 1] ?? 0) << 8) | (src[at + 2] ?? 0);
    const code = (window >> (24 - (bitPos & 7) - width)) & ((1 << width) - 1);
    bitPos += width;
    if (code === 257) break;
    if (code === 256) {
      next = 258;
      width = 9;
      prev = -1;
      continue;
    }
    if (prev < 0) {
      if (code > 255) break;
      emit(code);
      prev = code;
      continue;
    }
    if (code < next) {
      emit(code);
      if (next < 4096) {
        prefix[next] = prev;
        suffix[next] = head[code];
        head[next] = head[prev];
        length[next] = length[prev] + 1;
        next++;
      }
    } else if (code === next && next < 4096) {
      prefix[next] = prev;
      suffix[next] = head[prev];
      head[next] = head[prev];
      length[next] = length[prev] + 1;
      next++;
      emit(code);
    } else break;
    prev = code;
    if (next + 1 >= 1 << width && width < 12) width++;
  }
  return o === expected ? out : null;
}

function packBits(src: Uint8Array, expected: number): Uint8Array | null {
  const out = new Uint8Array(expected);
  let o = 0;
  for (let i = 0; i < src.length && o < expected; ) {
    const n = (src[i++] << 24) >> 24;
    if (n >= 0) {
      if (i + n + 1 > src.length || o + n + 1 > expected) return null;
      out.set(src.subarray(i, i + n + 1), o);
      i += n + 1;
      o += n + 1;
    } else if (n !== -128) {
      if (i === src.length || o + 1 - n > expected) return null;
      const v = src[i++];
      for (let k = 0; k < 1 - n; k++) out[o++] = v;
    }
  }
  return o === expected ? out : null;
}
