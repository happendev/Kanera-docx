import { ImageBudget } from './budget.js';
import type { InflateBytes } from '../util/compression.js';

/**
 * Baseline TIFF decoding to RGBA for the first image in the file: uncompressed, LZW, Deflate and
 * PackBits strips or tiles; bilevel, grey, palette, RGB and CMYK pixels of 1 to 16 bits; horizontal
 * prediction; associated or unassociated alpha. CCITT fax and JPEG compression are not supported.
 */
export function decodeTiff(d: Uint8Array, inflateZlib: InflateBytes, budget = new ImageBudget()): { width: number; height: number; data: Uint8Array } | null {
  if (d.length < 8) return null;
  const le = d[0] === 0x49;
  const u16 = (o: number) => (le ? d[o] | (d[o + 1] << 8) : (d[o] << 8) | d[o + 1]);
  const u32 = (o: number) => (le ? (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0 : ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0);
  const ifd = u32(4);
  if (ifd + 2 > d.length) return null;
  const tags = new Map<number, number[]>();
  const SIZES: Record<number, number> = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 16: 8 };
  let remainingTagValues = 1 << 20;
  for (let i = 0, n = u16(ifd); i < n; i++) {
    const e = ifd + 2 + i * 12;
    if (e + 12 > d.length) return null;
    const tag = u16(e);
    if (![256, 257, 258, 259, 262, 273, 277, 278, 279, 284, 317, 320, 322, 323, 324, 325, 338].includes(tag)) continue;
    const type = u16(e + 2);
    const count = u32(e + 4);
    const size = (SIZES[type] ?? 1) * count;
    const at = size <= 4 ? e + 8 : u32(e + 8);
    if (at + size > d.length || !SIZES[type]) return null;
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
  const one = (tag: number, def: number) => tags.get(tag)?.[0] ?? def;
  const width = one(256, 0);
  const height = one(257, 0);

  const compression = one(259, 1);
  const photometric = one(262, 1);
  const spp = one(277, 1);
  const bps = one(258, 1);
  const planar = one(284, 1);
  const predictor = one(317, 1);
  const extra = one(338, 0);
  if (![1, 5, 8, 32946, 32773].includes(compression) || ![1, 2, 4, 8, 16].includes(bps) || !Number.isInteger(spp) || spp < 1 || spp > 5) return null;

  const tiled = tags.has(322);
  const tileW = tiled ? one(322, width) : width;
  const tileH = tiled ? one(323, height) : one(278, height);
  const offsets = tags.get(tiled ? 324 : 273) ?? [];
  const counts = tags.get(tiled ? 325 : 279) ?? [];
  if (![width, height, tileW, tileH].every((v) => Number.isSafeInteger(v) && v > 0)) return null;
  const across = Math.ceil(width / tileW);
  const down = Math.ceil(height / tileH);
  // Edge tiles include padding that the decoder still has to expand.
  budget.reserve(tiled ? across * tileW : width, tiled ? down * tileH : height);
  const planes = planar === 2 ? spp : 1;
  if (offsets.length !== across * down * planes || (counts.length && counts.length !== offsets.length)) return null;
  const chunkSpp = planar === 2 ? 1 : spp;
  const rowBytes = Math.ceil((tileW * chunkSpp * bps) / 8);
  // Sample values per pixel, scaled to 0..255, gathered from every strip or tile.
  const samples = new Uint8Array(width * height * spp);
  for (let c = 0; c < offsets.length; c++) {
    const plane = Math.floor((c * planes) / offsets.length);
    const index = c % Math.ceil(offsets.length / planes);
    const x0 = tiled ? (index % across) * tileW : 0;
    const y0 = tiled ? Math.floor(index / across) * tileH : index * tileH;
    if (y0 >= height) continue;
    const start = offsets[c];
    const end = counts.length ? start + counts[c] : d.length;
    if (start > end || end > d.length) return null;
    const raw = d.subarray(start, end);
    const rows = tiled ? tileH : Math.min(tileH, height - y0);
    let buf: Uint8Array | null;
    try {
      buf = compression === 1 ? raw : compression === 5 ? lzw(raw, rowBytes * rows) : compression === 32773 ? packBits(raw, rowBytes * rows) : inflateZlib(raw, rowBytes * rows);
    } catch {
      buf = null;
    }
    if (!buf || buf.length < rowBytes * rows) return null;
    for (let r = 0; r < rows; r++) {
      const y = y0 + r;
      if (y >= height) break;
      const row = buf.subarray(r * rowBytes, (r + 1) * rowBytes);
      if (predictor === 2 && bps === 8) for (let i = chunkSpp; i < row.length; i++) row[i] = (row[i] + row[i - chunkSpp]) & 255;
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
  const out = new Uint8Array(width * height * 4);
  const premultiplied = extra === 1;
  for (let i = 0; i < width * height; i++) {
    const p = samples.subarray(i * spp, i * spp + spp);
    let r: number, g: number, b: number;
    let a = 255;
    if (photometric === 2) {
      [r, g, b] = [p[0], p[1] ?? p[0], p[2] ?? p[0]];
      if (spp > 3) a = p[3];
    } else if (photometric === 3 && map) {
      const n = map.length / 3;
      r = map[p[0]] >> 8;
      g = map[n + p[0]] >> 8;
      b = map[2 * n + p[0]] >> 8;
    } else if (photometric === 5) {
      const k = 255 - p[3];
      [r, g, b] = [((255 - p[0]) * k) / 255, ((255 - p[1]) * k) / 255, ((255 - p[2]) * k) / 255];
      if (spp > 4) a = p[4];
    } else {
      r = g = b = photometric === 0 ? 255 - p[0] : p[0];
      if (spp > 1) a = p[1];
    }
    if (premultiplied && a > 0 && a < 255) [r, g, b] = [(r * 255) / a, (g * 255) / a, (b * 255) / a];
    out.set([Math.min(255, r), Math.min(255, g), Math.min(255, b), a], i * 4);
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
  for (let i = 0; i < 256; i++) {
    suffix[i] = i;
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
  const first = (code: number) => {
    let c = code;
    while (prefix[c] >= 0) c = prefix[c];
    return suffix[c];
  };
  while (bitPos + width <= total && o < expected) {
    let code = 0;
    for (let i = 0; i < width; i++, bitPos++) code = (code << 1) | ((src[bitPos >> 3] >> (7 - (bitPos & 7))) & 1);
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
        suffix[next] = first(code);
        length[next] = length[prev] + 1;
        next++;
      }
    } else if (code === next && next < 4096) {
      prefix[next] = prev;
      suffix[next] = first(prev);
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
