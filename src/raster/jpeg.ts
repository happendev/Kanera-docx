// Baseline JPEG encoding of RGB pages: YCbCr with chroma halved both ways (4:2:0), the standard
// quantisation tables scaled by quality and the standard Huffman tables. Blank white areas, most
// of a document page, are coded straight from the canvas's painted spans without a transform.

const ZIGZAG = Int32Array.from([
  0, 1, 8, 16, 9, 2, 3, 10, 17, 24, 32, 25, 18, 11, 4, 5, 12, 19, 26, 33, 40, 48, 41, 34, 27, 20, 13, 6, 7, 14, 21, 28,
  35, 42, 49, 56, 57, 50, 43, 36, 29, 22, 15, 23, 30, 37, 44, 51, 58, 59, 52, 45, 38, 31, 39, 46, 53, 60, 61, 54, 47, 55, 62, 63,
]);

// Annex K quantisation tables, in natural (row-major) order.
const LUMA_Q = [
  16, 11, 10, 16, 24, 40, 51, 61, 12, 12, 14, 19, 26, 58, 60, 55, 14, 13, 16, 24, 40, 57, 69, 56, 14, 17, 22, 29, 51, 87, 80, 62,
  18, 22, 37, 56, 68, 109, 103, 77, 24, 35, 55, 64, 81, 104, 113, 92, 49, 64, 78, 87, 103, 121, 120, 101, 72, 92, 95, 98, 112, 100, 103, 99,
];
const CHROMA_Q = [
  17, 18, 24, 47, 99, 99, 99, 99, 18, 21, 26, 66, 99, 99, 99, 99, 24, 26, 56, 99, 99, 99, 99, 99, 47, 66, 99, 99, 99, 99, 99, 99,
  99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99,
];

// Annex K Huffman tables: code counts by length 1–16, then symbols.
const DC_LUMA = { bits: [0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0], vals: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] };
const DC_CHROMA = { bits: [0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0], vals: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] };
const AC_LUMA = {
  bits: [0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, 0x7d],
  vals: [
    0x01, 0x02, 0x03, 0x00, 0x04, 0x11, 0x05, 0x12, 0x21, 0x31, 0x41, 0x06, 0x13, 0x51, 0x61, 0x07, 0x22, 0x71, 0x14, 0x32, 0x81, 0x91, 0xa1, 0x08,
    0x23, 0x42, 0xb1, 0xc1, 0x15, 0x52, 0xd1, 0xf0, 0x24, 0x33, 0x62, 0x72, 0x82, 0x09, 0x0a, 0x16, 0x17, 0x18, 0x19, 0x1a, 0x25, 0x26, 0x27, 0x28,
    0x29, 0x2a, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4a, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58, 0x59,
    0x5a, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6a, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7a, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89,
    0x8a, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9a, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6, 0xa7, 0xa8, 0xa9, 0xaa, 0xb2, 0xb3, 0xb4, 0xb5, 0xb6,
    0xb7, 0xb8, 0xb9, 0xba, 0xc2, 0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8, 0xd9, 0xda, 0xe1, 0xe2,
    0xe3, 0xe4, 0xe5, 0xe6, 0xe7, 0xe8, 0xe9, 0xea, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9, 0xfa,
  ],
};
const AC_CHROMA = {
  bits: [0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, 0x77],
  vals: [
    0x00, 0x01, 0x02, 0x03, 0x11, 0x04, 0x05, 0x21, 0x31, 0x06, 0x12, 0x41, 0x51, 0x07, 0x61, 0x71, 0x13, 0x22, 0x32, 0x81, 0x08, 0x14, 0x42, 0x91,
    0xa1, 0xb1, 0xc1, 0x09, 0x23, 0x33, 0x52, 0xf0, 0x15, 0x62, 0x72, 0xd1, 0x0a, 0x16, 0x24, 0x34, 0xe1, 0x25, 0xf1, 0x17, 0x18, 0x19, 0x1a, 0x26,
    0x27, 0x28, 0x29, 0x2a, 0x35, 0x36, 0x37, 0x38, 0x39, 0x3a, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4a, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58,
    0x59, 0x5a, 0x63, 0x64, 0x65, 0x66, 0x67, 0x68, 0x69, 0x6a, 0x73, 0x74, 0x75, 0x76, 0x77, 0x78, 0x79, 0x7a, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87,
    0x88, 0x89, 0x8a, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9a, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6, 0xa7, 0xa8, 0xa9, 0xaa, 0xb2, 0xb3, 0xb4,
    0xb5, 0xb6, 0xb7, 0xb8, 0xb9, 0xba, 0xc2, 0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xd2, 0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8, 0xd9, 0xda,
    0xe2, 0xe3, 0xe4, 0xe5, 0xe6, 0xe7, 0xe8, 0xe9, 0xea, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9, 0xfa,
  ],
};

interface Huffman {
  code: Uint16Array;
  size: Uint8Array;
}

function huffman(spec: { bits: number[]; vals: number[] }): Huffman {
  const code = new Uint16Array(256);
  const size = new Uint8Array(256);
  let k = 0;
  let c = 0;
  for (let len = 1; len <= 16; len++) {
    for (let i = 0; i < spec.bits[len - 1]; i++, k++) {
      code[spec.vals[k]] = c++;
      size[spec.vals[k]] = len;
    }
    c <<= 1;
  }
  return { code, size };
}

const TABLES = { dcL: huffman(DC_LUMA), dcC: huffman(DC_CHROMA), acL: huffman(AC_LUMA), acC: huffman(AC_CHROMA) };

// AAN DCT output scale factors, folded into the quantisation divisors.
const AAN = [1, 1.387039845, 1.306562965, 1.175875602, 1, 0.785694958, 0.5411961, 0.275899379];

/** Quantisation table for `quality` (IJG scaling), natural order. */
function scaled(base: number[], quality: number): Uint8Array {
  const s = quality < 50 ? 5000 / quality : 200 - quality * 2;
  return Uint8Array.from(base, (q) => Math.max(1, Math.min(255, Math.floor((q * s + 50) / 100))));
}

/** Multipliers taking AAN DCT outputs (natural order) to quantised coefficients. */
function divisors(q: Uint8Array): Float32Array {
  const out = new Float32Array(64);
  for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++) out[r * 8 + c] = 1 / (q[r * 8 + c] * AAN[r] * AAN[c] * 8);
  return out;
}

/** Growable output with JPEG's bit packing: 0xFF data bytes are followed by a zero byte. */
class Writer {
  bytes = new Uint8Array(1 << 16);
  n = 0;
  private acc = 0;
  private bits = 0;

  private grow(extra: number): void {
    if (this.n + extra <= this.bytes.length) return;
    const next = new Uint8Array(Math.max(this.bytes.length * 2, this.n + extra));
    next.set(this.bytes.subarray(0, this.n));
    this.bytes = next;
  }

  raw(data: ArrayLike<number>): void {
    this.grow(data.length);
    for (let i = 0; i < data.length; i++) this.bytes[this.n++] = data[i];
  }

  u16(v: number): void {
    this.raw([v >> 8, v & 255]);
  }

  /** Append the low `size` bits of `value`, most significant first. */
  put(value: number, size: number): void {
    this.acc = (this.acc << size) | (value & ((1 << size) - 1));
    this.bits += size;
    if (this.bits < 8) return;
    this.grow(8);
    while (this.bits >= 8) {
      this.bits -= 8;
      const b = (this.acc >> this.bits) & 255;
      this.bytes[this.n++] = b;
      if (b === 255) this.bytes[this.n++] = 0;
    }
    this.acc &= (1 << this.bits) - 1;
  }

  /** Pad the last byte with one bits. */
  flush(): void {
    if (this.bits > 0) this.put(0x7f, 8 - this.bits);
  }
}

/** Bits needed for |v|. */
function category(v: number): number {
  return 32 - Math.clz32(v < 0 ? -v : v);
}

const block = new Float32Array(64);
const coef = new Int32Array(64);

/** Forward DCT (AAN) of `block` in place, then quantise into `coef` in zigzag order. */
function transform(div: Float32Array): void {
  const d = block;
  for (let o = 0; o < 64; o += 8) {
    const t0 = d[o] + d[o + 7];
    const t7 = d[o] - d[o + 7];
    const t1 = d[o + 1] + d[o + 6];
    const t6 = d[o + 1] - d[o + 6];
    const t2 = d[o + 2] + d[o + 5];
    const t5 = d[o + 2] - d[o + 5];
    const t3 = d[o + 3] + d[o + 4];
    const t4 = d[o + 3] - d[o + 4];
    let a = t0 + t3;
    const b = t0 - t3;
    let c = t1 + t2;
    const e = t1 - t2;
    d[o] = a + c;
    d[o + 4] = a - c;
    const z1 = (e + b) * 0.707106781;
    d[o + 2] = b + z1;
    d[o + 6] = b - z1;
    a = t4 + t5;
    c = t5 + t6;
    const f = t6 + t7;
    const z5 = (a - f) * 0.382683433;
    const z2 = 0.5411961 * a + z5;
    const z4 = 1.306562965 * f + z5;
    const z3 = c * 0.707106781;
    const z11 = t7 + z3;
    const z13 = t7 - z3;
    d[o + 5] = z13 + z2;
    d[o + 3] = z13 - z2;
    d[o + 1] = z11 + z4;
    d[o + 7] = z11 - z4;
  }
  for (let o = 0; o < 8; o++) {
    const t0 = d[o] + d[o + 56];
    const t7 = d[o] - d[o + 56];
    const t1 = d[o + 8] + d[o + 48];
    const t6 = d[o + 8] - d[o + 48];
    const t2 = d[o + 16] + d[o + 40];
    const t5 = d[o + 16] - d[o + 40];
    const t3 = d[o + 24] + d[o + 32];
    const t4 = d[o + 24] - d[o + 32];
    let a = t0 + t3;
    const b = t0 - t3;
    let c = t1 + t2;
    const e = t1 - t2;
    d[o] = a + c;
    d[o + 32] = a - c;
    const z1 = (e + b) * 0.707106781;
    d[o + 16] = b + z1;
    d[o + 48] = b - z1;
    a = t4 + t5;
    c = t5 + t6;
    const f = t6 + t7;
    const z5 = (a - f) * 0.382683433;
    const z2 = 0.5411961 * a + z5;
    const z4 = 1.306562965 * f + z5;
    const z3 = c * 0.707106781;
    const z11 = t7 + z3;
    const z13 = t7 - z3;
    d[o + 40] = z13 + z2;
    d[o + 24] = z13 - z2;
    d[o + 8] = z11 + z4;
    d[o + 56] = z11 - z4;
  }
  for (let i = 0; i < 64; i++) {
    const z = ZIGZAG[i];
    coef[i] = Math.round(d[z] * div[z]);
  }
}

/** Code `coef` (zigzag order) against the previous DC value; returns this block's DC. */
function code(w: Writer, prevDc: number, dc: Huffman, ac: Huffman): number {
  const diff = coef[0] - prevDc;
  const n = category(diff);
  w.put(dc.code[n], dc.size[n]);
  if (n) w.put(diff < 0 ? diff - 1 : diff, n);
  let last = 63;
  while (last > 0 && !coef[last]) last--;
  let run = 0;
  for (let i = 1; i <= last; i++) {
    const v = coef[i];
    if (!v) {
      run++;
      continue;
    }
    while (run > 15) {
      w.put(ac.code[0xf0], ac.size[0xf0]);
      run -= 16;
    }
    const s = category(v);
    const sym = (run << 4) | s;
    w.put(ac.code[sym], ac.size[sym]);
    w.put(v < 0 ? v - 1 : v, s);
    run = 0;
  }
  if (last < 63) w.put(ac.code[0], ac.size[0]);
  return coef[0];
}

/** Whether all of `block` is one value. */
function flatBlock(): boolean {
  const v = block[0];
  for (let i = 1; i < 64; i++) if (block[i] !== v) return false;
  return true;
}

/** Transform and code `block`, or for a flat one just its DC value (the transform gives no more). */
function codeBlock(w: Writer, div: Float32Array, prevDc: number, dc: Huffman, ac: Huffman): number {
  if (flatBlock()) return codeFlat(w, block[0], div, prevDc, dc, ac);
  transform(div);
  return code(w, prevDc, dc, ac);
}

/** Code a block whose every sample is `v` (level-shifted): a DC value and no AC terms. */
function codeFlat(w: Writer, v: number, div: Float32Array, prevDc: number, dc: Huffman, ac: Huffman): number {
  coef.fill(0);
  coef[0] = Math.round(v * 64 * div[0]);
  return code(w, prevDc, dc, ac);
}

/**
 * A JPEG of an RGB image at `quality` (1–100). Pixels outside each row's [lo, hi) are white.
 */
export function jpegFile(width: number, height: number, rgb: Uint8Array, lo: Int32Array, hi: Int32Array, quality: number, dpi: number): Uint8Array<ArrayBuffer> {
  const ql = scaled(LUMA_Q, quality);
  const qc = scaled(CHROMA_Q, quality);
  const divL = divisors(ql);
  const divC = divisors(qc);
  const w = new Writer();
  w.raw([0xff, 0xd8]);
  // JFIF, with the resolution in dots per inch.
  const d = Math.max(1, Math.min(65535, Math.round(dpi)));
  w.raw([0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 1, d >> 8, d & 255, d >> 8, d & 255, 0, 0]);
  w.raw([0xff, 0xdb, 0, 132, 0]);
  for (let i = 0; i < 64; i++) w.raw([ql[ZIGZAG[i]]]);
  w.raw([1]);
  for (let i = 0; i < 64; i++) w.raw([qc[ZIGZAG[i]]]);
  w.raw([0xff, 0xc0, 0, 17, 8]);
  w.u16(height);
  w.u16(width);
  w.raw([3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1]);
  w.raw([0xff, 0xc4]);
  w.u16(2 + 4 * 17 + DC_LUMA.vals.length + AC_LUMA.vals.length + DC_CHROMA.vals.length + AC_CHROMA.vals.length);
  for (const [cls, spec] of [[0x00, DC_LUMA], [0x10, AC_LUMA], [0x01, DC_CHROMA], [0x11, AC_CHROMA]] as const) {
    w.raw([cls]);
    w.raw(spec.bits);
    w.raw(spec.vals);
  }
  w.raw([0xff, 0xda, 0, 12, 3, 1, 0x00, 2, 0x11, 3, 0x11, 0, 63, 0]);

  const { dcL, dcC, acL, acC } = TABLES;
  let dcY = 0;
  let dcCb = 0;
  let dcCr = 0;
  const stride = width * 3;
  const ys = new Float32Array(256);
  const cbs = new Float32Array(64);
  const crs = new Float32Array(64);
  for (let my = 0; my < height; my += 16) {
    // The painted span across these 16 rows; MCUs outside it are white.
    let spanLo = width;
    let spanHi = 0;
    for (let y = my; y < Math.min(height, my + 16); y++) {
      if (lo[y] < spanLo) spanLo = lo[y];
      if (hi[y] > spanHi) spanHi = hi[y];
    }
    for (let mx = 0; mx < width; mx += 16) {
      if (mx + 16 <= spanLo || mx >= spanHi) {
        // White: Y 255, Cb and Cr 128 (level-shifted to 127 and 0).
        for (let k = 0; k < 4; k++) dcY = codeFlat(w, 127, divL, dcY, dcL, acL);
        dcCb = codeFlat(w, 0, divC, dcCb, dcC, acC);
        dcCr = codeFlat(w, 0, divC, dcCr, dcC, acC);
        continue;
      }
      // One colour across the whole MCU: its blocks are flat, worked out as the general case would.
      const first = my * stride + mx * 3;
      const r0 = rgb[first];
      const g0 = rgb[first + 1];
      const b0 = rgb[first + 2];
      let uniform = true;
      for (let j = 0; j < 16 && uniform; j++) {
        const row = Math.min(height - 1, my + j) * stride;
        for (let i = 0; i < 16; i++) {
          const o = row + Math.min(width - 1, mx + i) * 3;
          if (rgb[o] !== r0 || rgb[o + 1] !== g0 || rgb[o + 2] !== b0) {
            uniform = false;
            break;
          }
        }
      }
      if (uniform) {
        const yv = Math.fround(0.299 * r0 + 0.587 * g0 + 0.114 * b0 - 128);
        const chroma = (x: number) => Math.fround((x + x + x + x) / 4);
        for (let k = 0; k < 4; k++) dcY = codeFlat(w, yv, divL, dcY, dcL, acL);
        dcCb = codeFlat(w, chroma(-0.168736 * r0 - 0.331264 * g0 + 0.5 * b0), divC, dcCb, dcC, acC);
        dcCr = codeFlat(w, chroma(0.5 * r0 - 0.418688 * g0 - 0.081312 * b0), divC, dcCr, dcC, acC);
        continue;
      }
      // Each chroma sample averages a 2 × 2 group; MCUs past the right or bottom edge repeat its last pixels.
      const edge = mx + 16 > width || my + 16 > height;
      for (let cj = 0; cj < 8; cj++) {
        for (let ci = 0; ci < 8; ci++) {
          let cb = 0;
          let cr = 0;
          for (let k = 0; k < 4; k++) {
            const j = cj * 2 + (k >> 1);
            const i = ci * 2 + (k & 1);
            const o = edge ? Math.min(height - 1, my + j) * stride + Math.min(width - 1, mx + i) * 3 : (my + j) * stride + (mx + i) * 3;
            const r = rgb[o];
            const g = rgb[o + 1];
            const b = rgb[o + 2];
            ys[j * 16 + i] = 0.299 * r + 0.587 * g + 0.114 * b - 128;
            cb += -0.168736 * r - 0.331264 * g + 0.5 * b;
            cr += 0.5 * r - 0.418688 * g - 0.081312 * b;
          }
          cbs[cj * 8 + ci] = cb / 4;
          crs[cj * 8 + ci] = cr / 4;
        }
      }
      for (let k = 0; k < 4; k++) {
        const ox = (k & 1) * 8;
        const oy = (k >> 1) * 8;
        for (let j = 0; j < 8; j++) for (let i = 0; i < 8; i++) block[j * 8 + i] = ys[(oy + j) * 16 + ox + i];
        dcY = codeBlock(w, divL, dcY, dcL, acL);
      }
      block.set(cbs);
      dcCb = codeBlock(w, divC, dcCb, dcC, acC);
      block.set(crs);
      dcCr = codeBlock(w, divC, dcCr, dcC, acC);
    }
  }
  w.flush();
  w.raw([0xff, 0xd9]);
  return w.bytes.slice(0, w.n);
}
