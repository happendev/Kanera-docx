// JPEG decoding (Huffman-coded baseline, extended and progressive, 8-bit samples) with libjpeg's
// arithmetic: its integer inverse DCT, triangle ("fancy") chroma upsampling and YCbCr tables, the
// way Poppler's DCTDecode filter, built on libjpeg, reads the pictures the PDF writer embeds. CMYK
// pictures convert to RGB as Poppler's DeviceCMYK does.

/** Coefficient position (row-major) of each zig-zag index, padded so corrupt runs stop at 63. */
const NATURAL = new Int32Array(80).fill(63);
NATURAL.set([
  0, 1, 8, 16, 9, 2, 3, 10, 17, 24, 32, 25, 18, 11, 4, 5, 12, 19, 26, 33, 40, 48, 41, 34, 27, 20, 13, 6, 7, 14, 21, 28,
  35, 42, 49, 56, 57, 50, 43, 36, 29, 22, 15, 23, 30, 37, 44, 51, 58, 59, 52, 45, 38, 31, 39, 46, 53, 60, 61, 54, 47, 55, 62, 63,
]);

/** libjpeg's post-IDCT range limit: centred samples to 0-255, indexed by `value & 1023`. */
const RANGE = (() => {
  const t = new Uint8Array(1024);
  for (let i = 0; i < 1024; i++) t[i] = i < 128 ? i + 128 : i < 512 ? 255 : i < 896 ? 0 : i - 896;
  return t;
})();

/** libjpeg's YCbCr → RGB tables (jdcolor.c), 16-bit fixed point. */
const CR_R = new Int32Array(256);
const CB_B = new Int32Array(256);
const CR_G = new Int32Array(256);
const CB_G = new Int32Array(256);
for (let i = 0; i < 256; i++) {
  const x = i - 128;
  CR_R[i] = (91881 * x + 32768) >> 16;
  CB_B[i] = (116130 * x + 32768) >> 16;
  CR_G[i] = -46802 * x;
  CB_G[i] = -22554 * x + 32768;
}

/** Sums of a sample and a chroma term (-256 to 511), indexed from 256, clamped to 0–255. */
const LIMIT = Uint8Array.from({ length: 768 }, (_, i) => Math.min(255, Math.max(0, i - 256)));

const LITTLE_ENDIAN = new Uint8Array(new Uint32Array([1]).buffer)[0] === 1;

export interface JpegPixels {
  width: number;
  height: number;
  /** Colour components in the file: 1 (grey), 3 or 4 (CMYK). */
  components: number;
  /** Opaque RGBA. */
  data: Uint8Array;
}

interface Huffman {
  /** For the next 9 bits: the code's length << 8 | its symbol, or 0 for a longer code. */
  fast: Uint16Array;
  /** Largest code of each length (1-16), -1 when none; [17] is a sentinel. */
  maxcode: Int32Array;
  /** Symbol index minus code, for each length. */
  offset: Int32Array;
  symbols: Uint8Array;
}

function huffman(counts: Uint8Array, symbols: Uint8Array): Huffman {
  const fast = new Uint16Array(512);
  const maxcode = new Int32Array(18).fill(-1);
  const offset = new Int32Array(17);
  let code = 0;
  let p = 0;
  for (let len = 1; len <= 16; len++) {
    const n = counts[len - 1];
    offset[len] = p - code;
    for (let i = 0; i < n; i++, p++, code++) {
      if (len <= 9) {
        const shift = 9 - len;
        const entry = (len << 8) | symbols[p];
        for (let s = 0; s < 1 << shift; s++) fast[(code << shift) | s] = entry;
      }
    }
    if (n) maxcode[len] = code - 1;
    code <<= 1;
  }
  maxcode[17] = 0x7fffffff;
  return { fast, maxcode, offset, symbols };
}

interface Component {
  id: number;
  h: number;
  v: number;
  tq: number;
  /** Quantisation table (natural order), latched when the component's first scan starts. */
  quant: Int32Array | null;
  /** Blocks across and down that hold the component's samples, and allocated (whole MCUs). */
  bw: number;
  bh: number;
  stride: number;
  rows: number;
  coefs: Int16Array;
  /** Samples across and down (libjpeg's downsampled size). */
  sw: number;
  sh: number;
  dc: Huffman | null;
  ac: Huffman | null;
  pred: number;
}

/** Entropy-coded data: bits most significant first, with stuffed zero bytes removed. */
class Bits {
  private buf = 0;
  /** Bits held in `buf`, and how many of the lowest of them are made-up zeros past a marker. */
  private cnt = 0;
  private fake = 0;
  /** Bits ran out mid-segment; libjpeg then leaves the segment's later blocks empty. */
  short = false;
  constructor(
    private readonly d: Uint8Array,
    public pos: number,
  ) {}

  private fill(): void {
    const d = this.d;
    while (this.cnt <= 24) {
      let b = 0;
      if (!this.fake && this.pos < d.length) {
        b = d[this.pos];
        if (b === 0xff) {
          let n = this.pos + 1;
          while (n < d.length && d[n] === 0xff) n++;
          if (n < d.length && d[n] === 0) this.pos = n + 1;
          else {
            // A marker: the segment's data ends here, and zeros make up any bits still asked for.
            this.pos = n - 1;
            b = 0;
            this.fake = 8;
          }
        } else this.pos++;
      } else if (this.fake) this.fake += 8;
      else this.fake = 8;
      this.buf = (this.buf << 8) | b;
      this.cnt += 8;
    }
  }

  private take(n: number): void {
    this.cnt -= n;
    if (this.cnt < this.fake) this.short = true;
  }

  bits(n: number): number {
    if (!n) return 0;
    if (this.cnt < n) this.fill();
    this.take(n);
    return (this.buf >>> this.cnt) & ((1 << n) - 1);
  }

  bit(): number {
    if (this.cnt < 1) this.fill();
    this.take(1);
    return (this.buf >>> this.cnt) & 1;
  }

  /** The next Huffman symbol (0 for a code no table entry matches, as libjpeg does). */
  decode(h: Huffman): number {
    if (this.cnt < 16) this.fill();
    const e = h.fast[(this.buf >>> (this.cnt - 9)) & 511];
    if (e) {
      this.take(e >> 8);
      return e & 255;
    }
    for (let len = 10; len <= 16; len++) {
      const code = (this.buf >>> (this.cnt - len)) & ((1 << len) - 1);
      if (code <= h.maxcode[len]) {
        this.take(len);
        return h.symbols[(code + h.offset[len]) & 255];
      }
    }
    this.take(16);
    return 0;
  }

  /** `s` more bits as a signed value (HUFF_EXTEND). */
  receive(s: number): number {
    // Sizes past 16 bits only come from corrupt tables.
    if (s > 16) s = 16;
    const v = this.bits(s);
    return v < 1 << (s - 1) ? v - (1 << s) + 1 : v;
  }

  /** Drop buffered bits and continue after the next marker if it is a restart marker. */
  restart(): void {
    this.buf = 0;
    this.cnt = 0;
    this.fake = 0;
    this.short = false;
    const d = this.d;
    let p = this.pos;
    while (p + 1 < d.length && !(d[p] === 0xff && d[p + 1] !== 0 && d[p + 1] !== 0xff)) p++;
    if (p + 1 < d.length && d[p + 1] >= 0xd0 && d[p + 1] <= 0xd7) p += 2;
    this.pos = p;
  }

  /** Where the scan's data ends: the next marker other than a restart marker. */
  end(): number {
    const d = this.d;
    let p = Math.max(0, this.pos - 2);
    while (p + 1 < d.length && !(d[p] === 0xff && d[p + 1] !== 0 && d[p + 1] !== 0xff && !(d[p + 1] >= 0xd0 && d[p + 1] <= 0xd7))) p++;
    return p;
  }
}

/**
 * Decode a JPEG to RGBA. CMYK stored inverted under an Adobe marker is read as the PDF writer's
 * Decode array reads it.
 */
export function decodeJpegPixels(d: Uint8Array): JpegPixels {
  if (d[0] !== 0xff || d[1] !== 0xd8) throw new Error('JPEG: no start marker');
  const quant: Array<Int32Array | null> = [null, null, null, null];
  const dcTables: Array<Huffman | null> = [null, null, null, null];
  const acTables: Array<Huffman | null> = [null, null, null, null];
  let comps: Component[] = [];
  let width = 0;
  let height = 0;
  let progressive = false;
  let restartInterval = 0;
  let adobe: number | null = null;
  let jfif = false;
  let hmax = 1;
  let vmax = 1;
  let mcusX = 0;
  let mcusY = 0;
  let p = 2;
  let done = false;
  while (!done && p + 3 < d.length) {
    if (d[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = d[p + 1];
    if (marker === 0xff) {
      p++;
      continue;
    }
    if (marker === 0xd9) break;
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) {
      p += 2;
      continue;
    }
    const len = (d[p + 2] << 8) | d[p + 3];
    const s = p + 4;
    const e = Math.min(d.length, p + 2 + len);
    if (len < 2) throw new Error('JPEG: bad segment length');
    switch (marker) {
      case 0xdb: // DQT
        for (let q = s; q + 1 <= e; ) {
          const pq = d[q] >> 4;
          const tq = d[q] & 3;
          q++;
          const t = new Int32Array(64);
          for (let k = 0; k < 64; k++, q += pq ? 2 : 1) t[NATURAL[k]] = pq ? (d[q] << 8) | d[q + 1] : d[q];
          quant[tq] = t;
        }
        break;
      case 0xc4: // DHT
        for (let q = s; q + 17 <= e; ) {
          const tc = d[q] >> 4;
          const th = d[q] & 3;
          const counts = d.subarray(q + 1, q + 17);
          let n = 0;
          for (let i = 0; i < 16; i++) n += counts[i];
          const syms = new Uint8Array(256);
          syms.set(d.subarray(q + 17, Math.min(e, q + 17 + n)));
          (tc ? acTables : dcTables)[th] = huffman(counts, syms);
          q += 17 + n;
        }
        break;
      case 0xdd: // DRI
        restartInterval = (d[s] << 8) | d[s + 1];
        break;
      case 0xe0:
        if (d[s] === 0x4a && d[s + 1] === 0x46 && d[s + 2] === 0x49 && d[s + 3] === 0x46 && d[s + 4] === 0) jfif = true;
        break;
      case 0xee:
        if (len >= 14 && d[s] === 0x41 && d[s + 1] === 0x64 && d[s + 2] === 0x6f && d[s + 3] === 0x62 && d[s + 4] === 0x65) adobe = d[s + 11];
        break;
      case 0xc0:
      case 0xc1:
      case 0xc2: {
        if (d[s] !== 8) throw new Error('JPEG: only 8-bit samples are supported');
        progressive = marker === 0xc2;
        height = (d[s + 1] << 8) | d[s + 2];
        width = (d[s + 3] << 8) | d[s + 4];
        const n = d[s + 5];
        if (!width || !height || !n || n > 4) throw new Error('JPEG: bad frame header');
        comps = [];
        for (let i = 0; i < n; i++) {
          const o = s + 6 + i * 3;
          const h = d[o + 1] >> 4;
          const v = d[o + 1] & 15;
          if (h < 1 || h > 4 || v < 1 || v > 4) throw new Error('JPEG: bad sampling factors');
          comps.push({ id: d[o], h, v, tq: d[o + 2] & 3, quant: null, bw: 0, bh: 0, stride: 0, rows: 0, coefs: new Int16Array(0), sw: 0, sh: 0, dc: null, ac: null, pred: 0 });
        }
        hmax = Math.max(...comps.map((c) => c.h));
        vmax = Math.max(...comps.map((c) => c.v));
        mcusX = Math.ceil(width / (8 * hmax));
        mcusY = Math.ceil(height / (8 * vmax));
        for (const c of comps) {
          if (hmax % c.h || vmax % c.v) throw new Error('JPEG: fractional sampling is not supported');
          c.sw = Math.ceil((width * c.h) / hmax);
          c.sh = Math.ceil((height * c.v) / vmax);
          c.bw = Math.ceil(c.sw / 8);
          c.bh = Math.ceil(c.sh / 8);
          c.stride = mcusX * c.h;
          c.rows = mcusY * c.v;
          c.coefs = new Int16Array(c.stride * c.rows * 64);
        }
        break;
      }
      case 0xc3:
      case 0xc5:
      case 0xc6:
      case 0xc7:
      case 0xc9:
      case 0xca:
      case 0xcb:
      case 0xcd:
      case 0xce:
      case 0xcf:
        throw new Error('JPEG: lossless, hierarchical and arithmetic-coded JPEGs are not supported');
      case 0xda: {
        if (!comps.length) throw new Error('JPEG: scan before frame header');
        const ns = d[s];
        const scan: Component[] = [];
        for (let i = 0; i < ns; i++) {
          const c = comps.find((x) => x.id === d[s + 1 + i * 2]);
          if (!c) throw new Error('JPEG: scan names an unknown component');
          c.dc = dcTables[d[s + 2 + i * 2] >> 4];
          c.ac = acTables[d[s + 2 + i * 2] & 3];
          c.quant ??= quant[c.tq];
          scan.push(c);
        }
        const q = s + 1 + ns * 2;
        const ss = d[q];
        const se = d[q + 1];
        const ah = d[q + 2] >> 4;
        const al = d[q + 2] & 15;
        const bits = new Bits(d, e);
        decodeScan(bits, scan, progressive, ss, se, ah, al, restartInterval, mcusX, mcusY);
        p = bits.end();
        continue;
      }
    }
    p += 2 + len;
    if (marker === 0xd9) done = true;
  }
  if (!comps.length) throw new Error('JPEG: no frame header');
  for (const c of comps) if (!c.quant) c.quant = quant[c.tq] ?? new Int32Array(64);
  return output(comps, width, height, hmax, vmax, colorTransform(comps, adobe, jfif), adobe !== null);
}

function decodeScan(bits: Bits, scan: Component[], progressive: boolean, ss: number, se: number, ah: number, al: number, restartInterval: number, mcusX: number, mcusY: number): void {
  for (const c of scan) c.pred = 0;
  let eobrun = 0;
  const single = scan.length === 1;
  const c0 = scan[0];
  // A single component's scan runs over its own blocks; interleaved ones over whole MCUs.
  const unitsX = single ? c0.bw : mcusX;
  const total = single ? c0.bw * c0.bh : mcusX * mcusY;
  const block = (c: Component, row: number, col: number) => {
    const at = (row * c.stride + col) * 64;
    if (bits.short) return;
    const coefs = c.coefs;
    if (!progressive) {
      const t = bits.decode(c.dc!);
      c.pred += t ? bits.receive(t) : 0;
      coefs[at] = c.pred;
      const ac = c.ac!;
      for (let k = 1; k < 64; k++) {
        const rs = bits.decode(ac);
        const s = rs & 15;
        const r = rs >> 4;
        if (s) {
          k += r;
          coefs[at + NATURAL[k]] = bits.receive(s);
        } else {
          if (r !== 15) break;
          k += 15;
        }
      }
    } else if (ss === 0) {
      // DC scans: the first, or one more bit.
      if (ah === 0) {
        const t = bits.decode(c.dc!);
        c.pred += t ? bits.receive(t) : 0;
        coefs[at] = c.pred * (1 << al);
      } else if (bits.bit()) coefs[at] |= 1 << al;
    } else if (ah === 0) {
      if (eobrun > 0) {
        eobrun--;
        return;
      }
      const ac = c.ac!;
      for (let k = ss; k <= se; k++) {
        const rs = bits.decode(ac);
        const s = rs & 15;
        const r = rs >> 4;
        if (s) {
          k += r;
          coefs[at + NATURAL[k]] = bits.receive(s) * (1 << al);
        } else if (r === 15) k += 15;
        else {
          eobrun = (1 << r) - 1;
          if (r) eobrun += bits.bits(r);
          break;
        }
      }
    } else {
      // AC refinement (libjpeg's decode_mcu_AC_refine).
      const p1 = 1 << al;
      const m1 = -1 << al;
      let k = ss;
      if (eobrun === 0) {
        const ac = c.ac!;
        for (; k <= se; k++) {
          const rs = bits.decode(ac);
          let r = rs >> 4;
          let s = rs & 15;
          if (s) s = bits.bit() ? p1 : m1;
          else if (r !== 15) {
            eobrun = 1 << r;
            if (r) eobrun += bits.bits(r);
            break;
          }
          do {
            const pos = at + NATURAL[k];
            if (coefs[pos] !== 0) {
              if (bits.bit() && (coefs[pos] & p1) === 0) coefs[pos] += coefs[pos] >= 0 ? p1 : m1;
            } else if (--r < 0) break;
            k++;
          } while (k <= se);
          if (s) coefs[at + NATURAL[k]] = s;
        }
      }
      if (eobrun > 0) {
        for (; k <= se; k++) {
          const pos = at + NATURAL[k];
          if (coefs[pos] !== 0 && bits.bit() && (coefs[pos] & p1) === 0) coefs[pos] += coefs[pos] >= 0 ? p1 : m1;
        }
        eobrun--;
      }
    }
  };
  for (let n = 0; n < total; n++) {
    if (restartInterval && n && n % restartInterval === 0) {
      bits.restart();
      eobrun = 0;
      for (const c of scan) c.pred = 0;
    }
    const ux = n % unitsX;
    const uy = (n / unitsX) | 0;
    if (single) block(c0, uy, ux);
    else for (const c of scan) for (let v = 0; v < c.v; v++) for (let h = 0; h < c.h; h++) block(c, uy * c.v + v, ux * c.h + h);
  }
}

/** libjpeg's colour space choice, as Poppler's DCT stream makes it: 0 none, 1 YCbCr, 2 YCCK. */
function colorTransform(comps: Component[], adobe: number | null, jfif: boolean): number {
  if (adobe !== null) return comps.length === 4 ? (adobe ? 2 : 0) : adobe ? 1 : 0;
  if (comps.length === 3) return jfif || !(comps[0].id === 82 && comps[1].id === 71 && comps[2].id === 66) ? 1 : 0;
  return 0;
}

/** libjpeg's jpeg_idct_islow: one block's coefficients to 8 × 8 samples at `out[o]`, rows `stride` apart. */
function idct(coefs: Int16Array, at: number, q: Int32Array, out: Uint8Array, o: number, stride: number, ws: Int32Array): void {
  for (let c = 0; c < 8; c++) {
    const i = at + c;
    if (!coefs[i + 8] && !coefs[i + 16] && !coefs[i + 24] && !coefs[i + 32] && !coefs[i + 40] && !coefs[i + 48] && !coefs[i + 56]) {
      const dc = (coefs[i] * q[c]) << 2;
      for (let r = 0; r < 64; r += 8) ws[r + c] = dc;
      continue;
    }
    let z2 = coefs[i + 16] * q[c + 16];
    let z3 = coefs[i + 48] * q[c + 48];
    let z1 = (z2 + z3) * 4433;
    let tmp2 = z1 + z3 * -15137;
    let tmp3 = z1 + z2 * 6270;
    z2 = coefs[i] * q[c];
    z3 = coefs[i + 32] * q[c + 32];
    let tmp0 = (z2 + z3) * 8192;
    let tmp1 = (z2 - z3) * 8192;
    const tmp10 = tmp0 + tmp3;
    const tmp13 = tmp0 - tmp3;
    const tmp11 = tmp1 + tmp2;
    const tmp12 = tmp1 - tmp2;
    tmp0 = coefs[i + 56] * q[c + 56];
    tmp1 = coefs[i + 40] * q[c + 40];
    tmp2 = coefs[i + 24] * q[c + 24];
    tmp3 = coefs[i + 8] * q[c + 8];
    z1 = tmp0 + tmp3;
    z2 = tmp1 + tmp2;
    z3 = tmp0 + tmp2;
    let z4 = tmp1 + tmp3;
    const z5 = (z3 + z4) * 9633;
    tmp0 *= 2446;
    tmp1 *= 16819;
    tmp2 *= 25172;
    tmp3 *= 12299;
    z1 *= -7373;
    z2 *= -20995;
    z3 = z3 * -16069 + z5;
    z4 = z4 * -3196 + z5;
    tmp0 += z1 + z3;
    tmp1 += z2 + z4;
    tmp2 += z2 + z3;
    tmp3 += z1 + z4;
    ws[c] = (tmp10 + tmp3 + 1024) >> 11;
    ws[c + 56] = (tmp10 - tmp3 + 1024) >> 11;
    ws[c + 8] = (tmp11 + tmp2 + 1024) >> 11;
    ws[c + 48] = (tmp11 - tmp2 + 1024) >> 11;
    ws[c + 16] = (tmp12 + tmp1 + 1024) >> 11;
    ws[c + 40] = (tmp12 - tmp1 + 1024) >> 11;
    ws[c + 24] = (tmp13 + tmp0 + 1024) >> 11;
    ws[c + 32] = (tmp13 - tmp0 + 1024) >> 11;
  }
  for (let r = 0; r < 64; r += 8, o += stride) {
    if (!ws[r + 1] && !ws[r + 2] && !ws[r + 3] && !ws[r + 4] && !ws[r + 5] && !ws[r + 6] && !ws[r + 7]) {
      const v = RANGE[((ws[r] + 16) >> 5) & 1023];
      for (let k = 0; k < 8; k++) out[o + k] = v;
      continue;
    }
    let z2 = ws[r + 2];
    let z3 = ws[r + 6];
    let z1 = (z2 + z3) * 4433;
    let tmp2 = z1 + z3 * -15137;
    let tmp3 = z1 + z2 * 6270;
    let tmp0 = (ws[r] + ws[r + 4]) * 8192;
    let tmp1 = (ws[r] - ws[r + 4]) * 8192;
    const tmp10 = tmp0 + tmp3;
    const tmp13 = tmp0 - tmp3;
    const tmp11 = tmp1 + tmp2;
    const tmp12 = tmp1 - tmp2;
    tmp0 = ws[r + 7];
    tmp1 = ws[r + 5];
    tmp2 = ws[r + 3];
    tmp3 = ws[r + 1];
    z1 = tmp0 + tmp3;
    z2 = tmp1 + tmp2;
    z3 = tmp0 + tmp2;
    let z4 = tmp1 + tmp3;
    const z5 = (z3 + z4) * 9633;
    tmp0 *= 2446;
    tmp1 *= 16819;
    tmp2 *= 25172;
    tmp3 *= 12299;
    z1 *= -7373;
    z2 *= -20995;
    z3 = z3 * -16069 + z5;
    z4 = z4 * -3196 + z5;
    tmp0 += z1 + z3;
    tmp1 += z2 + z4;
    tmp2 += z2 + z3;
    tmp3 += z1 + z4;
    out[o] = RANGE[((tmp10 + tmp3 + 131072) >> 18) & 1023];
    out[o + 7] = RANGE[((tmp10 - tmp3 + 131072) >> 18) & 1023];
    out[o + 1] = RANGE[((tmp11 + tmp2 + 131072) >> 18) & 1023];
    out[o + 6] = RANGE[((tmp11 - tmp2 + 131072) >> 18) & 1023];
    out[o + 2] = RANGE[((tmp12 + tmp1 + 131072) >> 18) & 1023];
    out[o + 5] = RANGE[((tmp12 - tmp1 + 131072) >> 18) & 1023];
    out[o + 3] = RANGE[((tmp13 + tmp0 + 131072) >> 18) & 1023];
    out[o + 4] = RANGE[((tmp13 - tmp0 + 131072) >> 18) & 1023];
  }
}

/** A component's samples: every block through the inverse DCT. */
function plane(c: Component): { data: Uint8Array; stride: number } {
  const stride = c.bw * 8;
  const data = new Uint8Array(stride * c.bh * 8);
  const q = c.quant!;
  const ws = new Int32Array(64);
  for (let row = 0; row < c.bh; row++) {
    for (let col = 0; col < c.bw; col++) idct(c.coefs, (row * c.stride + col) * 64, q, data, row * 8 * stride + col * 8, stride, ws);
  }
  return { data, stride };
}

/** One output row of component samples at full size, upsampled as libjpeg's jdsample.c does. */
function upsampleRow(c: Component, src: { data: Uint8Array; stride: number }, hf: number, vf: number, y: number, width: number, out: Uint8Array): void {
  const { data, stride } = src;
  const sw = c.sw;
  const fancy = (hf === 2 && (vf === 1 || vf === 2) && sw > 2) || (hf === 1 && vf === 2);
  if (!fancy) {
    const row = Math.min(c.sh - 1, (y / vf) | 0) * stride;
    if (hf === 1) out.set(data.subarray(row, row + width));
    else for (let x = 0; x < width; x++) out[x] = data[row + ((x / hf) | 0)];
    return;
  }
  const inrow = (y / vf) | 0;
  const near = inrow * stride;
  if (vf === 1) {
    // h2v1: three quarters of the nearer sample and one of the further one.
    let v = data[near];
    out[0] = v;
    out[1] = (v * 3 + data[near + 1] + 2) >> 2;
    let o = 2;
    for (let i = 1; i < sw - 1; i++, o += 2) {
      v = data[near + i] * 3;
      out[o] = (v + data[near + i - 1] + 1) >> 2;
      out[o + 1] = (v + data[near + i + 1] + 2) >> 2;
    }
    v = data[near + sw - 1];
    out[o] = (v * 3 + data[near + sw - 2] + 1) >> 2;
    out[o + 1] = v;
    return;
  }
  // The row above the top and below the bottom repeat the edge rows.
  const other = Math.max(0, Math.min(c.sh - 1, y % 2 === 0 ? inrow - 1 : inrow + 1)) * stride;
  if (hf === 1) {
    const bias = y % 2 === 0 ? 1 : 2;
    for (let x = 0; x < width; x++) out[x] = (data[near + x] * 3 + data[other + x] + bias) >> 2;
    return;
  }
  // h2v2: 9/16, 3/16, 3/16 and 1/16 of the four nearest samples.
  let thiscol = data[near] * 3 + data[other];
  let nextcol = data[near + 1] * 3 + data[other + 1];
  out[0] = (thiscol * 4 + 8) >> 4;
  out[1] = (thiscol * 3 + nextcol + 7) >> 4;
  let lastcol = thiscol;
  thiscol = nextcol;
  let o = 2;
  for (let i = 2; i < sw; i++, o += 2) {
    nextcol = data[near + i] * 3 + data[other + i];
    out[o] = (thiscol * 3 + lastcol + 8) >> 4;
    out[o + 1] = (thiscol * 3 + nextcol + 7) >> 4;
    lastcol = thiscol;
    thiscol = nextcol;
  }
  out[o] = (thiscol * 3 + lastcol + 8) >> 4;
  out[o + 1] = (thiscol * 4 + 7) >> 4;
}

function output(comps: Component[], width: number, height: number, hmax: number, vmax: number, transform: number, inverted: boolean): JpegPixels {
  const n = comps.length;
  const out = new Uint8Array(width * height * 4);
  const rows = comps.map(() => new Uint8Array(Math.max(width, ...comps.map((c) => c.bw * 8)) + 16));
  const planes = comps.map(plane);
  for (let y = 0; y < height; y++) {
    for (let i = 0; i < n; i++) upsampleRow(comps[i], planes[i], hmax / comps[i].h, vmax / comps[i].v, y, width, rows[i]);
    convertRow(rows, n, transform, inverted, out, y * width * 4, width);
  }
  return { width, height, components: n, data: out };
}

function convertRow(rows: Uint8Array[], n: number, transform: number, inverted: boolean, out: Uint8Array, o: number, width: number): void {
  const [a, b, c, k] = rows;
  if (LITTLE_ENDIAN && n !== 4) {
    // Whole opaque pixels at once: red in the low byte.
    const words = new Uint32Array(out.buffer, out.byteOffset + o, width);
    if (n < 3) for (let x = 0; x < width; x++) words[x] = (a[x] * 0x10101) | 0xff000000;
    else if (transform) {
      for (let x = 0; x < width; x++) {
        const y = a[x] + 256;
        const cb = b[x];
        const cr = c[x];
        words[x] = LIMIT[y + CR_R[cr]] | (LIMIT[y + ((CB_G[cb] + CR_G[cr]) >> 16)] << 8) | (LIMIT[y + CB_B[cb]] << 16) | 0xff000000;
      }
    } else for (let x = 0; x < width; x++) words[x] = a[x] | (b[x] << 8) | (c[x] << 16) | 0xff000000;
    return;
  }
  if (n === 1 || n === 2) {
    for (let x = 0; x < width; x++, o += 4) {
      out[o] = out[o + 1] = out[o + 2] = a[x];
      out[o + 3] = 255;
    }
    return;
  }
  if (n === 3) {
    for (let x = 0; x < width; x++, o += 4) {
      const y = a[x] + 256;
      const cb = b[x];
      const cr = c[x];
      out[o] = transform ? LIMIT[y + CR_R[cr]] : a[x];
      out[o + 1] = transform ? LIMIT[y + ((CB_G[cb] + CR_G[cr]) >> 16)] : cb;
      out[o + 2] = transform ? LIMIT[y + CB_B[cb]] : cr;
      out[o + 3] = 255;
    }
    return;
  }
  for (let x = 0; x < width; x++, o += 4) {
    let c0 = a[x];
    let c1 = b[x];
    let c2 = c[x];
    if (transform) {
      // YCCK: YCbCr to RGB, inverted to CMY; K passes through.
      const y = c0 + 256;
      c0 = 255 - LIMIT[y + CR_R[c2]];
      c1 = 255 - LIMIT[y + ((CB_G[c1] + CR_G[c2]) >> 16)];
      c2 = 255 - LIMIT[y + CB_B[b[x]]];
    }
    cmykToRgb(c0, c1, c2, k[x], inverted, out, o);
  }
}

/**
 * Poppler's DeviceCMYK to RGB (GfxDeviceCMYKColorSpace::getRGBLine), after the Decode array the
 * PDF writer gives Adobe's inverted CMYK; RGB bytes at `out[o]`.
 */
function cmykToRgb(cb: number, mb: number, yb: number, kb: number, inverted: boolean, out: Uint8Array, o: number): void {
  const c = (inverted ? 255 - cb : cb) / 255;
  const m = (inverted ? 255 - mb : mb) / 255;
  const y = (inverted ? 255 - yb : yb) / 255;
  const k = (inverted ? 255 - kb : kb) / 255;
  const c1 = 1 - c;
  const m1 = 1 - m;
  const y1 = 1 - y;
  const k1 = 1 - k;
  let x = c1 * m1 * y1 * k1;
  let r = x;
  let g = x;
  let b = x;
  x = c1 * m1 * y1 * k;
  r += 0.1373 * x;
  g += 0.1216 * x;
  b += 0.1255 * x;
  x = c1 * m1 * y * k1;
  r += x;
  g += 0.949 * x;
  x = c1 * m1 * y * k;
  r += 0.1098 * x;
  g += 0.102 * x;
  x = c1 * m * y1 * k1;
  r += 0.9255 * x;
  b += 0.549 * x;
  x = c1 * m * y1 * k;
  r += 0.1412 * x;
  x = c1 * m * y * k1;
  r += 0.9294 * x;
  g += 0.1098 * x;
  b += 0.1412 * x;
  x = c1 * m * y * k;
  r += 0.1333 * x;
  x = c * m1 * y1 * k1;
  g += 0.6784 * x;
  b += 0.9373 * x;
  x = c * m1 * y1 * k;
  g += 0.0588 * x;
  b += 0.1412 * x;
  x = c * m1 * y * k1;
  g += 0.651 * x;
  b += 0.3137 * x;
  x = c * m1 * y * k;
  g += 0.0745 * x;
  x = c * m * y1 * k1;
  r += 0.1804 * x;
  g += 0.1922 * x;
  b += 0.5725 * x;
  x = c * m * y1 * k;
  b += 0.0078 * x;
  x = c * m * y * k1;
  r += 0.2118 * x;
  g += 0.2119 * x;
  b += 0.2235 * x;
  out[o] = (Math.min(1, Math.max(0, r)) * 255) | 0;
  out[o + 1] = (Math.min(1, Math.max(0, g)) * 255) | 0;
  out[o + 2] = (Math.min(1, Math.max(0, b)) * 255) | 0;
  out[o + 3] = 255;
}
