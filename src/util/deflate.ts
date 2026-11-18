/**
 * DEFLATE compressor (LZ77 with hash chains + dynamic Huffman blocks), zlib-wrapped for PDF
 * FlateDecode streams. Favors simplicity over speed; output is valid for any conforming inflater.
 */

const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193,
  12289, 16385, 24577,
];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

const LEN_CODE = new Uint16Array(259);
const DIST_CODE_LOW = new Uint8Array(512);
(() => {
  for (let i = 0; i < 29; i++) {
    const end = i === 28 ? 259 : LEN_BASE[i + 1];
    for (let l = LEN_BASE[i]; l < end && l < 259; l++) LEN_CODE[l] = i;
  }
  LEN_CODE[258] = 28;
  for (let i = 0; i < 30; i++) {
    for (let d = DIST_BASE[i]; d < (i === 29 ? 32769 : DIST_BASE[i + 1]); d++) {
      if (d <= 512) DIST_CODE_LOW[d - 1] = i;
    }
  }
})();

function distCode(d: number): number {
  if (d <= 512) return DIST_CODE_LOW[d - 1];
  // Binary search for larger distances.
  let lo = 0;
  let hi = 29;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (DIST_BASE[mid] <= d) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

class BitWriter {
  buf = new Uint8Array(1 << 16);
  len = 0;
  bitBuf = 0;
  bitCnt = 0;

  private grow(extra: number): void {
    if (this.len + extra <= this.buf.length) return;
    let n = this.buf.length * 2;
    while (n < this.len + extra) n *= 2;
    const nb = new Uint8Array(n);
    nb.set(this.buf.subarray(0, this.len));
    this.buf = nb;
  }

  write(value: number, n: number): void {
    // value fits in n <= 16 bits
    this.bitBuf |= value << this.bitCnt;
    this.bitCnt += n;
    while (this.bitCnt >= 8) {
      this.grow(1);
      this.buf[this.len++] = this.bitBuf & 255;
      this.bitBuf >>>= 8;
      this.bitCnt -= 8;
    }
  }

  flushByte(): void {
    if (this.bitCnt > 0) {
      this.grow(1);
      this.buf[this.len++] = this.bitBuf & 255;
      this.bitBuf = 0;
      this.bitCnt = 0;
    }
  }

  bytes(data: Uint8Array): void {
    this.grow(data.length);
    this.buf.set(data, this.len);
    this.len += data.length;
  }
}

/** Build length-limited Huffman code lengths for the given frequencies. */
function codeLengths(freqs: ArrayLike<number>, maxBits: number): Uint8Array {
  const n = freqs.length;
  const lengths = new Uint8Array(n);
  const used: number[] = [];
  for (let i = 0; i < n; i++) if (freqs[i] > 0) used.push(i);
  if (used.length === 0) return lengths;
  if (used.length === 1) {
    lengths[used[0]] = 1;
    return lengths;
  }
  // Standard Huffman via a simple priority queue (array sort; n <= 288).
  type Node = { w: number; sym: number; l?: Node; r?: Node };
  let nodes: Node[] = used.map((s) => ({ w: freqs[s], sym: s }));
  while (nodes.length > 1) {
    nodes.sort((a, b) => a.w - b.w);
    const a = nodes.shift()!;
    const b = nodes.shift()!;
    nodes.push({ w: a.w + b.w, sym: -1, l: a, r: b });
  }
  const depth = new Map<number, number>();
  const walk = (node: Node, d: number) => {
    if (node.sym >= 0) depth.set(node.sym, Math.max(d, 1));
    else {
      walk(node.l!, d + 1);
      walk(node.r!, d + 1);
    }
  };
  walk(nodes[0], 0);
  // Count lengths, then limit using the JPEG Annex K adjustment.
  const blCount = new Array<number>(n + 2).fill(0);
  for (const d of depth.values()) blCount[d]++;
  for (let i = n + 1; i > maxBits; i--) {
    while (blCount[i] > 0) {
      let j = i - 2;
      while (j > 0 && blCount[j] === 0) j--;
      blCount[i] -= 2;
      blCount[i - 1]++;
      blCount[j + 1] += 2;
      blCount[j]--;
    }
  }
  // Assign the shortest lengths to the most frequent symbols.
  const sorted = used.slice().sort((a, b) => freqs[b] - freqs[a] || a - b);
  let k = 0;
  for (let len = 1; len <= maxBits; len++) {
    for (let c = 0; c < blCount[len]; c++) lengths[sorted[k++]] = len;
  }
  return lengths;
}

function canonicalCodes(lengths: Uint8Array): Uint16Array {
  const codes = new Uint16Array(lengths.length);
  const blCount = new Uint16Array(16);
  for (const l of lengths) if (l) blCount[l]++;
  const next = new Uint16Array(16);
  let code = 0;
  for (let bits = 1; bits < 16; bits++) {
    code = (code + blCount[bits - 1]) << 1;
    next[bits] = code;
  }
  for (let i = 0; i < lengths.length; i++) {
    const l = lengths[i];
    if (l) {
      // Reverse for LSB-first emission.
      let c = next[l]++;
      let r = 0;
      for (let b = 0; b < l; b++) {
        r = (r << 1) | (c & 1);
        c >>>= 1;
      }
      codes[i] = r;
    }
  }
  return codes;
}

/** Symbols collected for one block: literals (0..255), or length/distance pairs. */
class SymbolBuffer {
  // Each entry: lit/len symbol value in low 16 bits; for matches we store length and distance separately.
  kind: Uint8Array;
  a: Uint16Array;
  b: Uint16Array;
  n = 0;
  constructor(cap: number) {
    this.kind = new Uint8Array(cap);
    this.a = new Uint16Array(cap);
    this.b = new Uint16Array(cap);
  }
}

function writeBlock(bw: BitWriter, syms: SymbolBuffer, final: boolean): void {
  const litFreq = new Uint32Array(286);
  const distFreq = new Uint32Array(30);
  for (let i = 0; i < syms.n; i++) {
    if (syms.kind[i] === 0) litFreq[syms.a[i]]++;
    else {
      litFreq[257 + LEN_CODE[syms.a[i]]]++;
      distFreq[distCode(syms.b[i])]++;
    }
  }
  litFreq[256] = 1;
  let distUsed = false;
  for (const f of distFreq) if (f) distUsed = true;
  if (!distUsed) distFreq[0] = 1;

  const litLens = codeLengths(litFreq, 15);
  const distLens = codeLengths(distFreq, 15);
  let hlit = 286;
  while (hlit > 257 && litLens[hlit - 1] === 0) hlit--;
  let hdist = 30;
  while (hdist > 1 && distLens[hdist - 1] === 0) hdist--;

  // Run-length encode the code lengths.
  const all = new Uint8Array(hlit + hdist);
  all.set(litLens.subarray(0, hlit), 0);
  all.set(distLens.subarray(0, hdist), hlit);
  const rle: Array<[number, number]> = [];
  for (let i = 0; i < all.length; ) {
    const v = all[i];
    let run = 1;
    while (i + run < all.length && all[i + run] === v) run++;
    let left = run;
    if (v === 0) {
      while (left >= 11) {
        const r = Math.min(left, 138);
        rle.push([18, r - 11]);
        left -= r;
      }
      if (left >= 3) {
        rle.push([17, left - 3]);
        left = 0;
      }
    } else {
      rle.push([v, 0]);
      left--;
      while (left >= 3) {
        const r = Math.min(left, 6);
        rle.push([16, r - 3]);
        left -= r;
      }
    }
    while (left-- > 0) rle.push([v, 0]);
    i += run;
  }
  const clFreq = new Uint32Array(19);
  for (const [s] of rle) clFreq[s]++;
  const clLens = codeLengths(clFreq, 7);
  const clCodes = canonicalCodes(clLens);
  let hclen = 19;
  while (hclen > 4 && clLens[CL_ORDER[hclen - 1]] === 0) hclen--;

  bw.write(final ? 1 : 0, 1);
  bw.write(2, 2);
  bw.write(hlit - 257, 5);
  bw.write(hdist - 1, 5);
  bw.write(hclen - 4, 4);
  for (let i = 0; i < hclen; i++) bw.write(clLens[CL_ORDER[i]], 3);
  for (const [s, extra] of rle) {
    bw.write(clCodes[s], clLens[s]);
    if (s === 16) bw.write(extra, 2);
    else if (s === 17) bw.write(extra, 3);
    else if (s === 18) bw.write(extra, 7);
  }

  const litCodes = canonicalCodes(litLens);
  const distCodes = canonicalCodes(distLens);
  for (let i = 0; i < syms.n; i++) {
    if (syms.kind[i] === 0) {
      const s = syms.a[i];
      bw.write(litCodes[s], litLens[s]);
    } else {
      const len = syms.a[i];
      const lc = LEN_CODE[len];
      bw.write(litCodes[257 + lc], litLens[257 + lc]);
      if (LEN_EXTRA[lc]) bw.write(len - LEN_BASE[lc], LEN_EXTRA[lc]);
      const d = syms.b[i];
      const dc = distCode(d);
      bw.write(distCodes[dc], distLens[dc]);
      if (DIST_EXTRA[dc]) bw.write(d - DIST_BASE[dc], DIST_EXTRA[dc]);
    }
  }
  bw.write(litCodes[256], litLens[256]);
}

const WINDOW = 32768;
const HASH_BITS = 15;
const HASH_SIZE = 1 << HASH_BITS;
const MAX_CHAIN = 48;
const BLOCK_SYMBOLS = 1 << 15;

/** Raw DEFLATE compression. */
export function deflateRaw(data: Uint8Array): Uint8Array {
  const bw = new BitWriter();
  if (data.length === 0) {
    bw.write(1, 1);
    bw.write(1, 2);
    bw.write(0, 7); // end-of-block in fixed Huffman (code 0000000)
    bw.flushByte();
    return bw.buf.slice(0, bw.len);
  }
  const head = new Int32Array(HASH_SIZE).fill(-1);
  const prev = new Int32Array(WINDOW).fill(-1);
  const syms = new SymbolBuffer(BLOCK_SYMBOLS + 2);
  const n = data.length;
  const hash = (i: number) => ((data[i] << 10) ^ (data[i + 1] << 5) ^ data[i + 2]) & (HASH_SIZE - 1);
  const insert = (i: number) => {
    if (i + 2 >= n) return;
    const h = hash(i);
    prev[i & (WINDOW - 1)] = head[h];
    head[h] = i;
  };
  const findMatch = (i: number): [number, number] => {
    if (i + 2 >= n) return [0, 0];
    let best = 0;
    let bestDist = 0;
    let cand = head[hash(i)];
    let chain = MAX_CHAIN;
    const maxLen = Math.min(258, n - i);
    while (cand >= 0 && i - cand <= WINDOW && chain-- > 0) {
      if (data[cand + best] === data[i + best]) {
        let l = 0;
        while (l < maxLen && data[cand + l] === data[i + l]) l++;
        if (l > best) {
          best = l;
          bestDist = i - cand;
          if (l === maxLen) break;
        }
      }
      const nextCand = prev[cand & (WINDOW - 1)];
      if (nextCand >= cand) break;
      cand = nextCand;
    }
    return best >= 3 ? [best, bestDist] : [0, 0];
  };

  let i = 0;
  while (i < n) {
    let [len, dist] = findMatch(i);
    if (len > 0 && len < 32 && i + 1 < n) {
      // Lazy matching: prefer a longer match starting at the next byte.
      insert(i);
      const [len2, dist2] = findMatch(i + 1);
      if (len2 > len) {
        syms.kind[syms.n] = 0;
        syms.a[syms.n++] = data[i];
        i++;
        insert(i);
        len = len2;
        dist = dist2;
      }
      if (len >= 3) {
        syms.kind[syms.n] = 1;
        syms.a[syms.n] = len;
        syms.b[syms.n++] = dist;
        for (let k = 1; k < len; k++) insert(i + k);
        i += len;
      }
    } else if (len >= 3) {
      syms.kind[syms.n] = 1;
      syms.a[syms.n] = len;
      syms.b[syms.n++] = dist;
      for (let k = 0; k < len; k++) insert(i + k);
      i += len;
    } else {
      insert(i);
      syms.kind[syms.n] = 0;
      syms.a[syms.n++] = data[i];
      i++;
    }
    if (syms.n >= BLOCK_SYMBOLS) {
      writeBlock(bw, syms, false);
      syms.n = 0;
    }
  }
  writeBlock(bw, syms, true);
  bw.flushByte();
  return bw.buf.slice(0, bw.len);
}

function adler32(data: Uint8Array): number {
  let a = 1;
  let b = 0;
  const MOD = 65521;
  for (let i = 0; i < data.length; ) {
    const end = Math.min(i + 3800, data.length);
    for (; i < end; i++) {
      a += data[i];
      b += a;
    }
    a %= MOD;
    b %= MOD;
  }
  return ((b << 16) | a) >>> 0;
}

/** zlib-wrapped DEFLATE, as used by PDF FlateDecode and PNG. */
export function deflateZlib(data: Uint8Array): Uint8Array {
  const body = deflateRaw(data);
  const out = new Uint8Array(body.length + 6);
  out[0] = 0x78;
  out[1] = 0x9c;
  out.set(body, 2);
  const ad = adler32(data);
  out[out.length - 4] = (ad >>> 24) & 255;
  out[out.length - 3] = (ad >>> 16) & 255;
  out[out.length - 2] = (ad >>> 8) & 255;
  out[out.length - 1] = ad & 255;
  return out;
}
