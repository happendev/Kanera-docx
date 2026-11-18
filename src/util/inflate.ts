/**
 * RFC 1951 (raw DEFLATE) decoder, plus a zlib (RFC 1950) wrapper.
 *
 * Self-contained so the library has no runtime dependencies and stays synchronous in both Node and
 * the browser. Decoding uses lookup tables keyed by the next 9 bits, falling back to canonical
 * bit-by-bit decoding for longer codes.
 */

const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193,
  12289, 16385, 24577,
];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

const FAST_BITS = 9;

class Huffman {
  counts = new Uint16Array(16);
  symbols: Uint16Array;
  /** fast[bits] = (symbol << 4) | length, or 0 when the code is longer than FAST_BITS. */
  fast = new Uint32Array(1 << FAST_BITS);

  constructor(lengths: ArrayLike<number>, n: number) {
    this.symbols = new Uint16Array(n);
    for (let i = 0; i < n; i++) this.counts[lengths[i]]++;
    this.counts[0] = 0;
    const offs = new Uint16Array(16);
    for (let i = 1; i < 16; i++) offs[i] = offs[i - 1] + this.counts[i - 1];
    for (let i = 0; i < n; i++) if (lengths[i]) this.symbols[offs[lengths[i]]++] = i;

    // Build the fast table from canonical codes (codes are stored MSB-first, read LSB-first).
    let code = 0;
    let k = 0;
    for (let len = 1; len <= 15; len++) {
      for (let j = 0; j < this.counts[len]; j++, k++, code++) {
        if (len <= FAST_BITS) {
          const rev = reverse(code, len);
          const sym = this.symbols[k];
          for (let fill = rev; fill < 1 << FAST_BITS; fill += 1 << len) this.fast[fill] = (sym << 4) | len;
        }
      }
      code <<= 1;
    }
  }
}

function reverse(code: number, len: number): number {
  let r = 0;
  for (let i = 0; i < len; i++) {
    r = (r << 1) | (code & 1);
    code >>>= 1;
  }
  return r;
}

let FIXED_LIT: Huffman | null = null;
let FIXED_DIST: Huffman | null = null;

function fixedTables(): [Huffman, Huffman] {
  if (!FIXED_LIT || !FIXED_DIST) {
    const l = new Uint8Array(288);
    for (let i = 0; i < 144; i++) l[i] = 8;
    for (let i = 144; i < 256; i++) l[i] = 9;
    for (let i = 256; i < 280; i++) l[i] = 7;
    for (let i = 280; i < 288; i++) l[i] = 8;
    FIXED_LIT = new Huffman(l, 288);
    const d = new Uint8Array(30).fill(5);
    FIXED_DIST = new Huffman(d, 30);
  }
  return [FIXED_LIT, FIXED_DIST];
}

class BitReader {
  pos: number;
  bitBuf = 0;
  bitCnt = 0;
  constructor(
    public data: Uint8Array,
    start: number,
  ) {
    this.pos = start;
  }

  need(n: number): void {
    while (this.bitCnt < n) {
      const byte = this.pos < this.data.length ? this.data[this.pos] : 0;
      this.pos++;
      this.bitBuf |= byte << this.bitCnt;
      this.bitCnt += 8;
    }
  }

  bits(n: number): number {
    if (n === 0) return 0;
    this.need(n);
    const v = this.bitBuf & ((1 << n) - 1);
    this.bitBuf >>>= n;
    this.bitCnt -= n;
    return v;
  }

  decode(h: Huffman): number {
    this.need(FAST_BITS);
    const entry = h.fast[this.bitBuf & ((1 << FAST_BITS) - 1)];
    if (entry) {
      const len = entry & 15;
      this.bitBuf >>>= len;
      this.bitCnt -= len;
      return entry >>> 4;
    }
    // Slow canonical decode for long codes.
    let code = 0;
    let first = 0;
    let index = 0;
    for (let len = 1; len <= 15; len++) {
      code |= this.bits(1);
      const count = h.counts[len];
      if (code - count < first) return h.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new Error('inflate: invalid Huffman code');
  }

  alignToByte(): void {
    const drop = this.bitCnt & 7;
    this.bitBuf >>>= drop;
    this.bitCnt -= drop;
  }
}

class Output {
  buf: Uint8Array;
  len = 0;
  constructor(size: number) {
    this.buf = new Uint8Array(Math.max(size, 1024));
  }
  ensure(extra: number): void {
    if (this.len + extra <= this.buf.length) return;
    let n = this.buf.length * 2;
    while (n < this.len + extra) n *= 2;
    const nb = new Uint8Array(n);
    nb.set(this.buf.subarray(0, this.len));
    this.buf = nb;
  }
}

/** Decode raw DEFLATE data. `sizeHint` preallocates the output when the size is known. */
export function inflateRaw(data: Uint8Array, sizeHint = 0, start = 0): Uint8Array {
  const br = new BitReader(data, start);
  const out = new Output(sizeHint || data.length * 4);
  let final = 0;
  do {
    final = br.bits(1);
    const type = br.bits(2);
    if (type === 0) {
      br.alignToByte();
      // Remaining whole bytes in the bit buffer come first.
      const header: number[] = [];
      while (br.bitCnt >= 8 && header.length < 4) {
        header.push(br.bitBuf & 255);
        br.bitBuf >>>= 8;
        br.bitCnt -= 8;
      }
      while (header.length < 4) header.push(data[br.pos++]);
      const len = header[0] | (header[1] << 8);
      out.ensure(len);
      // Any bytes still buffered belong to the stored payload.
      let copied = 0;
      while (br.bitCnt >= 8 && copied < len) {
        out.buf[out.len++] = br.bitBuf & 255;
        br.bitBuf >>>= 8;
        br.bitCnt -= 8;
        copied++;
      }
      const rest = len - copied;
      out.buf.set(data.subarray(br.pos, br.pos + rest), out.len);
      out.len += rest;
      br.pos += rest;
    } else if (type === 1 || type === 2) {
      let lit: Huffman;
      let dist: Huffman;
      if (type === 1) {
        [lit, dist] = fixedTables();
      } else {
        const hlit = br.bits(5) + 257;
        const hdist = br.bits(5) + 1;
        const hclen = br.bits(4) + 4;
        const clLens = new Uint8Array(19);
        for (let i = 0; i < hclen; i++) clLens[CL_ORDER[i]] = br.bits(3);
        const cl = new Huffman(clLens, 19);
        const lens = new Uint8Array(hlit + hdist);
        let i = 0;
        while (i < hlit + hdist) {
          const sym = br.decode(cl);
          if (sym < 16) lens[i++] = sym;
          else if (sym === 16) {
            if (i === 0) throw new Error('inflate: repeat with no previous length');
            const prev = lens[i - 1];
            let r = 3 + br.bits(2);
            while (r--) lens[i++] = prev;
          } else if (sym === 17) {
            let r = 3 + br.bits(3);
            while (r--) lens[i++] = 0;
          } else {
            let r = 11 + br.bits(7);
            while (r--) lens[i++] = 0;
          }
        }
        lit = new Huffman(lens.subarray(0, hlit), hlit);
        dist = new Huffman(lens.subarray(hlit), hdist);
      }
      for (;;) {
        const sym = br.decode(lit);
        if (sym < 256) {
          out.ensure(1);
          out.buf[out.len++] = sym;
        } else if (sym === 256) {
          break;
        } else {
          const li = sym - 257;
          if (li >= 29) throw new Error('inflate: bad length symbol');
          const length = LEN_BASE[li] + br.bits(LEN_EXTRA[li]);
          const di = br.decode(dist);
          if (di >= 30) throw new Error('inflate: bad distance symbol');
          const distance = DIST_BASE[di] + br.bits(DIST_EXTRA[di]);
          if (distance > out.len) throw new Error('inflate: distance too far back');
          out.ensure(length);
          const b = out.buf;
          let src = out.len - distance;
          for (let k = 0; k < length; k++) b[out.len++] = b[src++];
        }
      }
    } else {
      throw new Error('inflate: invalid block type');
    }
  } while (!final);
  return out.buf.subarray(0, out.len);
}

/** Decode zlib-wrapped data (2-byte header, DEFLATE body, Adler-32 trailer which is not verified). */
export function inflateZlib(data: Uint8Array, sizeHint = 0): Uint8Array {
  if (data.length < 2) throw new Error('zlib: truncated');
  const cmf = data[0];
  const flg = data[1];
  if ((cmf & 15) !== 8 || ((cmf << 8) | flg) % 31 !== 0) throw new Error('zlib: bad header');
  const start = flg & 0x20 ? 6 : 2;
  return inflateRaw(data, sizeHint, start);
}
