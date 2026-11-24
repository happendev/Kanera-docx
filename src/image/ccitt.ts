// CCITT fax decoding (ITU-T T.4 and T.6) to packed 1-bit rows, for TIFF compressions 2, 3 and 4.

/** How the rows are coded: Modified Huffman (TIFF compression 2), T.4 (3) or T.6 (4). */
export type FaxCoding = 'mh' | 't4' | 't6';

export interface FaxOptions {
  coding: FaxCoding;
  /** T.4 only: rows may be coded two-dimensionally, each EOL followed by a tag bit. */
  twoDimensional?: boolean;
}

/**
 * Decode `rows` rows of `columns` pixels to packed rows, most significant bit first, each padded to
 * a whole byte: 0 for pixels coded white, 1 for black. Throws on malformed data.
 */
export function decodeFax(src: Uint8Array, columns: number, rows: number, { coding, twoDimensional = false }: FaxOptions): Uint8Array {
  const rowBytes = Math.ceil(columns / 8);
  const out = new Uint8Array(rowBytes * rows);
  const bits = new BitReader(src);
  // Changing elements: the positions where the colour flips, starting white, then sentinels.
  let ref = new Int32Array(columns + 4);
  let cur = new Int32Array(columns + 4);
  let refLen = 0;
  ref.fill(columns, 0, 4);
  for (let y = 0; y < rows; y++) {
    let twoD = coding === 't6';
    if (coding === 'mh') {
      if (y > 0) bits.align();
    } else if (coding === 't4') {
      // Fill bits (zeros) may pad the EOL that starts each row.
      while (bits.peek(12) === 0 && !bits.done()) bits.skip(1);
      if (bits.peek(12) === 1) {
        bits.skip(12);
        if (twoDimensional) twoD = bits.read(1) === 0;
      }
    }
    if (coding === 't6' && bits.peek(24) === 0x001001) break; // EOFB: the remaining rows stay white.
    const curLen = twoD ? decode2D(bits, ref, refLen, cur, columns) : decode1D(bits, cur, columns);
    cur.fill(columns, curLen, curLen + 4);
    fillRow(out, y * rowBytes, cur, curLen);
    [ref, cur] = [cur, ref];
    refLen = curLen;
  }
  return out;
}

/** One row coded as alternating white and black runs. */
function decode1D(bits: BitReader, cur: Int32Array, columns: number): number {
  let len = 0;
  let pos = 0;
  let black = false;
  while (pos < columns) {
    pos += readRun(bits, black);
    if (pos > columns) throw new Error('CCITT: run past the end of the row');
    if (pos < columns) len = addChange(cur, len, pos);
    black = !black;
  }
  return len;
}

/** One row coded against the row above (`ref`), with pass, horizontal and vertical modes. */
function decode2D(bits: BitReader, ref: Int32Array, refLen: number, cur: Int32Array, columns: number): number {
  let len = 0;
  let a0 = -1;
  let black = false;
  let bi = 0;
  while (a0 < columns) {
    // b1: the first change on the row above, right of a0, to the colour opposite a0's.
    bi = Math.max(0, bi - 2);
    while (bi < refLen && ref[bi] <= a0) bi++;
    if ((bi & 1) !== (black ? 1 : 0)) bi++;
    const b1 = bi < refLen ? ref[bi] : columns;
    const b2 = bi + 1 < refLen ? ref[bi + 1] : columns;
    const mode = MODES[bits.peek(7)];
    if (!mode) throw new Error('CCITT: invalid mode code');
    bits.skip(mode >> 8);
    const m = (mode & 0xff) - 4;
    if (m === 4) {
      // Pass: a0 moves under b2 with its colour unchanged.
      a0 = b2;
    } else if (m === 5) {
      // Horizontal: two runs, the first in a0's colour.
      const start = Math.max(a0, 0);
      const a1 = start + readRun(bits, black);
      const a2 = a1 + readRun(bits, !black);
      if (a2 > columns) throw new Error('CCITT: run past the end of the row');
      if (a1 < columns) len = addChange(cur, len, a1);
      if (a2 < columns) len = addChange(cur, len, a2);
      a0 = a2;
    } else {
      // Vertical: a1 lies within three pixels of b1.
      const a1 = b1 + m;
      if (a1 < 0 || a1 > columns || a1 < a0) throw new Error('CCITT: vertical change out of range');
      if (a1 < columns) len = addChange(cur, len, a1);
      a0 = a1;
      black = !black;
    }
  }
  return len;
}

/** Record a colour change; one at the same place as the last cancels it. */
function addChange(cur: Int32Array, len: number, pos: number): number {
  if (len > 0 && cur[len - 1] === pos) return len - 1;
  if (len > 0 && cur[len - 1] > pos) throw new Error('CCITT: changes out of order');
  cur[len] = pos;
  return len + 1;
}

function fillRow(out: Uint8Array, at: number, cur: Int32Array, len: number): void {
  for (let i = 0; i < len; i += 2) {
    const from = cur[i];
    const to = i + 1 < len ? cur[i + 1] : cur[len];
    for (let x = from; x < to; x++) out[at + (x >> 3)] |= 0x80 >> (x & 7);
  }
}

/** A run length: make-up codes of 64 or more, then a terminating code under 64. */
function readRun(bits: BitReader, black: boolean): number {
  const table = black ? BLACK : WHITE;
  let total = 0;
  for (;;) {
    const entry = table[bits.peek(13)];
    if (entry < 0) throw new Error('CCITT: invalid run code');
    bits.skip(entry >>> 16);
    const run = entry & 0xffff;
    total += run;
    if (run < 64) return total;
  }
}

class BitReader {
  private pos = 0;
  private readonly total: number;
  constructor(private readonly d: Uint8Array) {
    this.total = d.length * 8;
  }
  done(): boolean {
    return this.pos >= this.total;
  }
  /** The next n (up to 24) bits, zeros past the end. */
  peek(n: number): number {
    const at = this.pos >> 3;
    const d = this.d;
    const word = ((d[at] ?? 0) << 24 | (d[at + 1] ?? 0) << 16 | (d[at + 2] ?? 0) << 8 | (d[at + 3] ?? 0)) >>> 0;
    return ((word << (this.pos & 7)) >>> 0) >>> (32 - n);
  }
  skip(n: number): void {
    this.pos += n;
    if (this.pos > this.total + 32) throw new Error('CCITT: data ends early');
  }
  read(n: number): number {
    const v = this.peek(n);
    this.skip(n);
    return v;
  }
  align(): void {
    this.pos = (this.pos + 7) & ~7;
  }
}

/** For each 13-bit window, the run whose code starts it, as length << 16 | run, or -1. */
function runTable(codes: string): Int32Array {
  const table = new Int32Array(1 << 13).fill(-1);
  for (const entry of codes.trim().split(/\s+/)) {
    const [code, run] = entry.split(':');
    const shift = 13 - code.length;
    const first = parseInt(code, 2) << shift;
    table.fill((code.length << 16) | Number(run), first, first + (1 << shift));
  }
  return table;
}

// Make-up codes from 1792 to 2560, shared by both colours.
const EXTENDED = `
00000001000:1792 00000001100:1856 00000001101:1920 000000010010:1984 000000010011:2048 000000010100:2112
000000010101:2176 000000010110:2240 000000010111:2304 000000011100:2368 000000011101:2432 000000011110:2496
000000011111:2560`;

const WHITE = runTable(`
00110101:0 000111:1 0111:2 1000:3 1011:4 1100:5 1110:6 1111:7 10011:8 10100:9 00111:10 01000:11 001000:12
000011:13 110100:14 110101:15 101010:16 101011:17 0100111:18 0001100:19 0001000:20 0010111:21 0000011:22
0000100:23 0101000:24 0101011:25 0010011:26 0100100:27 0011000:28 00000010:29 00000011:30 00011010:31
00011011:32 00010010:33 00010011:34 00010100:35 00010101:36 00010110:37 00010111:38 00101000:39 00101001:40
00101010:41 00101011:42 00101100:43 00101101:44 00000100:45 00000101:46 00001010:47 00001011:48 01010010:49
01010011:50 01010100:51 01010101:52 00100100:53 00100101:54 01011000:55 01011001:56 01011010:57 01011011:58
01001010:59 01001011:60 00110010:61 00110011:62 00110100:63
11011:64 10010:128 010111:192 0110111:256 00110110:320 00110111:384 01100100:448 01100101:512 01101000:576
01100111:640 011001100:704 011001101:768 011010010:832 011010011:896 011010100:960 011010101:1024
011010110:1088 011010111:1152 011011000:1216 011011001:1280 011011010:1344 011011011:1408 010011000:1472
010011001:1536 010011010:1600 011000:1664 010011011:1728 ${EXTENDED}`);

const BLACK = runTable(`
0000110111:0 010:1 11:2 10:3 011:4 0011:5 0010:6 00011:7 000101:8 000100:9 0000100:10 0000101:11 0000111:12
00000100:13 00000111:14 000011000:15 0000010111:16 0000011000:17 0000001000:18 00001100111:19 00001101000:20
00001101100:21 00000110111:22 00000101000:23 00000010111:24 00000011000:25 000011001010:26 000011001011:27
000011001100:28 000011001101:29 000001101000:30 000001101001:31 000001101010:32 000001101011:33
000011010010:34 000011010011:35 000011010100:36 000011010101:37 000011010110:38 000011010111:39
000001101100:40 000001101101:41 000011011010:42 000011011011:43 000001010100:44 000001010101:45
000001010110:46 000001010111:47 000001100100:48 000001100101:49 000001010010:50 000001010011:51
000000100100:52 000000110111:53 000000111000:54 000000100111:55 000000101000:56 000001011000:57
000001011001:58 000000101011:59 000000101100:60 000001011010:61 000001100110:62 000001100111:63
0000001111:64 000011001000:128 000011001001:192 000001011011:256 000000110011:320 000000110100:384
000000110101:448 0000001101100:512 0000001101101:576 0000001001010:640 0000001001011:704 0000001001100:768
0000001001101:832 0000001110010:896 0000001110011:960 0000001110100:1024 0000001110101:1088
0000001110110:1152 0000001110111:1216 0000001010010:1280 0000001010011:1344 0000001010100:1408
0000001010101:1472 0000001011010:1536 0000001011011:1600 0000001100100:1664 0000001100101:1728 ${EXTENDED}`);

/**
 * Two-dimensional mode codes by their first 7 bits: length << 8 | (offset + 4), where offsets −3
 * to 3 are vertical modes, 4 is pass and 5 is horizontal. 0 where no mode code starts.
 */
const MODES = (() => {
  const table = new Uint16Array(128);
  const codes: [string, number][] = [['1', 0], ['011', 1], ['000011', 2], ['0000011', 3], ['010', -1], ['000010', -2], ['0000010', -3], ['0001', 4], ['001', 5]];
  for (const [code, m] of codes) {
    const shift = 7 - code.length;
    const first = parseInt(code, 2) << shift;
    table.fill((code.length << 8) | (m + 4), first, first + (1 << shift));
  }
  return table;
})();
