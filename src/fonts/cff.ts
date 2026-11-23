// Outlines of CFF and CFF2 fonts (OpenType with PostScript outlines): Type 2 charstrings run to
// cubic curves in font units. Hints are skipped, CFF2 variation deltas leave the default instance.

import { OutlineBuilder, type Outline } from './outline.js';
import type { SfntFont } from './sfnt.js';

interface Index {
  /** Offsets of each item's start, one more than the item count. */
  starts: number[];
}

interface FontDict {
  subrs: Index | null;
  subrStart: number;
  /** CFF2: the default item variation data, for the blend operator's region count. */
  vsindex: number;
}

interface Cff {
  d: Uint8Array;
  version: number;
  charStrings: Index;
  globalSubrs: Index;
  fonts: FontDict[];
  /** Font dict of each glyph (CID-keyed and CFF2 fonts); null when there is one. */
  select: ((gid: number) => number) | null;
  /** Region counts of the variation store's item variation data (CFF2). */
  regions: number[];
  /** Font units per charstring unit, when the font matrix isn't the head table's em. */
  scale: number;
}

const cache = new WeakMap<SfntFont, Cff | null>();
const MAX_OPS = 100_000;
const MAX_CALL_DEPTH = 10;

class Truncated extends Error {}

function u8(d: Uint8Array, o: number): number {
  if (o < 0 || o >= d.length) throw new Truncated('Truncated CFF table');
  return d[o];
}
function u16(d: Uint8Array, o: number): number {
  return (u8(d, o) << 8) | u8(d, o + 1);
}
function u32(d: Uint8Array, o: number): number {
  return ((u8(d, o) << 24) | (u8(d, o + 1) << 16) | (u8(d, o + 2) << 8) | u8(d, o + 3)) >>> 0;
}
function offset(d: Uint8Array, o: number, size: number): number {
  let v = 0;
  for (let i = 0; i < size; i++) v = v * 256 + u8(d, o + i);
  return v;
}

/** An INDEX at `o`; `end` receives the byte after it. */
function readIndex(d: Uint8Array, o: number, cff2: boolean, end?: { at: number }): Index {
  const count = cff2 ? u32(d, o) : u16(d, o);
  const head = cff2 ? 4 : 2;
  if (!count) {
    if (end) end.at = o + head;
    return { starts: [o + head] };
  }
  const size = u8(d, o + head);
  if (size < 1 || size > 4) throw new Error('Invalid CFF index');
  const base = o + head + 1 + (count + 1) * size - 1;
  const starts = new Array<number>(count + 1);
  for (let i = 0; i <= count; i++) {
    starts[i] = base + offset(d, o + head + 1 + i * size, size);
    if (i && starts[i] < starts[i - 1]) throw new Error('Invalid CFF index');
  }
  if (starts[count] > d.length) throw new Truncated('Truncated CFF index');
  if (end) end.at = starts[count];
  return { starts };
}

function count(index: Index): number {
  return index.starts.length - 1;
}

/** A DICT's entries: operator (escaped ones as 1200 + code) to operands. */
function readDict(d: Uint8Array, start: number, end: number): Map<number, number[]> {
  const out = new Map<number, number[]>();
  let operands: number[] = [];
  for (let p = start; p < end; ) {
    const b = u8(d, p);
    if (b <= 21) {
      let op = b;
      p++;
      if (b === 12) op = 1200 + u8(d, p++);
      out.set(op, operands);
      operands = [];
    } else if (b === 28) {
      operands.push((u16(d, p + 1) << 16) >> 16);
      p += 3;
    } else if (b === 29) {
      operands.push(u32(d, p + 1) | 0);
      p += 5;
    } else if (b === 30) {
      let s = '';
      p++;
      for (let done = false; !done && p < end; p++) {
        for (const nibble of [u8(d, p) >> 4, u8(d, p) & 15]) {
          if (nibble === 15) {
            done = true;
            break;
          }
          s += nibble <= 9 ? String(nibble) : ['.', 'E', 'E-', '', '-'][nibble - 10];
        }
      }
      operands.push(Number(s) || 0);
    } else if (b >= 32 && b <= 246) {
      operands.push(b - 139);
      p++;
    } else if (b >= 247 && b <= 250) {
      operands.push((b - 247) * 256 + u8(d, p + 1) + 108);
      p += 2;
    } else if (b >= 251 && b <= 254) {
      operands.push(-(b - 251) * 256 - u8(d, p + 1) - 108);
      p += 2;
    } else p++;
  }
  return out;
}

function fontDict(d: Uint8Array, top: Map<number, number[]>): FontDict {
  const priv = top.get(18);
  if (!priv || priv.length < 2) return { subrs: null, subrStart: 0, vsindex: 0 };
  const [size, at] = priv;
  if (size < 0 || at < 0 || at + size > d.length) throw new Truncated('Truncated CFF private dict');
  const dict = readDict(d, at, at + size);
  const subrOffset = dict.get(19)?.[0];
  return {
    subrs: subrOffset ? readIndex(d, at + subrOffset, false) : null,
    subrStart: at,
    vsindex: dict.get(22)?.[0] ?? 0,
  };
}

function parse(font: SfntFont): Cff | null {
  const cff2 = font.tables.has('CFF2');
  const table = font.tableBytes(cff2 ? 'CFF2' : 'CFF ');
  if (!table) return null;
  const d = table;
  const version = u8(d, 0);
  let top: Map<number, number[]>;
  let globalSubrs: Index;
  if (version === 2) {
    const hdrSize = u8(d, 2);
    const topSize = u16(d, 3);
    top = readDict(d, hdrSize, hdrSize + topSize);
    globalSubrs = readIndex(d, hdrSize + topSize, true);
  } else {
    const end = { at: 0 };
    readIndex(d, u8(d, 2), false, end);
    const tops = readIndex(d, end.at, false, end);
    if (!count(tops)) return null;
    top = readDict(d, tops.starts[0], tops.starts[1]);
    readIndex(d, end.at, false, end);
    globalSubrs = readIndex(d, end.at, false);
  }
  if ((top.get(1206)?.[0] ?? 2) !== 2) return null;
  const charStringsAt = top.get(17)?.[0];
  if (!charStringsAt) return null;
  const charStrings = readIndex(d, charStringsAt, version === 2);

  let fonts: FontDict[] = [];
  let select: Cff['select'] = null;
  const fdArray = top.get(1236)?.[0];
  if (fdArray) {
    const fds = readIndex(d, fdArray, version === 2);
    for (let i = 0; i < count(fds); i++) fonts.push(fontDict(d, readDict(d, fds.starts[i], fds.starts[i + 1])));
    const fdSelect = top.get(1237)?.[0];
    if (fdSelect) select = readSelect(d, fdSelect, count(charStrings));
  }
  if (!fonts.length) fonts = [fontDict(d, top)];

  const regions: number[] = [];
  const vstore = top.get(24)?.[0];
  if (version === 2 && vstore) {
    // The store follows a 16-bit length: format, region list offset, then item variation data.
    const at = vstore + 2;
    const n = u16(d, at + 6);
    for (let i = 0; i < n; i++) {
      const data = at + u32(d, at + 8 + 4 * i);
      regions.push(u16(d, data + 4));
    }
  }

  const matrix = top.get(1207);
  const scale = matrix && matrix.length === 6 && matrix[0] > 0 ? matrix[0] * font.unitsPerEm : 1;
  return { d, version, charStrings, globalSubrs, fonts, select, regions, scale: Math.abs(scale - 1) < 1e-6 ? 1 : scale };
}

function readSelect(d: Uint8Array, at: number, glyphs: number): (gid: number) => number {
  const format = u8(d, at);
  if (format === 0) return (gid) => (gid < glyphs ? u8(d, at + 1 + gid) : 0);
  if (format === 3 || format === 4) {
    const wide = format === 4;
    const n = wide ? u32(d, at + 1) : u16(d, at + 1);
    const rec = wide ? 6 : 3;
    const first = at + (wide ? 5 : 3);
    const firstOf = (i: number) => (wide ? u32(d, first + i * rec) : u16(d, first + i * rec));
    const fdOf = (i: number) => (wide ? u16(d, first + i * rec + 4) : u8(d, first + i * rec + 2));
    return (gid) => {
      let lo = 0;
      let hi = n - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (firstOf(mid) > gid) hi = mid - 1;
        else if (mid + 1 < n && firstOf(mid + 1) <= gid) lo = mid + 1;
        else return fdOf(mid);
      }
      return 0;
    };
  }
  throw new Error('Unsupported CFF FDSelect format');
}

function bias(index: Index | null): number {
  const n = index ? count(index) : 0;
  return n < 1240 ? 107 : n < 33900 ? 1131 : 32768;
}

/** Glyph `gid`'s outline from the font's CFF or CFF2 table, or null when it can't be read. */
export function cffOutline(font: SfntFont, gid: number): Outline | null {
  let cff = cache.get(font);
  if (cff === undefined) {
    try {
      cff = parse(font);
    } catch {
      cff = null;
    }
    cache.set(font, cff);
  }
  if (!cff || gid >= count(cff.charStrings)) return null;
  const fd = cff.fonts[cff.select ? cff.select(gid) : 0] ?? cff.fonts[0];
  const out = new OutlineBuilder();
  run(cff, fd, cff.charStrings.starts[gid], cff.charStrings.starts[gid + 1], out);
  const outline = out.finish();
  if (cff.scale !== 1) for (let i = 0; i < outline.pts.length; i++) outline.pts[i] *= cff.scale;
  if (cff.scale !== 1) outline.bounds = outline.bounds.map((v) => v * cff.scale) as Outline['bounds'];
  return outline;
}

function run(cff: Cff, fd: FontDict, start: number, end: number, out: OutlineBuilder): void {
  const { d } = cff;
  const cff2 = cff.version === 2;
  const stack: number[] = [];
  const transient: number[] = new Array(32).fill(0);
  let stems = 0;
  let x = 0;
  let y = 0;
  let open = false;
  let widthSeen = cff2;
  let ops = 0;
  let vsindex = fd.vsindex;
  const localBias = bias(fd.subrs);
  const globalBias = bias(cff.globalSubrs);

  const moveTo = (dx: number, dy: number) => {
    x += dx;
    y += dy;
    out.move(x, y);
    open = true;
  };
  const lineTo = (dx: number, dy: number) => {
    x += dx;
    y += dy;
    out.line(x, y);
  };
  const curveTo = (a: number, b: number, c: number, e: number, f: number, g: number) => {
    const x1 = x + a;
    const y1 = y + b;
    const x2 = x1 + c;
    const y2 = y1 + e;
    x = x2 + f;
    y = y2 + g;
    out.cubic(x1, y1, x2, y2, x, y);
  };
  /** Drop the width a glyph's first stack-clearing operator may carry ahead of its arguments. */
  const width = (odd: boolean) => {
    if (!widthSeen && odd && stack.length) stack.shift();
    widthSeen = true;
  };
  const stemHints = () => {
    width(stack.length % 2 === 1);
    stems += stack.length >> 1;
    stack.length = 0;
  };

  const exec = (from: number, to: number, depth: number): boolean => {
    if (depth > MAX_CALL_DEPTH) throw new Error('CFF subroutines nest too deeply');
    let p = from;
    while (p < to) {
      if (++ops > MAX_OPS) throw new Error('CFF charstring too long');
      const b = d[p];
      if (b >= 32 || b === 28) {
        if (b === 28) {
          stack.push((u16(d, p + 1) << 16) >> 16);
          p += 3;
        } else if (b <= 246) {
          stack.push(b - 139);
          p++;
        } else if (b <= 250) {
          stack.push((b - 247) * 256 + u8(d, p + 1) + 108);
          p += 2;
        } else if (b <= 254) {
          stack.push(-(b - 251) * 256 - u8(d, p + 1) - 108);
          p += 2;
        } else {
          stack.push((u32(d, p + 1) | 0) / 65536);
          p += 5;
        }
        if (stack.length > 513) throw new Error('CFF stack overflow');
        continue;
      }
      p++;
      switch (b) {
        case 1:
        case 3:
        case 18:
        case 23:
          stemHints();
          break;
        case 19:
        case 20:
          if (stack.length) stemHints();
          widthSeen = true;
          p += (stems + 7) >> 3;
          break;
        case 21:
          width(stack.length > 2);
          moveTo(stack[0] ?? 0, stack[1] ?? 0);
          stack.length = 0;
          break;
        case 22:
          width(stack.length > 1);
          moveTo(stack[0] ?? 0, 0);
          stack.length = 0;
          break;
        case 4:
          width(stack.length > 1);
          moveTo(0, stack[0] ?? 0);
          stack.length = 0;
          break;
        case 5:
          for (let i = 0; i + 1 < stack.length; i += 2) lineTo(stack[i], stack[i + 1]);
          stack.length = 0;
          break;
        case 6:
        case 7: {
          let horizontal = b === 6;
          for (let i = 0; i < stack.length; i++) {
            if (horizontal) lineTo(stack[i], 0);
            else lineTo(0, stack[i]);
            horizontal = !horizontal;
          }
          stack.length = 0;
          break;
        }
        case 8:
          for (let i = 0; i + 5 < stack.length; i += 6) curveTo(stack[i], stack[i + 1], stack[i + 2], stack[i + 3], stack[i + 4], stack[i + 5]);
          stack.length = 0;
          break;
        case 24: {
          let i = 0;
          for (; i + 7 < stack.length; i += 6) curveTo(stack[i], stack[i + 1], stack[i + 2], stack[i + 3], stack[i + 4], stack[i + 5]);
          if (i + 1 < stack.length) lineTo(stack[i], stack[i + 1]);
          stack.length = 0;
          break;
        }
        case 25: {
          let i = 0;
          for (; i + 7 < stack.length; i += 2) lineTo(stack[i], stack[i + 1]);
          if (i + 5 < stack.length) curveTo(stack[i], stack[i + 1], stack[i + 2], stack[i + 3], stack[i + 4], stack[i + 5]);
          stack.length = 0;
          break;
        }
        case 26: {
          let i = 0;
          let dx1 = 0;
          if (stack.length % 2) dx1 = stack[i++];
          for (; i + 3 < stack.length; i += 4) {
            curveTo(dx1, stack[i], stack[i + 1], stack[i + 2], 0, stack[i + 3]);
            dx1 = 0;
          }
          stack.length = 0;
          break;
        }
        case 27: {
          let i = 0;
          let dy1 = 0;
          if (stack.length % 2) dy1 = stack[i++];
          for (; i + 3 < stack.length; i += 4) {
            curveTo(stack[i], dy1, stack[i + 1], stack[i + 2], stack[i + 3], 0);
            dy1 = 0;
          }
          stack.length = 0;
          break;
        }
        case 30:
        case 31: {
          let horizontal = b === 31;
          for (let i = 0; i + 3 < stack.length; i += 4) {
            const last = i + 5 === stack.length ? stack[i + 4] : 0;
            if (horizontal) curveTo(stack[i], 0, stack[i + 1], stack[i + 2], last, stack[i + 3]);
            else curveTo(0, stack[i], stack[i + 1], stack[i + 2], stack[i + 3], last);
            horizontal = !horizontal;
          }
          stack.length = 0;
          break;
        }
        case 10:
        case 29: {
          const subrs = b === 10 ? fd.subrs : cff.globalSubrs;
          const index = (stack.pop() ?? 0) + (b === 10 ? localBias : globalBias);
          if (!subrs || index < 0 || index >= count(subrs)) throw new Error('CFF subroutine out of range');
          if (exec(subrs.starts[index], subrs.starts[index + 1], depth + 1)) return true;
          break;
        }
        case 11:
          return false;
        case 14:
          // endchar; a seac accent (four more operands) isn't composed.
          width(stack.length === 1 || stack.length === 5);
          stack.length = 0;
          if (open) out.close();
          return true;
        case 15:
          vsindex = stack.pop() ?? 0;
          stack.length = 0;
          break;
        case 16: {
          // blend: keep the default values, drop each region's deltas.
          const n = stack.pop() ?? 0;
          const k = cff.regions[vsindex] ?? 0;
          const base = stack.length - n * (k + 1);
          if (n < 0 || base < 0) throw new Error('Invalid CFF2 blend');
          stack.length = base + n;
          break;
        }
        case 12: {
          const e = u8(d, p++);
          const s = stack;
          switch (e) {
            case 35: // flex
              if (s.length >= 12) {
                curveTo(s[0], s[1], s[2], s[3], s[4], s[5]);
                curveTo(s[6], s[7], s[8], s[9], s[10], s[11]);
              }
              s.length = 0;
              break;
            case 34: // hflex
              if (s.length >= 7) {
                const y0 = y;
                curveTo(s[0], 0, s[1], s[2], s[3], 0);
                curveTo(s[4], 0, s[5], y0 - y, s[6], 0);
              }
              s.length = 0;
              break;
            case 36: // hflex1
              if (s.length >= 9) {
                const y0 = y;
                curveTo(s[0], s[1], s[2], s[3], s[4], 0);
                curveTo(s[5], 0, s[6], s[7], s[8], y0 - (y + s[7]));
              }
              s.length = 0;
              break;
            case 37: // flex1
              if (s.length >= 11) {
                const [x0, y0] = [x, y];
                let dx = 0;
                let dy = 0;
                for (let i = 0; i < 10; i += 2) {
                  dx += s[i];
                  dy += s[i + 1];
                }
                curveTo(s[0], s[1], s[2], s[3], s[4], s[5]);
                const [cx1, cy1] = [x + s[6], y + s[7]];
                const [cx2, cy2] = [cx1 + s[8], cy1 + s[9]];
                const [ex, ey] = Math.abs(dx) > Math.abs(dy) ? [x0 + dx + s[10], y0] : [x0, y0 + dy + s[10]];
                out.cubic(cx1, cy1, cx2, cy2, ex, ey);
                x = ex;
                y = ey;
              }
              s.length = 0;
              break;
            case 3: {
              const [a, c] = [s.pop() ?? 0, s.pop() ?? 0];
              s.push(a && c ? 1 : 0);
              break;
            }
            case 4: {
              const [a, c] = [s.pop() ?? 0, s.pop() ?? 0];
              s.push(a || c ? 1 : 0);
              break;
            }
            case 5:
              s.push(s.pop() ? 0 : 1);
              break;
            case 9:
              s.push(Math.abs(s.pop() ?? 0));
              break;
            case 10:
              s.push((s.pop() ?? 0) + (s.pop() ?? 0));
              break;
            case 11: {
              const c = s.pop() ?? 0;
              s.push((s.pop() ?? 0) - c);
              break;
            }
            case 12: {
              const c = s.pop() ?? 1;
              s.push(c ? (s.pop() ?? 0) / c : 0);
              break;
            }
            case 14:
              s.push(-(s.pop() ?? 0));
              break;
            case 15:
              s.push(s.pop() === s.pop() ? 1 : 0);
              break;
            case 18:
              s.pop();
              break;
            case 20: {
              const i = s.pop() ?? 0;
              const v = s.pop() ?? 0;
              if (i >= 0 && i < transient.length) transient[i] = v;
              break;
            }
            case 21: {
              const i = s.pop() ?? 0;
              s.push(i >= 0 && i < transient.length ? transient[i] : 0);
              break;
            }
            case 22: {
              const [v2, v1, s2, s1] = [s.pop() ?? 0, s.pop() ?? 0, s.pop() ?? 0, s.pop() ?? 0];
              s.push(v1 <= v2 ? s1 : s2);
              break;
            }
            case 23:
              s.push(0.5);
              break;
            case 24:
              s.push((s.pop() ?? 0) * (s.pop() ?? 0));
              break;
            case 26:
              s.push(Math.sqrt(Math.max(0, s.pop() ?? 0)));
              break;
            case 27:
              s.push(s[s.length - 1] ?? 0);
              break;
            case 28: {
              const a = s.pop() ?? 0;
              const c = s.pop() ?? 0;
              s.push(a, c);
              break;
            }
            case 29: {
              const i = s.pop() ?? 0;
              s.push(s[s.length - 1 - Math.max(0, i)] ?? 0);
              break;
            }
            case 30: {
              const j = s.pop() ?? 0;
              const n = s.pop() ?? 0;
              if (n > 0 && n <= s.length) {
                const items = s.splice(s.length - n, n);
                const shift = ((j % n) + n) % n;
                s.push(...items.slice(n - shift), ...items.slice(0, n - shift));
              }
              break;
            }
            default:
              s.length = 0;
          }
          break;
        }
        default:
          stack.length = 0;
      }
    }
    return false;
  };

  exec(start, end, 0);
  if (open) out.close();
}
