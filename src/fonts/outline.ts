// Glyph outlines in font units (y up), for drawing glyphs as paths: TrueType (glyf) contours with
// their quadratic curves, and CFF or CFF2 charstrings with their cubic ones.

import { cffOutline } from './cff.js';
import type { SfntFont } from './sfnt.js';

export const MOVE = 0;
export const LINE = 1;
export const QUAD = 2;
export const CUBIC = 3;
export const CLOSE = 4;

/** A glyph outline: commands (MOVE, LINE, QUAD, CUBIC, CLOSE) and their points, x then y. */
export interface Outline {
  cmds: Uint8Array;
  /** 2 numbers per MOVE or LINE, 4 per QUAD, 6 per CUBIC, none per CLOSE. */
  pts: Float32Array;
  /** Bounds of the points [xMin, yMin, xMax, yMax]; all zero for an empty outline. */
  bounds: [number, number, number, number];
}

/** Collects an outline's commands and points. */
export class OutlineBuilder {
  readonly cmds: number[] = [];
  readonly pts: number[] = [];
  private open = false;

  move(x: number, y: number): void {
    if (this.open) this.close();
    this.cmds.push(MOVE);
    this.pts.push(x, y);
    this.open = true;
  }
  line(x: number, y: number): void {
    this.cmds.push(LINE);
    this.pts.push(x, y);
  }
  quad(cx: number, cy: number, x: number, y: number): void {
    this.cmds.push(QUAD);
    this.pts.push(cx, cy, x, y);
  }
  cubic(c1x: number, c1y: number, c2x: number, c2y: number, x: number, y: number): void {
    this.cmds.push(CUBIC);
    this.pts.push(c1x, c1y, c2x, c2y, x, y);
  }
  close(): void {
    if (!this.open) return;
    this.cmds.push(CLOSE);
    this.open = false;
  }

  finish(): Outline {
    this.close();
    const pts = Float32Array.from(this.pts);
    const bounds: Outline['bounds'] = [0, 0, 0, 0];
    if (pts.length) {
      bounds[0] = bounds[2] = pts[0];
      bounds[1] = bounds[3] = pts[1];
      for (let i = 2; i < pts.length; i += 2) {
        const x = pts[i];
        const y = pts[i + 1];
        if (x < bounds[0]) bounds[0] = x;
        else if (x > bounds[2]) bounds[2] = x;
        if (y < bounds[1]) bounds[1] = y;
        else if (y > bounds[3]) bounds[3] = y;
      }
    }
    return { cmds: Uint8Array.from(this.cmds), pts, bounds };
  }
}

const EMPTY: Outline = { cmds: new Uint8Array(0), pts: new Float32Array(0), bounds: [0, 0, 0, 0] };
// A font is untrusted input: composites nest to a bounded depth and total size.
const MAX_COMPOSITE_DEPTH = 8;
const MAX_OUTLINE_POINTS = 200_000;

const outlines = new WeakMap<SfntFont, Map<number, Outline>>();
/** Outline bytes read separately from the font's layout tables (see SfntFont.readSubsetGlyphs). */
const detached = new WeakMap<SfntFont, Map<number, Uint8Array>>();

/** Whether glyph outlines are read from outside the font's tables, and must be loaded first. */
export function needsOutlineLoad(font: SfntFont): boolean {
  return !font.table('glyf') && !font.isCFF && font.readSubsetGlyphs !== undefined;
}

/** Load the outlines of `gids` for a font whose outlines are kept apart from its tables. */
export async function loadOutlines(font: SfntFont, gids: Iterable<number>): Promise<void> {
  if (!needsOutlineLoad(font)) return;
  let have = detached.get(font);
  if (!have) detached.set(font, (have = new Map()));
  const missing = [...new Set(gids)].filter((g) => !have.has(g));
  if (!missing.length) return;
  const read = await font.readSubsetGlyphs!(missing);
  for (const [gid, bytes] of read) have.set(gid, bytes);
}

/** Glyph `gid`'s outline in font units, empty when it has none or can't be read. */
export function glyphPath(font: SfntFont, gid: number): Outline {
  let cache = outlines.get(font);
  if (!cache) outlines.set(font, (cache = new Map()));
  let out = cache.get(gid);
  if (out === undefined) {
    try {
      out = readOutline(font, gid);
    } catch {
      out = EMPTY;
    }
    cache.set(gid, out);
  }
  return out;
}

function readOutline(font: SfntFont, gid: number): Outline {
  if (!Number.isInteger(gid) || gid < 0 || gid >= font.numGlyphs) return EMPTY;
  if (font.isCFF) return cffOutline(font, gid) ?? EMPTY;
  const b = new OutlineBuilder();
  appendTrueType(font, gid, [1, 0, 0, 1, 0, 0], b, 0, { points: 0 });
  return b.cmds.length ? b.finish() : EMPTY;
}

type Matrix = [number, number, number, number, number, number];

function glyphBytes(font: SfntFont, gid: number): Uint8Array | null {
  const glyf = font.table('glyf');
  if (glyf) {
    const range = font.glyphRange(gid, glyf);
    return range ? font.data.subarray(range[0], range[1]) : null;
  }
  return detached.get(font)?.get(gid) ?? null;
}

function appendTrueType(font: SfntFont, gid: number, m: Matrix, out: OutlineBuilder, depth: number, budget: { points: number }): void {
  if (depth > MAX_COMPOSITE_DEPTH) throw new Error('Composite glyph too deep');
  const d = glyphBytes(font, gid);
  if (!d || d.length < 10) return;
  const end = d.length;
  const u8 = (o: number) => {
    if (o >= end) throw new Error('Truncated glyph');
    return d[o];
  };
  const u16 = (o: number) => (u8(o) << 8) | u8(o + 1);
  const i16 = (o: number) => (u16(o) << 16) >> 16;
  const contours = i16(0);
  if (contours < 0) {
    let p = 10;
    for (;;) {
      const flags = u16(p);
      const component = u16(p + 2);
      p += 4;
      let dx = 0;
      let dy = 0;
      if (flags & 0x0001) {
        if (flags & 0x0002) [dx, dy] = [i16(p), i16(p + 2)];
        p += 4;
      } else {
        if (flags & 0x0002) [dx, dy] = [(u8(p) << 24) >> 24, (u8(p + 1) << 24) >> 24];
        p += 2;
      }
      let [a, b, c, e] = [1, 0, 0, 1];
      if (flags & 0x0008) {
        a = e = i16(p) / 16384;
        p += 2;
      } else if (flags & 0x0040) {
        [a, e] = [i16(p) / 16384, i16(p + 2) / 16384];
        p += 4;
      } else if (flags & 0x0080) {
        [a, b, c, e] = [i16(p) / 16384, i16(p + 2) / 16384, i16(p + 4) / 16384, i16(p + 6) / 16384];
        p += 8;
      }
      // A scaled component offset (rare) goes through the component's matrix too.
      if (flags & 0x0800) [dx, dy] = [a * dx + c * dy, b * dx + e * dy];
      appendTrueType(font, component, multiply([a, b, c, e, dx, dy], m), out, depth + 1, budget);
      if (!(flags & 0x0020)) break;
    }
    return;
  }
  if (!contours) return;
  const ends = new Array<number>(contours);
  for (let i = 0; i < contours; i++) ends[i] = u16(10 + 2 * i);
  const count = ends[contours - 1] + 1;
  budget.points += count;
  if (budget.points > MAX_OUTLINE_POINTS) throw new Error('Glyph outline too large');
  let p = 10 + 2 * contours;
  p += 2 + u16(p);
  const flags = new Uint8Array(count);
  for (let i = 0; i < count; ) {
    const f = u8(p++);
    flags[i++] = f;
    if (f & 0x08) {
      const repeat = u8(p++);
      for (let k = 0; k < repeat && i < count; k++) flags[i++] = f;
    }
  }
  const xs = new Float64Array(count);
  const ys = new Float64Array(count);
  let v = 0;
  for (let i = 0; i < count; i++) {
    const f = flags[i];
    if (f & 0x02) v += f & 0x10 ? u8(p++) : -u8(p++);
    else if (!(f & 0x10)) {
      v += i16(p);
      p += 2;
    }
    xs[i] = v;
  }
  v = 0;
  for (let i = 0; i < count; i++) {
    const f = flags[i];
    if (f & 0x04) v += f & 0x20 ? u8(p++) : -u8(p++);
    else if (!(f & 0x20)) {
      v += i16(p);
      p += 2;
    }
    ys[i] = v;
  }
  const [ma, mb, mc, md, me, mf] = m;
  for (let i = 0; i < count; i++) {
    const x = xs[i];
    const y = ys[i];
    xs[i] = ma * x + mc * y + me;
    ys[i] = mb * x + md * y + mf;
  }
  let first = 0;
  for (const last of ends) {
    if (last < first || last >= count) throw new Error('Invalid contour');
    const n = last - first + 1;
    if (n > 1) {
      const on = (k: number) => (flags[first + (k % n)] & 1) === 1;
      const X = (k: number) => xs[first + (k % n)];
      const Y = (k: number) => ys[first + (k % n)];
      // Start on an on-curve point, or midway between two off-curve ones when there is none.
      let s = 0;
      while (s < n && !on(s)) s++;
      let bx: number;
      let by: number;
      let from: number;
      if (s === n) {
        bx = (X(n - 1) + X(0)) / 2;
        by = (Y(n - 1) + Y(0)) / 2;
        from = 0;
      } else {
        bx = X(s);
        by = Y(s);
        from = s + 1;
      }
      out.move(bx, by);
      let hasControl = false;
      let cx = 0;
      let cy = 0;
      for (let k = from; k < from + n - (s === n ? 0 : 1); k++) {
        const qx = X(k);
        const qy = Y(k);
        if (on(k)) {
          if (hasControl) out.quad(cx, cy, qx, qy);
          else out.line(qx, qy);
          hasControl = false;
        } else {
          if (hasControl) out.quad(cx, cy, (cx + qx) / 2, (cy + qy) / 2);
          cx = qx;
          cy = qy;
          hasControl = true;
        }
      }
      if (hasControl) out.quad(cx, cy, bx, by);
      out.close();
    }
    first = last + 1;
  }
}

/** The matrix that applies `a`, then `b`. */
function multiply(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[1] * b[2],
    a[0] * b[1] + a[1] * b[3],
    a[2] * b[0] + a[3] * b[2],
    a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4],
    a[4] * b[1] + a[5] * b[3] + b[5],
  ];
}
