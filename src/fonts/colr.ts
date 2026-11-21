// Colour glyphs (OpenType COLR and CPAL tables): the layers of Segoe UI Emoji (COLR version 0) and
// the gradients, transforms and compositing of Noto Color Emoji (version 1), read into a paint tree
// that the PDF writer draws. Glyph outlines (TrueType glyf) are read as PDF path operators.

import { num } from '../pdf/objects.js';
import type { SfntFont } from './sfnt.js';

/** A colour: `rgb` RRGGBB, or null for the text's own colour (palette index 0xFFFF). */
export interface PaintColor {
  rgb: string | null;
  alpha: number;
}

export interface ColorStop {
  offset: number;
  color: PaintColor;
}

/** A 2D affine matrix [a b c d e f], as PDF's cm takes it. */
export type Matrix = [number, number, number, number, number, number];

export type ColorPaint =
  | { kind: 'layers'; layers: ColorPaint[] }
  | { kind: 'solid'; color: PaintColor }
  /** Colour runs from (x0, y0) at offset 0 to (x1, y1) at 1, constant along perpendicular lines. */
  | { kind: 'linear'; stops: ColorStop[]; x0: number; y0: number; x1: number; y1: number }
  /** Colour runs from the circle (x0, y0, r0) at offset 0 to the circle (x1, y1, r1) at 1. */
  | { kind: 'radial'; stops: ColorStop[]; x0: number; y0: number; r0: number; x1: number; y1: number; r1: number }
  /** `paint` drawn inside the outline of glyph `gid`. */
  | { kind: 'glyph'; gid: number; paint: ColorPaint }
  | { kind: 'transform'; matrix: Matrix; paint: ColorPaint }
  /** `source` composited onto `backdrop` by one of COLR's composite modes. */
  | { kind: 'composite'; mode: number; source: ColorPaint; backdrop: ColorPaint };

export interface ColorGlyph {
  paint: ColorPaint;
  /** Where the glyph draws, in font units [xMin, yMin, xMax, yMax], when the font says. */
  clip: [number, number, number, number] | null;
  /** Whether any colour is the text's own (so the glyph's look depends on it). */
  foreground: boolean;
}

export const COMPOSITE = {
  CLEAR: 0, SRC: 1, DEST: 2, SRC_OVER: 3, DEST_OVER: 4, SRC_IN: 5, DEST_IN: 6, SRC_OUT: 7, DEST_OUT: 8,
  SRC_ATOP: 9, DEST_ATOP: 10, XOR: 11, PLUS: 12,
} as const;

/** PDF blend modes for COLR's composite modes 13 and up (screen to luminosity). */
export const BLEND_MODES = ['Screen', 'Overlay', 'Darken', 'Lighten', 'ColorDodge', 'ColorBurn', 'HardLight', 'SoftLight', 'Difference', 'Exclusion', 'Multiply', 'Hue', 'Saturation', 'Color', 'Luminosity'];

// A font is untrusted input: paint graphs are walked to a bounded depth and size.
const MAX_DEPTH = 64;
const MAX_NODES = 20_000;
const MAX_OUTLINE_POINTS = 200_000;

class Bytes {
  constructor(readonly d: Uint8Array, readonly start: number, readonly end: number) {}
  private need(o: number, n: number): void {
    if (o < this.start || o + n > this.end) throw new Error('Truncated colour table');
  }
  u8(o: number): number {
    this.need(o, 1);
    return this.d[o];
  }
  u16(o: number): number {
    this.need(o, 2);
    return (this.d[o] << 8) | this.d[o + 1];
  }
  i16(o: number): number {
    const v = this.u16(o);
    return v & 0x8000 ? v - 0x10000 : v;
  }
  u24(o: number): number {
    this.need(o, 3);
    return (this.d[o] << 16) | (this.d[o + 1] << 8) | this.d[o + 2];
  }
  u32(o: number): number {
    this.need(o, 4);
    return ((this.d[o] << 24) | (this.d[o + 1] << 16) | (this.d[o + 2] << 8) | this.d[o + 3]) >>> 0;
  }
  i32(o: number): number {
    return this.u32(o) | 0;
  }
  f2dot14(o: number): number {
    return this.i16(o) / 16384;
  }
  fixed(o: number): number {
    return this.i32(o) / 65536;
  }
}

interface ColorTables {
  r: Bytes;
  colr: number;
  version: number;
  palette: Array<{ rgb: string; alpha: number }>;
  glyphs: Map<number, ColorGlyph | null>;
}

const tablesCache = new WeakMap<SfntFont, ColorTables | null>();

function colorTables(font: SfntFont): ColorTables | null {
  let t = tablesCache.get(font);
  if (t !== undefined) return t;
  t = null;
  const colr = font.table('COLR');
  const cpal = font.table('CPAL');
  // Layers are drawn from TrueType outlines.
  if (colr && cpal && font.table('glyf') && font.table('loca')) {
    try {
      const r = new Bytes(font.data, colr.offset, colr.offset + colr.length);
      const p = new Bytes(font.data, cpal.offset, cpal.offset + cpal.length);
      // The first palette: CPAL records are blue, green, red, alpha.
      const entries = p.u16(cpal.offset + 2);
      const records = cpal.offset + p.u32(cpal.offset + 8);
      const first = p.u16(cpal.offset + 12);
      const palette: ColorTables['palette'] = [];
      for (let i = 0; i < entries; i++) {
        const o = records + 4 * (first + i);
        const hex = (v: number) => v.toString(16).padStart(2, '0');
        palette.push({ rgb: (hex(p.u8(o + 2)) + hex(p.u8(o + 1)) + hex(p.u8(o))).toUpperCase(), alpha: p.u8(o + 3) / 255 });
      }
      t = { r, colr: colr.offset, version: r.u16(colr.offset), palette, glyphs: new Map() };
    } catch {
      t = null;
    }
  }
  tablesCache.set(font, t);
  return t;
}

/** Whether the font has colour glyphs (COLR and CPAL tables) with TrueType outlines. */
export function hasColorGlyphs(font: SfntFont): boolean {
  return colorTables(font) !== null;
}

/** The colour glyph for `gid`, or null when it has none or its tables are malformed. */
export function colorGlyph(font: SfntFont, gid: number): ColorGlyph | null {
  const t = colorTables(font);
  if (!t) return null;
  let g = t.glyphs.get(gid);
  if (g === undefined) {
    try {
      g = readColorGlyph(t, gid);
    } catch {
      g = null;
    }
    t.glyphs.set(gid, g);
  }
  return g;
}

/** Binary search of records `size` bytes long, sorted by the glyph id at their start. */
function findRecord(r: Bytes, start: number, count: number, size: number, gid: number): number {
  let lo = 0;
  let hi = count - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const g = r.u16(start + mid * size);
    if (g === gid) return start + mid * size;
    if (g < gid) lo = mid + 1;
    else hi = mid - 1;
  }
  return -1;
}

function readColorGlyph(t: ColorTables, gid: number): ColorGlyph | null {
  const { r, colr } = t;
  let foreground = false;
  let nodes = 0;
  const color = (index: number, alpha: number): PaintColor => {
    if (index === 0xffff) {
      foreground = true;
      return { rgb: null, alpha };
    }
    const c = t.palette[index];
    if (!c) throw new Error('Palette index out of range');
    return { rgb: c.rgb, alpha: c.alpha * alpha };
  };

  // Version 0: layers of glyphs, each in one palette colour. A font with both versions' records
  // for a glyph draws its version 0 layers, as Word for Mac draws Segoe UI Emoji's.
  const count = r.u16(colr + 2);
  const base = count ? findRecord(r, colr + r.u32(colr + 4), count, 6, gid) : -1;
  if (base >= 0) {
    const first = r.u16(base + 2);
    const n = r.u16(base + 4);
    const layerRecords = colr + r.u32(colr + 8);
    if (first + n > r.u16(colr + 12)) throw new Error('Layer index out of range');
    const layers: ColorPaint[] = [];
    for (let k = 0; k < n; k++) {
      const rec = layerRecords + 4 * (first + k);
      layers.push({ kind: 'glyph', gid: r.u16(rec), paint: { kind: 'solid', color: color(r.u16(rec + 2), 1) } });
    }
    return { paint: { kind: 'layers', layers }, clip: null, foreground };
  }

  if (t.version >= 1) {
    const list = r.u32(colr + 14);
    const at = list ? findRecord(r, colr + list + 4, r.u32(colr + list), 6, gid) : -1;
    if (at >= 0) {
      const layerList = r.u32(colr + 18);
      const clipList = r.u32(colr + 22);
      const active = new Set<number>([gid]);
      const colorLine = (o: number, varStops: boolean): ColorStop[] => {
        const n = r.u16(o + 1);
        const size = varStops ? 10 : 6;
        const stops: ColorStop[] = [];
        for (let k = 0; k < n; k++) {
          const s = o + 3 + k * size;
          stops.push({ offset: r.f2dot14(s), color: color(r.u16(s + 2), r.f2dot14(s + 4)) });
        }
        return stops.sort((a, b) => a.offset - b.offset);
      };
      const transformed = (o: number, matrix: Matrix, depth: number): ColorPaint => ({ kind: 'transform', matrix, paint: paint(o + r.u24(o + 1), depth + 1) });
      const around = (m: Matrix, cx: number, cy: number): Matrix => [m[0], m[1], m[2], m[3], cx - m[0] * cx - m[2] * cy, cy - m[1] * cx - m[3] * cy];
      const rotation = (turns: number): Matrix => {
        const a = turns * Math.PI;
        return [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0];
      };
      const skew = (x: number, y: number): Matrix => [1, Math.tan(y * Math.PI), -Math.tan(x * Math.PI), 1, 0, 0];
      const paint = (o: number, depth: number): ColorPaint => {
        if (depth > MAX_DEPTH || ++nodes > MAX_NODES) throw new Error('Colour glyph too complex');
        const format = r.u8(o);
        switch (format) {
          case 1: {
            const n = r.u8(o + 1);
            const first = r.u32(o + 2);
            if (!layerList || first + n > r.u32(colr + layerList)) throw new Error('Layer index out of range');
            const layers: ColorPaint[] = [];
            for (let k = 0; k < n; k++) layers.push(paint(colr + layerList + r.u32(colr + layerList + 4 + 4 * (first + k)), depth + 1));
            return { kind: 'layers', layers };
          }
          case 2:
          case 3:
            return { kind: 'solid', color: color(r.u16(o + 1), r.f2dot14(o + 3)) };
          case 4:
          case 5: {
            const stops = colorLine(o + r.u24(o + 1), format === 5);
            const [x0, y0, x1, y1, x2, y2] = [4, 6, 8, 10, 12, 14].map((k) => r.i16(o + k));
            // Colour is constant along lines parallel to p0p2, so it runs along the perpendicular to
            // p0p2 through p0, to where p1 projects onto it.
            const [nx, ny] = [y2 - y0, -(x2 - x0)];
            const nn = nx * nx + ny * ny;
            const d = nn ? ((x1 - x0) * nx + (y1 - y0) * ny) / nn : 0;
            return nn ? { kind: 'linear', stops, x0, y0, x1: x0 + nx * d, y1: y0 + ny * d } : { kind: 'linear', stops, x0, y0, x1, y1 };
          }
          case 6:
          case 7: {
            const stops = colorLine(o + r.u24(o + 1), format === 7);
            return { kind: 'radial', stops, x0: r.i16(o + 4), y0: r.i16(o + 6), r0: r.u16(o + 8), x1: r.i16(o + 10), y1: r.i16(o + 12), r1: r.u16(o + 14) };
          }
          case 8:
          case 9: {
            // Sweep gradients have no PDF shading; their middle colour stands in.
            const stops = colorLine(o + r.u24(o + 1), format === 9);
            if (!stops.length) return { kind: 'layers', layers: [] };
            return { kind: 'solid', color: stops[Math.floor(stops.length / 2)].color };
          }
          case 10:
            return { kind: 'glyph', gid: r.u16(o + 4), paint: paint(o + r.u24(o + 1), depth + 1) };
          case 11: {
            const target = r.u16(o + 1);
            const rec = findRecord(r, colr + list + 4, r.u32(colr + list), 6, target);
            if (rec < 0 || active.has(target)) throw new Error('Colour glyph reference out of range');
            active.add(target);
            const out = paint(colr + list + r.u32(rec + 2), depth + 1);
            active.delete(target);
            return out;
          }
          case 12:
          case 13: {
            const m = o + r.u24(o + 4);
            return transformed(o, [r.fixed(m), r.fixed(m + 4), r.fixed(m + 8), r.fixed(m + 12), r.fixed(m + 16), r.fixed(m + 20)], depth);
          }
          case 14:
          case 15:
            return transformed(o, [1, 0, 0, 1, r.i16(o + 4), r.i16(o + 6)], depth);
          case 16:
          case 17:
            return transformed(o, [r.f2dot14(o + 4), 0, 0, r.f2dot14(o + 6), 0, 0], depth);
          case 18:
          case 19:
            return transformed(o, around([r.f2dot14(o + 4), 0, 0, r.f2dot14(o + 6), 0, 0], r.i16(o + 8), r.i16(o + 10)), depth);
          case 20:
          case 21:
            return transformed(o, [r.f2dot14(o + 4), 0, 0, r.f2dot14(o + 4), 0, 0], depth);
          case 22:
          case 23:
            return transformed(o, around([r.f2dot14(o + 4), 0, 0, r.f2dot14(o + 4), 0, 0], r.i16(o + 6), r.i16(o + 8)), depth);
          case 24:
          case 25:
            return transformed(o, rotation(r.f2dot14(o + 4)), depth);
          case 26:
          case 27:
            return transformed(o, around(rotation(r.f2dot14(o + 4)), r.i16(o + 6), r.i16(o + 8)), depth);
          case 28:
          case 29:
            return transformed(o, skew(r.f2dot14(o + 4), r.f2dot14(o + 6)), depth);
          case 30:
          case 31:
            return transformed(o, around(skew(r.f2dot14(o + 4), r.f2dot14(o + 6)), r.i16(o + 8), r.i16(o + 10)), depth);
          case 32:
            return { kind: 'composite', mode: r.u8(o + 4), source: paint(o + r.u24(o + 1), depth + 1), backdrop: paint(o + r.u24(o + 5), depth + 1) };
          default:
            throw new Error(`Unknown paint format ${format}`);
        }
      };
      const root = paint(colr + list + r.u32(at + 2), 0);
      let clip: ColorGlyph['clip'] = null;
      if (clipList) {
        const cl = colr + clipList;
        const n = r.u32(cl + 1);
        for (let lo = 0, hi = n - 1; lo <= hi; ) {
          const mid = (lo + hi) >> 1;
          const rec = cl + 5 + mid * 7;
          const [first, last] = [r.u16(rec), r.u16(rec + 2)];
          if (gid < first) hi = mid - 1;
          else if (gid > last) lo = mid + 1;
          else {
            const box = cl + r.u24(rec + 4);
            clip = [r.i16(box + 1), r.i16(box + 3), r.i16(box + 5), r.i16(box + 7)];
            break;
          }
        }
      }
      return { paint: root, clip, foreground };
    }
  }

  return null;
}

const outlineCache = new WeakMap<SfntFont, Map<number, string>>();

/**
 * Glyph `gid`'s outline as PDF path operators in font units (y up), quadratic curves raised to
 * cubic; empty when the glyph has none or isn't a TrueType outline.
 */
export function glyphOutline(font: SfntFont, gid: number): string {
  let cache = outlineCache.get(font);
  if (!cache) outlineCache.set(font, (cache = new Map()));
  let out = cache.get(gid);
  if (out === undefined) {
    const parts: string[] = [];
    try {
      appendOutline(font, gid, [1, 0, 0, 1, 0, 0], parts, 0, { points: 0 });
      out = parts.join('');
    } catch {
      out = '';
    }
    cache.set(gid, out);
  }
  return out;
}

function appendOutline(font: SfntFont, gid: number, m: Matrix, out: string[], depth: number, budget: { points: number }): void {
  if (depth > 8) throw new Error('Composite glyph too deep');
  const glyf = font.table('glyf');
  const range = glyf && font.glyphRange(gid, glyf);
  if (!range || range[1] - range[0] < 10) return;
  const [start, end] = range;
  const r = new Bytes(font.data, start, end);
  const contours = r.i16(start);
  if (contours < 0) {
    // A composite: components with offsets and optional scales.
    let p = start + 10;
    for (;;) {
      const flags = r.u16(p);
      const component = r.u16(p + 2);
      p += 4;
      let dx = 0;
      let dy = 0;
      if (flags & 0x0001) {
        if (flags & 0x0002) [dx, dy] = [r.i16(p), r.i16(p + 2)];
        p += 4;
      } else {
        if (flags & 0x0002) [dx, dy] = [(r.u8(p) << 24) >> 24, (r.u8(p + 1) << 24) >> 24];
        p += 2;
      }
      let [a, b, c, d] = [1, 0, 0, 1];
      if (flags & 0x0008) {
        a = d = r.f2dot14(p);
        p += 2;
      } else if (flags & 0x0040) {
        [a, d] = [r.f2dot14(p), r.f2dot14(p + 2)];
        p += 4;
      } else if (flags & 0x0080) {
        [a, b, c, d] = [r.f2dot14(p), r.f2dot14(p + 2), r.f2dot14(p + 4), r.f2dot14(p + 6)];
        p += 8;
      }
      // The component's points go through its own matrix, then the parent's.
      const local: Matrix = [a, b, c, d, dx, dy];
      appendOutline(font, component, multiply(local, m), out, depth + 1, budget);
      if (!(flags & 0x0020)) break;
    }
    return;
  }
  const ends: number[] = [];
  for (let i = 0; i < contours; i++) ends.push(r.u16(start + 10 + 2 * i));
  const count = contours ? ends[contours - 1] + 1 : 0;
  budget.points += count;
  if (budget.points > MAX_OUTLINE_POINTS) throw new Error('Glyph outline too large');
  let p = start + 10 + 2 * contours;
  p += 2 + r.u16(p);
  const flags = new Uint8Array(count);
  for (let i = 0; i < count; ) {
    const f = r.u8(p++);
    flags[i++] = f;
    if (f & 0x08) {
      const repeat = r.u8(p++);
      for (let k = 0; k < repeat && i < count; k++) flags[i++] = f;
    }
  }
  const xs = new Float64Array(count);
  const ys = new Float64Array(count);
  let v = 0;
  for (let i = 0; i < count; i++) {
    const f = flags[i];
    if (f & 0x02) v += f & 0x10 ? r.u8(p++) : -r.u8(p++);
    else if (!(f & 0x10)) {
      v += r.i16(p);
      p += 2;
    }
    xs[i] = v;
  }
  v = 0;
  for (let i = 0; i < count; i++) {
    const f = flags[i];
    if (f & 0x04) v += f & 0x20 ? r.u8(p++) : -r.u8(p++);
    else if (!(f & 0x20)) {
      v += r.i16(p);
      p += 2;
    }
    ys[i] = v;
  }
  const X = (x: number, y: number) => m[0] * x + m[2] * y + m[4];
  const Y = (x: number, y: number) => m[1] * x + m[3] * y + m[5];
  const pt = (x: number, y: number) => `${num(X(x, y))} ${num(Y(x, y))}`;
  let first = 0;
  for (const last of ends) {
    const n = last - first + 1;
    if (n > 1) {
      const on = (k: number) => (flags[first + ((k + n) % n)] & 1) === 1;
      const at = (k: number): [number, number] => [xs[first + ((k + n) % n)], ys[first + ((k + n) % n)]];
      // Start on an on-curve point, or midway between two off-curve ones when there is none.
      let s = 0;
      while (s < n && !on(s)) s++;
      const order: number[] = [];
      let begin: [number, number];
      if (s === n) {
        const [a, b] = [at(n - 1), at(0)];
        begin = [(a[0] + b[0]) / 2, (a[1] + b[1]) / 2];
        for (let k = 0; k < n; k++) order.push(k);
      } else {
        begin = at(s);
        for (let k = 1; k < n; k++) order.push(s + k);
      }
      let cur = begin;
      out.push(`${pt(...cur)} m `);
      let control: [number, number] | null = null;
      const quad = (c: [number, number], e: [number, number]) => {
        const c1: [number, number] = [cur[0] + (2 / 3) * (c[0] - cur[0]), cur[1] + (2 / 3) * (c[1] - cur[1])];
        const c2: [number, number] = [e[0] + (2 / 3) * (c[0] - e[0]), e[1] + (2 / 3) * (c[1] - e[1])];
        out.push(`${pt(...c1)} ${pt(...c2)} ${pt(...e)} c `);
        cur = e;
      };
      for (const k of order) {
        const q = at(k);
        if (on(k)) {
          if (control) quad(control, q);
          else {
            out.push(`${pt(...q)} l `);
            cur = q;
          }
          control = null;
        } else if (control) {
          quad(control, [(control[0] + q[0]) / 2, (control[1] + q[1]) / 2]);
          control = q;
        } else control = q;
      }
      if (control) quad(control, begin);
      out.push('h ');
    }
    first = last + 1;
  }
}

/** The matrix that applies `a`, then `b`. */
export function multiply(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[1] * b[2],
    a[0] * b[1] + a[1] * b[3],
    a[2] * b[0] + a[3] * b[2],
    a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4],
    a[4] * b[1] + a[5] * b[3] + b[5],
  ];
}
