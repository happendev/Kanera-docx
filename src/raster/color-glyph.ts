// Colour glyphs (COLR version 0 layers and version 1 paint graphs, see fonts/colr.ts) painted into
// a premultiplied layer the size of the glyph, then composited onto the page as one group.
// Gradients extend from their end stops, as the PDF writer's shadings do.

import { BLEND_MODES, COMPOSITE, colorGlyph, hasColorGlyphs, type ColorPaint, type ColorStop, type PaintColor } from '../fonts/colr.js';
import { glyphPath } from '../fonts/outline.js';
import type { SfntFont } from '../fonts/sfnt.js';
import type { Canvas } from './canvas.js';
import { Coverage } from './coverage.js';
import { concat, flattenOutline, invert, Polylines, type Matrix } from './geometry.js';
import type { ColorGlyphs } from './text.js';

type Box = [number, number, number, number];

interface Rendered {
  /** Offset of the layer's top-left pixel from the glyph origin's pixel. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Premultiplied RGBA, 0–255. */
  data: Uint8Array;
}

const MAX_CACHE_BYTES = 16 * 1024 * 1024;

export class ColrPainter implements ColorGlyphs {
  private readonly coverage = new Coverage();
  private readonly flat = new Polylines();
  private readonly cache = new Map<string, Rendered | null>();
  private readonly ids = new WeakMap<SfntFont, number>();
  private nextId = 1;
  private bytes = 0;

  has(font: SfntFont, gid: number): boolean {
    return hasColorGlyphs(font) && colorGlyph(font, gid) !== null;
  }

  draw(canvas: Canvas, font: SfntFont, gid: number, m: Matrix, fg: number, alpha: number): void {
    const glyph = colorGlyph(font, gid);
    if (!glyph) return;
    // Quarter-pixel positions share one rendering, as plain glyphs do.
    let ix = Math.floor(m[4]);
    let iy = Math.floor(m[5]);
    let fx = Math.round((m[4] - ix) * 4);
    let fy = Math.round((m[5] - iy) * 4);
    if (fx === 4) {
      ix++;
      fx = 0;
    }
    if (fy === 4) {
      iy++;
      fy = 0;
    }
    let id = this.ids.get(font);
    if (id === undefined) this.ids.set(font, (id = this.nextId++));
    const key = `${id}|${gid}|${m[0]}|${m[1]}|${m[2]}|${m[3]}|${fx}|${fy}|${glyph.foreground ? fg : ''}`;
    let r = this.cache.get(key);
    if (r === undefined) {
      const local: Matrix = [m[0], m[1], m[2], m[3], fx / 4, fy / 4];
      const [x0, y0, x1, y1] = font.bbox;
      const pad = font.unitsPerEm / 4;
      const box: Box = glyph.clip ?? [Math.min(x0, 0) - pad, Math.min(y0, 0) - pad, Math.max(x1, font.unitsPerEm) + pad, Math.max(y1, font.unitsPerEm) + pad];
      r = this.render(font, glyph.paint, box, local, fg);
      if (r) {
        this.bytes += r.data.length;
        if (this.bytes > MAX_CACHE_BYTES) {
          this.cache.clear();
          this.bytes = r.data.length;
        }
      }
      this.cache.set(key, r);
    }
    if (r) composite(canvas, r, ix + r.x, iy + r.y, alpha);
  }

  private render(font: SfntFont, paint: ColorPaint, box: Box, m: Matrix, fg: number): Rendered | null {
    const xs: number[] = [];
    const ys: number[] = [];
    for (const [x, y] of [[box[0], box[1]], [box[2], box[1]], [box[0], box[3]], [box[2], box[3]]]) {
      xs.push(m[0] * x + m[2] * y + m[4]);
      ys.push(m[1] * x + m[3] * y + m[5]);
    }
    const wx = Math.floor(Math.min(...xs));
    const wy = Math.floor(Math.min(...ys));
    const w = Math.ceil(Math.max(...xs)) - wx;
    const h = Math.ceil(Math.max(...ys)) - wy;
    if (w <= 0 || h <= 0 || w * h > 4_000_000) return null;
    // Device pixels of the layer, from its own origin.
    const device = concat(m, [1, 0, 0, 1, -wx, -wy]);
    const painter = new Painter(font, fg, w, h, this.coverage, this.flat);
    const layer = new Float32Array(w * h * 4);
    painter.paint(paint, [1, 0, 0, 1, 0, 0], device, layer, null);
    const data = new Uint8Array(w * h * 4);
    let any = false;
    for (let i = 0; i < layer.length; i++) {
      const v = layer[i] * 255;
      data[i] = v <= 0 ? 0 : v >= 255 ? 255 : (v + 0.5) | 0;
      if (data[i]) any = true;
    }
    return any ? { x: wx, y: wy, w, h, data } : null;
  }
}

class Painter {
  private readonly fgColor: [number, number, number];

  constructor(
    private readonly font: SfntFont,
    fg: number,
    private readonly w: number,
    private readonly h: number,
    private readonly coverage: Coverage,
    private readonly flat: Polylines,
  ) {
    this.fgColor = [((fg >> 16) & 255) / 255, ((fg >> 8) & 255) / 255, (fg & 255) / 255];
  }

  private rgb(c: PaintColor): [number, number, number] {
    if (c.rgb === null) return this.fgColor;
    const v = parseInt(c.rgb, 16) || 0;
    return [((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255];
  }

  /** Paint `p` (in a space `pm` takes to font units, `d` on to layer pixels) over `layer`. */
  paint(p: ColorPaint, pm: Matrix, d: Matrix, layer: Float32Array, mask: Float32Array | null): void {
    switch (p.kind) {
      case 'layers':
        for (const l of p.layers) this.paint(l, pm, d, layer, mask);
        return;
      case 'solid': {
        const a = p.color.alpha;
        if (a <= 0) return;
        const [r, g, b] = this.rgb(p.color);
        for (let i = 0, o = 0; i < this.w * this.h; i++, o += 4) {
          const k = mask ? a * mask[i] : a;
          if (k <= 0) continue;
          over(layer, o, r * k, g * k, b * k, k);
        }
        return;
      }
      case 'glyph': {
        const outline = glyphPath(this.font, p.gid);
        if (!outline.cmds.length) return;
        const flat = this.flat;
        flat.reset();
        flattenOutline(outline.cmds, outline.pts, concat(pm, d), flat);
        const clip = new Float32Array(this.w * this.h);
        const cov = this.coverage;
        if (!cov.begin(0, 0, this.w, this.h)) return;
        cov.addPolylines(flat);
        const w = this.w;
        cov.sweep(false, (y, from, to, row) => {
          for (let x = from; x < to; x++) clip[y * w + x] = mask ? row[x] * mask[y * w + x] : row[x];
        });
        this.paint(p.paint, pm, d, layer, clip);
        return;
      }
      case 'transform': {
        const [a, b, c, e, f, g] = p.matrix;
        if (!Number.isFinite(a + b + c + e + f + g) || Math.abs(a * e - b * c) < 1e-9) return;
        this.paint(p.paint, concat(p.matrix, pm), d, layer, mask);
        return;
      }
      case 'linear':
      case 'radial':
        this.gradient(p, pm, d, layer, mask);
        return;
      case 'composite': {
        const n = this.w * this.h * 4;
        const src = new Float32Array(n);
        const dst = new Float32Array(n);
        this.paint(p.source, pm, d, src, null);
        this.paint(p.backdrop, pm, d, dst, null);
        compose(p.mode, src, dst);
        for (let i = 0, o = 0; i < this.w * this.h; i++, o += 4) {
          const k = mask ? mask[i] : 1;
          if (k <= 0 || dst[o + 3] <= 0) continue;
          over(layer, o, dst[o] * k, dst[o + 1] * k, dst[o + 2] * k, dst[o + 3] * k);
        }
        return;
      }
    }
  }

  private gradient(p: Extract<ColorPaint, { kind: 'linear' | 'radial' }>, pm: Matrix, d: Matrix, layer: Float32Array, mask: Float32Array | null): void {
    const stops = p.stops;
    if (!stops.length) return;
    const [o0, o1] = [stops[0].offset, stops[stops.length - 1].offset];
    if (o1 - o0 < 1e-6) return this.paint({ kind: 'solid', color: stops[stops.length - 1].color }, pm, d, layer, mask);
    const inv = invert(concat(pm, d));
    if (!inv) return;
    // As the PDF writer does, the geometry moves to the first and last stops, which run 0 to 1.
    const at = (a: number, b: number, t: number) => a + (b - a) * t;
    const norm = stops.map((s): ColorStop => ({ offset: (s.offset - o0) / (o1 - o0), color: s.color }));
    const colors = norm.map((s) => [...this.rgb(s.color), s.color.alpha]);
    let position: (x: number, y: number) => number;
    if (p.kind === 'linear') {
      const x0 = at(p.x0, p.x1, o0);
      const y0 = at(p.y0, p.y1, o0);
      const dx = at(p.x0, p.x1, o1) - x0;
      const dy = at(p.y0, p.y1, o1) - y0;
      const len2 = dx * dx + dy * dy || 1;
      position = (x, y) => ((x - x0) * dx + (y - y0) * dy) / len2;
    } else {
      const cx0 = at(p.x0, p.x1, o0);
      const cy0 = at(p.y0, p.y1, o0);
      const r0 = Math.max(0, at(p.r0, p.r1, o0));
      const cdx = at(p.x0, p.x1, o1) - cx0;
      const cdy = at(p.y0, p.y1, o1) - cy0;
      const dr = Math.max(0, at(p.r0, p.r1, o1)) - r0;
      position = (x, y) => conical(x - cx0, y - cy0, cdx, cdy, r0, dr);
    }
    const w = this.w;
    for (let j = 0; j < this.h; j++) {
      for (let i = 0; i < w; i++) {
        const k = mask ? mask[j * w + i] : 1;
        if (k <= 0) continue;
        const px = i + 0.5;
        const py = j + 0.5;
        let t = position(inv[0] * px + inv[2] * py + inv[4], inv[1] * px + inv[3] * py + inv[5]);
        if (Number.isNaN(t)) continue;
        t = t < 0 ? 0 : t > 1 ? 1 : t;
        let s = 1;
        while (s < norm.length - 1 && norm[s].offset < t) s++;
        const a = colors[s - 1];
        const b = colors[s];
        const span = norm[s].offset - norm[s - 1].offset;
        const f = span > 1e-6 ? Math.max(0, Math.min(1, (t - norm[s - 1].offset) / span)) : 1;
        const alpha = (a[3] + (b[3] - a[3]) * f) * k;
        if (alpha <= 0) continue;
        over(layer, (j * w + i) * 4, (a[0] + (b[0] - a[0]) * f) * alpha, (a[1] + (b[1] - a[1]) * f) * alpha, (a[2] + (b[2] - a[2]) * f) * alpha, alpha);
      }
    }
  }
}

/**
 * The largest t at which the point (px, py), relative to the first circle's centre, lies on the
 * circle moving by (cdx, cdy) and growing by dr from radius r0, with a radius of at least zero.
 */
function conical(px: number, py: number, cdx: number, cdy: number, r0: number, dr: number): number {
  const a = cdx * cdx + cdy * cdy - dr * dr;
  const b = px * cdx + py * cdy + r0 * dr;
  const c = px * px + py * py - r0 * r0;
  if (Math.abs(a) < 1e-9) {
    if (Math.abs(b) < 1e-12) return NaN;
    const t = c / (2 * b);
    return r0 + t * dr >= 0 ? t : NaN;
  }
  const disc = b * b - a * c;
  if (disc < 0) return NaN;
  const sq = Math.sqrt(disc);
  const t1 = (b + sq) / a;
  const t2 = (b - sq) / a;
  const [hi, lo] = t1 > t2 ? [t1, t2] : [t2, t1];
  if (r0 + hi * dr >= 0) return hi;
  if (r0 + lo * dr >= 0) return lo;
  return NaN;
}

function over(layer: Float32Array, o: number, r: number, g: number, b: number, a: number): void {
  const keep = 1 - a;
  layer[o] = r + layer[o] * keep;
  layer[o + 1] = g + layer[o + 1] * keep;
  layer[o + 2] = b + layer[o + 2] * keep;
  layer[o + 3] = a + layer[o + 3] * keep;
}

/** Composite premultiplied `src` onto `dst` in place by COLR composite mode `mode`. */
function compose(mode: number, src: Float32Array, dst: Float32Array): void {
  const blend = mode >= 13 ? BLEND_MODES[mode - 13] : undefined;
  for (let o = 0; o < src.length; o += 4) {
    const sa = src[o + 3];
    const ba = dst[o + 3];
    let fs: number;
    let fb: number;
    switch (mode) {
      case COMPOSITE.CLEAR: [fs, fb] = [0, 0]; break;
      case COMPOSITE.SRC: [fs, fb] = [1, 0]; break;
      case COMPOSITE.DEST: [fs, fb] = [0, 1]; break;
      case COMPOSITE.DEST_OVER: [fs, fb] = [1 - ba, 1]; break;
      case COMPOSITE.SRC_IN: [fs, fb] = [ba, 0]; break;
      case COMPOSITE.DEST_IN: [fs, fb] = [0, sa]; break;
      case COMPOSITE.SRC_OUT: [fs, fb] = [1 - ba, 0]; break;
      case COMPOSITE.DEST_OUT: [fs, fb] = [0, 1 - sa]; break;
      case COMPOSITE.SRC_ATOP: [fs, fb] = [ba, 1 - sa]; break;
      case COMPOSITE.DEST_ATOP: [fs, fb] = [1 - ba, sa]; break;
      case COMPOSITE.XOR: [fs, fb] = [1 - ba, 1 - sa]; break;
      case COMPOSITE.PLUS:
        for (let k = 0; k < 4; k++) dst[o + k] = Math.min(1, src[o + k] + dst[o + k]);
        continue;
      default:
        [fs, fb] = [1, 1 - sa];
    }
    if (!blend) {
      for (let k = 0; k < 4; k++) dst[o + k] = src[o + k] * fs + dst[o + k] * fb;
      continue;
    }
    // Separable and non-separable blend modes, on straight colours.
    const cs = sa > 0 ? [src[o] / sa, src[o + 1] / sa, src[o + 2] / sa] : [0, 0, 0];
    const cb = ba > 0 ? [dst[o] / ba, dst[o + 1] / ba, dst[o + 2] / ba] : [0, 0, 0];
    const mixed = blendColor(blend, cb, cs);
    for (let k = 0; k < 3; k++) dst[o + k] = (1 - ba) * src[o + k] + (1 - sa) * dst[o + k] + sa * ba * mixed[k];
    dst[o + 3] = sa + ba - sa * ba;
  }
}

function blendColor(mode: string, b: number[], s: number[]): number[] {
  const sep = (f: (cb: number, cs: number) => number) => [f(b[0], s[0]), f(b[1], s[1]), f(b[2], s[2])];
  const hard = (cb: number, cs: number) => (cs <= 0.5 ? cb * 2 * cs : cb + (2 * cs - 1) - cb * (2 * cs - 1));
  switch (mode) {
    case 'Multiply': return sep((cb, cs) => cb * cs);
    case 'Screen': return sep((cb, cs) => cb + cs - cb * cs);
    case 'Overlay': return sep((cb, cs) => hard(cs, cb));
    case 'Darken': return sep(Math.min);
    case 'Lighten': return sep(Math.max);
    case 'ColorDodge': return sep((cb, cs) => (cb === 0 ? 0 : cs >= 1 ? 1 : Math.min(1, cb / (1 - cs))));
    case 'ColorBurn': return sep((cb, cs) => (cb >= 1 ? 1 : cs <= 0 ? 0 : 1 - Math.min(1, (1 - cb) / cs)));
    case 'HardLight': return sep(hard);
    case 'SoftLight': return sep((cb, cs) => {
      if (cs <= 0.5) return cb - (1 - 2 * cs) * cb * (1 - cb);
      const dd = cb <= 0.25 ? ((16 * cb - 12) * cb + 4) * cb : Math.sqrt(cb);
      return cb + (2 * cs - 1) * (dd - cb);
    });
    case 'Difference': return sep((cb, cs) => Math.abs(cb - cs));
    case 'Exclusion': return sep((cb, cs) => cb + cs - 2 * cb * cs);
    case 'Hue': return setLum(setSat(s, sat(b)), lum(b));
    case 'Saturation': return setLum(setSat(b, sat(s)), lum(b));
    case 'Color': return setLum(s, lum(b));
    case 'Luminosity': return setLum(b, lum(s));
    default: return s;
  }
}

const lum = (c: number[]) => 0.3 * c[0] + 0.59 * c[1] + 0.11 * c[2];
const sat = (c: number[]) => Math.max(...c) - Math.min(...c);

function setLum(c: number[], l: number): number[] {
  const d = l - lum(c);
  const out = c.map((v) => v + d);
  const L = lum(out);
  const n = Math.min(...out);
  const x = Math.max(...out);
  return out.map((v) => (n < 0 ? L + ((v - L) * L) / (L - n || 1) : x > 1 ? L + ((v - L) * (1 - L)) / (x - L || 1) : v));
}

function setSat(c: number[], s: number): number[] {
  const max = Math.max(...c);
  const min = Math.min(...c);
  if (max <= min) return [0, 0, 0];
  return c.map((v) => (v === max ? s : v === min ? 0 : ((v - min) * s) / (max - min)));
}

/** Composite a premultiplied layer with its top-left at (x0, y0), faded by `alpha`, through the clip. */
function composite(canvas: Canvas, r: Rendered, x0: number, y0: number, alpha: number): void {
  const c = canvas.clip;
  const { data, width } = canvas;
  const mask = c.mask;
  const ys = Math.max(y0, c.y0);
  const ye = Math.min(y0 + r.h, c.y1);
  const xs = Math.max(x0, c.x0);
  const xe = Math.min(x0 + r.w, c.x1);
  if (xs < xe) canvas.touch(xs, ys, xe, ye);
  for (let y = ys; y < ye; y++) {
    const mo = mask ? (y - c.my) * c.mw - c.mx : 0;
    for (let x = xs; x < xe; x++) {
      const s = ((y - y0) * r.w + (x - x0)) * 4;
      const sa = r.data[s + 3];
      if (!sa) continue;
      let k = alpha;
      if (x === c.x0) k *= c.e0;
      if (x === c.x1 - 1) k *= c.e1;
      if (mask) k = (k * mask[mo + x]) / 255;
      if (k <= 0) continue;
      const o = (y * width + x) * 3;
      const keep = 1 - (sa / 255) * k;
      data[o] = (r.data[s] * k + data[o] * keep + 0.5) | 0;
      data[o + 1] = (r.data[s + 1] * k + data[o + 1] * keep + 0.5) | 0;
      data[o + 2] = (r.data[s + 2] * k + data[o + 2] * keep + 0.5) | 0;
    }
  }
}
