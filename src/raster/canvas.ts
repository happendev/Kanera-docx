// An opaque RGB page in device pixels, painted by coverage: solid fills, coverage masks (glyphs)
// and per-pixel shaders (images, gradients), all through a stack of anti-aliased clips.

import { Coverage } from './coverage.js';
import { Supersampled } from './supersample.js';
import type { Polylines } from './geometry.js';

/**
 * Where painting may reach: whole-pixel bounds, and an optional coverage mask inside them.
 *
 * Rectangular clips work as Poppler's do: glyphs reach every pixel the rectangle touches, while
 * vector paint and images are clipped across in quarter pixels, so the first and last columns
 * can be partly covered (`e0`, `e1`, with the vector-edge gamma applied).
 */
export interface Clip {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  /** End of the columns glyphs may reach (at most x1). */
  gx1: number;
  /** The rectangle's first and last quarter-pixel columns, inclusive. */
  q0: number;
  q1: number;
  /** Coverage of columns x0 and x1 - 1 by vector paint. */
  e0: number;
  e1: number;
  /** Coverage 0–255, `mw` wide, with its top-left pixel at (mx, my). */
  mask: Uint8Array | null;
  mx: number;
  my: number;
  mw: number;
}

/**
 * Whose anti-aliasing a page follows. 'poppler': Splash's 4 × 4 samples with a 1.5 gamma on vector
 * edges, thin strokes snapped to whole pixels and glyph baselines on whole pixels. 'mupdf': plain
 * area coverage, strokes where they fall, rectangular clips widened to whole pixels and glyphs
 * placed on quarter pixels both ways.
 */
export type RasterStyle = 'poppler' | 'mupdf';

/**
 * Source colours for pixels x0..x1-1 of row y: red, green and blue (0–255) and alpha (0–1),
 * interleaved from out[0].
 */
export type Shader = (y: number, x0: number, x1: number, out: Float32Array) => void;

/**
 * Coverage of anti-aliased vector edges, raised to the power 1.5 as Poppler's Splash rasteriser
 * does: edges look lighter and thinner than plain area coverage. Glyphs and images are exempt.
 */
const GAMMA = (() => {
  const t = new Float32Array(257);
  for (let i = 0; i <= 256; i++) t[i] = (i / 256) ** 1.5;
  return t;
})();
export const vectorCoverage = (c: number): number => GAMMA[(c * 256 + 0.5) | 0];

export class Canvas {
  readonly data: Uint8Array;
  /** Per row, the pixels [lo, hi) that may differ from white (lo ≥ hi: the row is white). */
  readonly lo: Int32Array;
  readonly hi: Int32Array;
  clip: Clip;
  private readonly stack: Clip[] = [];
  private readonly coverage = new Coverage();
  /** In Poppler's style, vector paint is sampled as Splash samples it; glyphs get their exact area. */
  private readonly samples = new Supersampled();
  private readonly poppler: boolean;
  private source = new Float32Array(0);

  constructor(
    readonly width: number,
    readonly height: number,
    readonly style: RasterStyle = 'poppler',
  ) {
    this.poppler = style === 'poppler';
    this.data = new Uint8Array(width * height * 3).fill(255);
    this.lo = new Int32Array(height).fill(width);
    this.hi = new Int32Array(height);
    this.clip = fullClip(width, height);
  }

  /** Clear to white with no clip, for the next page. */
  reset(): void {
    const { data, lo, hi, width } = this;
    for (let y = 0; y < this.height; y++) {
      if (lo[y] < hi[y]) data.fill(255, (y * width + lo[y]) * 3, (y * width + hi[y]) * 3);
      lo[y] = width;
      hi[y] = 0;
    }
    this.stack.length = 0;
    this.clip = fullClip(this.width, this.height);
  }

  /** Note that pixels [x0, x1) × [y0, y1) may have been painted a colour other than white. */
  touch(x0: number, y0: number, x1: number, y1: number): void {
    const { lo, hi } = this;
    for (let y = y0; y < y1; y++) {
      if (x0 < lo[y]) lo[y] = x0;
      if (x1 > hi[y]) hi[y] = x1;
    }
  }

  /** Bounds of `p` in whole pixels, inside the clip; null when nothing would be painted. */
  private window(p: Polylines): [number, number, number, number] | null {
    const b = p.bounds();
    if (!b) return null;
    const c = this.clip;
    const x0 = Math.max(c.x0, Math.floor(b[0]));
    const y0 = Math.max(c.y0, Math.floor(b[1]));
    const x1 = Math.min(c.x1, Math.ceil(b[2]) + 1);
    const y1 = Math.min(c.y1, Math.ceil(b[3]) + 1);
    return x0 < x1 && y0 < y1 ? [x0, y0, x1, y1] : null;
  }

  save(): void {
    this.stack.push(this.clip);
  }

  restore(): void {
    const clip = this.stack.pop();
    if (clip) this.clip = clip;
  }

  /** Clip to the device-space rectangle [x0, x1] × [y0, y1]; paired with restore(). */
  clipRect(x0: number, y0: number, x1: number, y1: number): void {
    const c = this.clip;
    if (!this.poppler) {
      // MuPDF clips to every pixel the rectangle reaches into (by more than its 1/17 × 1/15 samples).
      const ex1 = Math.min(c.x1, Math.ceil(Math.floor(x1 * 17) / 17));
      this.clip = { ...c, x0: Math.max(c.x0, Math.floor(x0)), y0: Math.max(c.y0, Math.floor(y0)), x1: ex1, y1: Math.min(c.y1, Math.ceil(Math.floor(y1 * 15) / 15)), gx1: Math.min(c.gx1, ex1) };
      return;
    }
    const q0 = Math.max(c.q0, Math.floor(x0 * 4));
    const q1 = Math.min(c.q1, Math.floor(x1 * 4));
    this.clip = edged({
      ...c,
      x0: Math.max(c.x0, Math.floor(q0 / 4)),
      y0: Math.max(c.y0, Math.floor(y0)),
      x1: Math.min(c.x1, Math.floor(q1 / 4) + 1),
      y1: Math.min(c.y1, Math.ceil(y1)),
      gx1: Math.min(c.gx1, Math.ceil(x1)),
      q0,
      q1,
    });
  }

  /** Clip to the inside of `p`, anti-aliased; paired with restore(). */
  clipPath(p: Polylines, evenOdd: boolean): void {
    const c = this.clip;
    const win = this.window(p);
    if (!win) {
      this.clip = { ...c, x1: c.x0, y1: c.y0, gx1: c.x0 };
      return;
    }
    const [x0, y0, x1, y1] = win;
    const mw = x1 - x0;
    const mask = new Uint8Array(mw * (y1 - y0));
    let tx0 = x1;
    let ty0 = y1;
    let tx1 = x0;
    let ty1 = y0;
    const parent = c.mask;
    const poppler = this.poppler;
    const cov = poppler ? this.samples : this.coverage;
    cov.begin(x0, y0, x1, y1);
    cov.addPolylines(p);
    cov.sweep(evenOdd, (y, from, to, row) => {
      const o = (y - y0) * mw;
      const po = parent ? (y - c.my) * c.mw - c.mx + x0 : 0;
      let any = false;
      for (let x = from; x < to; x++) {
        let v = (poppler ? vectorCoverage(row[x]) : row[x]) * 255;
        if (parent) v = (v * parent[po + x]) / 255;
        // The mask takes in the enclosing rectangle's partly covered columns.
        if (x + x0 === c.x0) v *= c.e0;
        if (x + x0 === c.x1 - 1) v *= c.e1;
        const b = (v + 0.5) | 0;
        if (b) {
          mask[o + x] = b;
          any = true;
          if (x < tx0 - x0) tx0 = x + x0;
          if (x >= tx1 - x0) tx1 = x + x0 + 1;
        }
      }
      if (any) {
        if (y < ty0) ty0 = y;
        if (y >= ty1) ty1 = y + 1;
      }
    });
    const x1m = Math.max(tx0, tx1);
    this.clip = { ...c, x0: tx0, y0: ty0, x1: x1m, y1: Math.max(ty0, ty1), gx1: Math.min(c.gx1, x1m), e0: 1, e1: 1, mask, mx: x0, my: y0, mw };
  }

  /** Fill `p` with one colour (0xRRGGBB) at `alpha`: `vector` paint, or a glyph's outline. */
  fill(p: Polylines, evenOdd: boolean, rgb: number, alpha: number, vector = true): void {
    if (!(alpha > 0)) return;
    const win = this.window(p);
    if (!win) return;
    const [x0, y0, x1, y1] = win;
    // Poppler samples vector paint on its own grid and lightens its edges.
    const sampled = vector && this.poppler;
    const cov = sampled ? this.samples : this.coverage;
    cov.begin(x0, y0, x1, y1);
    cov.addPolylines(p);
    const r = (rgb >> 16) & 255;
    const g = (rgb >> 8) & 255;
    const b = rgb & 255;
    const { data, width, lo, hi } = this;
    const clip = this.clip;
    const mask = clip.mask;
    // White on a white page stays white, so it never widens the painted span.
    const white = rgb === 0xffffff;
    // Glyphs stop at the clip's glyph columns; vector paint is faded across its edge columns.
    const end = vector ? x1 - x0 : Math.min(x1, clip.gx1) - x0;
    const l = vector && clip.e0 < 1 ? clip.x0 - x0 : -1;
    const r1 = vector && clip.e1 < 1 ? clip.x1 - 1 - x0 : -1;
    const opaque = !mask && alpha >= 0.9995;
    // Translucent paint over runs of whole pixels goes through a table of the blend.
    let table: Uint8Array | null = null;
    cov.sweep(evenOdd, (y, from, to, row) => {
      if (to > end) to = end;
      if (from >= to) return;
      if (!white) {
        if (x0 + from < lo[y]) lo[y] = x0 + from;
        if (x0 + to > hi[y]) hi[y] = x0 + to;
      }
      const mo = mask ? (y - clip.my) * clip.mw - clip.mx + x0 : 0;
      let o = (y * width + x0 + from) * 3;
      for (let x = from; x < to; x++, o += 3) {
        const c = row[x];
        // Rows can span gaps, such as the inside of an outline.
        if (c === 0) continue;
        if (c === 1 && opaque && x !== l && x !== r1) {
          // A run of whole pixels inside the shape.
          let e = x + 1;
          while (e < to && row[e] === 1 && e !== r1) e++;
          solid(data, o, e - x, r, g, b);
          o += (e - x - 1) * 3;
          x = e - 1;
          continue;
        }
        if (c === 1 && !mask && x !== l && x !== r1 && alpha > 0.0005) {
          let e = x + 1;
          while (e < to && row[e] === 1 && e !== r1) e++;
          if (e - x >= BLEND_RUN) {
            table ??= blendTable(r, g, b, alpha);
            blendRun(data, o, e - x, table);
            o += (e - x - 1) * 3;
            x = e - 1;
            continue;
          }
        }
        let a = (sampled && c < 1 ? vectorCoverage(c) : c) * alpha;
        if (x === l) a *= clip.e0;
        if (x === r1) a *= clip.e1;
        if (mask) a = (a * mask[mo + x]) / 255;
        if (a <= 0.0005) continue;
        if (a >= 0.9995) {
          data[o] = r;
          data[o + 1] = g;
          data[o + 2] = b;
        } else {
          data[o] = (data[o] + (r - data[o]) * a + 0.5) | 0;
          data[o + 1] = (data[o + 1] + (g - data[o + 1]) * a + 0.5) | 0;
          data[o + 2] = (data[o + 2] + (b - data[o + 2]) * a + 0.5) | 0;
        }
      }
    });
  }

  /** Fill whole pixels [x0, x1) × [y0, y1) with one colour at `alpha`. */
  fillRect(x0: number, y0: number, x1: number, y1: number, rgb: number, alpha: number): void {
    if (!(alpha > 0)) return;
    const c = this.clip;
    x0 = Math.max(x0, c.x0);
    y0 = Math.max(y0, c.y0);
    x1 = Math.min(x1, c.x1);
    y1 = Math.min(y1, c.y1);
    if (x0 >= x1 || y0 >= y1) return;
    if (rgb !== 0xffffff) this.touch(x0, y0, x1, y1);
    const r = (rgb >> 16) & 255;
    const g = (rgb >> 8) & 255;
    const b = rgb & 255;
    const { data, width } = this;
    const mask = c.mask;
    const edges = (x0 === c.x0 && c.e0 < 1) || (x1 === c.x1 && c.e1 < 1);
    if (!mask && !edges && alpha >= 0.9995) {
      // One row, then copies of it.
      const first = (y0 * width + x0) * 3;
      const n = (x1 - x0) * 3;
      solid(data, first, x1 - x0, r, g, b);
      for (let y = y0 + 1; y < y1; y++) data.copyWithin((y * width + x0) * 3, first, first + n);
      return;
    }
    if (!mask && !edges && alpha > 0.0005 && (x1 - x0) * (y1 - y0) >= BLEND_RUN) {
      const table = blendTable(r, g, b, alpha);
      for (let y = y0; y < y1; y++) blendRun(data, (y * width + x0) * 3, x1 - x0, table);
      return;
    }
    for (let y = y0; y < y1; y++) {
      let o = (y * width + x0) * 3;
      const mo = mask ? (y - c.my) * c.mw - c.mx : 0;
      for (let x = x0; x < x1; x++, o += 3) {
        let a = mask ? (alpha * mask[mo + x]) / 255 : alpha;
        if (x === c.x0) a *= c.e0;
        if (x === c.x1 - 1) a *= c.e1;
        if (a >= 0.9995) {
          data[o] = r;
          data[o + 1] = g;
          data[o + 2] = b;
          continue;
        }
        if (a <= 0.0005) continue;
        data[o] = (data[o] + (r - data[o]) * a + 0.5) | 0;
        data[o + 1] = (data[o + 1] + (g - data[o + 1]) * a + 0.5) | 0;
        data[o + 2] = (data[o + 2] + (b - data[o + 2]) * a + 0.5) | 0;
      }
    }
  }

  /**
   * Fill the device rectangle [x0, x1] × [y0, y1] with one colour at `alpha`, each pixel by the
   * area it covers: what fill() paints for the rectangle in MuPDF's style, faster.
   */
  fillArea(x0: number, y0: number, x1: number, y1: number, rgb: number, alpha: number): void {
    if (!(alpha > 0) || !(x0 < x1 && y0 < y1)) return;
    const c = this.clip;
    const px0 = Math.max(c.x0, Math.floor(x0));
    const py0 = Math.max(c.y0, Math.floor(y0));
    const px1 = Math.min(c.x1, Math.ceil(x1));
    const py1 = Math.min(c.y1, Math.ceil(y1));
    if (px0 >= px1 || py0 >= py1) return;
    if (rgb !== 0xffffff) this.touch(px0, py0, px1, py1);
    const r = (rgb >> 16) & 255;
    const g = (rgb >> 8) & 255;
    const b = rgb & 255;
    const { data, width } = this;
    const mask = c.mask;
    // Columns px0..px1-1 are whole but for the first and last, which the rectangle may cover in part.
    const first = Math.min(px0 + 1, x1) - Math.max(px0, x0);
    const last = px1 - 1 === px0 ? first : Math.min(px1, x1) - (px1 - 1);
    const opaque = !mask && alpha >= 0.9995;
    for (let y = py0; y < py1; y++) {
      const cy = Math.min(y + 1, y1) - Math.max(y, y0);
      const mo = mask ? (y - c.my) * c.mw - c.mx : 0;
      let o = (y * width + px0) * 3;
      for (let x = px0; x < px1; x++, o += 3) {
        const cx = x === px0 ? first : x === px1 - 1 ? last : 1;
        if (cx === 1 && cy === 1 && opaque && x !== px1 - 1) {
          // The row's whole pixels, in one go.
          const n = px1 - 1 - x;
          solid(data, o, n, r, g, b);
          o += (n - 1) * 3;
          x += n - 1;
          continue;
        }
        let a = cx * cy * alpha;
        if (mask) a = (a * mask[mo + x]) / 255;
        if (a <= 0.0005) continue;
        if (a >= 0.9995) {
          data[o] = r;
          data[o + 1] = g;
          data[o + 2] = b;
        } else {
          data[o] = (data[o] + (r - data[o]) * a + 0.5) | 0;
          data[o + 1] = (data[o + 1] + (g - data[o + 1]) * a + 0.5) | 0;
          data[o + 2] = (data[o + 2] + (b - data[o + 2]) * a + 0.5) | 0;
        }
      }
    }
  }

  /**
   * Composite a coverage mask (0–255, `mw` × `mh`) with its top-left pixel at (dx, dy): a glyph,
   * or with `vector` a vector-drawn shape (clipped as vector paint is).
   */
  blit(mask: Uint8Array, mw: number, mh: number, dx: number, dy: number, rgb: number, alpha: number, vector = false): void {
    const c = this.clip;
    const x0 = Math.max(dx, c.x0);
    const y0 = Math.max(dy, c.y0);
    const x1 = Math.min(dx + mw, vector ? c.x1 : c.gx1);
    const y1 = Math.min(dy + mh, c.y1);
    if (x0 >= x1 || y0 >= y1 || !(alpha > 0)) return;
    if (rgb !== 0xffffff) this.touch(x0, y0, x1, y1);
    const r = (rgb >> 16) & 255;
    const g = (rgb >> 8) & 255;
    const b = rgb & 255;
    const { data, width } = this;
    const clipMask = c.mask;
    const k = alpha / 255;
    const l = vector && c.e0 < 1 ? c.x0 : -1;
    const r1 = vector && c.e1 < 1 ? c.x1 - 1 : -1;
    for (let y = y0; y < y1; y++) {
      let o = (y * width + x0) * 3;
      let m = (y - dy) * mw + (x0 - dx);
      const co = clipMask ? (y - c.my) * c.mw - c.mx : 0;
      for (let x = x0; x < x1; x++, o += 3, m++) {
        const v = mask[m];
        if (!v) continue;
        let a = v * k;
        if (x === l) a *= c.e0;
        if (x === r1) a *= c.e1;
        if (clipMask) a = (a * clipMask[co + x]) / 255;
        if (a >= 0.9995) {
          data[o] = r;
          data[o + 1] = g;
          data[o + 2] = b;
        } else if (a > 0.0005) {
          data[o] = (data[o] + (r - data[o]) * a + 0.5) | 0;
          data[o + 1] = (data[o + 1] + (g - data[o + 1]) * a + 0.5) | 0;
          data[o + 2] = (data[o + 2] + (b - data[o + 2]) * a + 0.5) | 0;
        }
      }
    }
  }

  /**
   * Paint `shader`'s colours at `alpha`: inside `p` when given, else over the whole-pixel `region`,
   * both within the clip.
   */
  shade(p: Polylines | null, evenOdd: boolean, region: [number, number, number, number] | null, shader: Shader, alpha: number): void {
    if (!(alpha > 0)) return;
    const c = this.clip;
    const poppler = this.poppler;
    const paintRow = (y: number, from: number, to: number, cov: Float32Array | null, covX0: number) => {
      if (from >= to) return;
      this.touch(from, y, to, y + 1);
      const n = to - from;
      if (this.source.length < n * 4) this.source = new Float32Array(Math.max(n * 4, this.source.length * 2));
      const src = this.source;
      shader(y, from, to, src);
      const { data, width } = this;
      const mask = c.mask;
      const mo = mask ? (y - c.my) * c.mw - c.mx : 0;
      let o = (y * width + from) * 3;
      for (let x = from, s = 0; x < to; x++, o += 3, s += 4) {
        let a = src[s + 3] * alpha;
        if (cov) a *= poppler ? vectorCoverage(cov[x - covX0]) : cov[x - covX0];
        if (x === c.x0) a *= c.e0;
        if (x === c.x1 - 1) a *= c.e1;
        if (mask) a = (a * mask[mo + x]) / 255;
        if (a <= 0.0005) continue;
        if (a >= 0.9995) {
          data[o] = src[s];
          data[o + 1] = src[s + 1];
          data[o + 2] = src[s + 2];
        } else {
          data[o] = (data[o] + (src[s] - data[o]) * a + 0.5) | 0;
          data[o + 1] = (data[o + 1] + (src[s + 1] - data[o + 1]) * a + 0.5) | 0;
          data[o + 2] = (data[o + 2] + (src[s + 2] - data[o + 2]) * a + 0.5) | 0;
        }
      }
    };
    if (p) {
      const win = this.window(p);
      if (!win) return;
      let [x0, y0, x1, y1] = win;
      if (region) {
        x0 = Math.max(x0, region[0]);
        y0 = Math.max(y0, region[1]);
        x1 = Math.min(x1, region[2]);
        y1 = Math.min(y1, region[3]);
        if (x0 >= x1 || y0 >= y1) return;
      }
      const cov = poppler ? this.samples : this.coverage;
      cov.begin(x0, y0, x1, y1);
      cov.addPolylines(p);
      cov.sweep(evenOdd, (y, from, to, row) => paintRow(y, x0 + from, x0 + to, row, x0));
      return;
    }
    const x0 = Math.max(c.x0, region ? region[0] : c.x0);
    const y0 = Math.max(c.y0, region ? region[1] : c.y0);
    const x1 = Math.min(c.x1, region ? region[2] : c.x1);
    const y1 = Math.min(c.y1, region ? region[3] : c.y1);
    for (let y = y0; y < y1; y++) paintRow(y, x0, x1, null, 0);
  }
}

function fullClip(width: number, height: number): Clip {
  return { x0: 0, y0: 0, x1: width, y1: height, gx1: width, q0: 0, q1: width * 4 - 1, e0: 1, e1: 1, mask: null, mx: 0, my: 0, mw: 0 };
}

/** `c` with the vector coverage of its first and last columns worked out from its quarter pixels. */
function edged(c: Clip): Clip {
  const part = (x: number) => {
    const n = Math.min(c.q1, x * 4 + 3) - Math.max(c.q0, x * 4) + 1;
    return n >= 4 ? 1 : n <= 0 ? 0 : vectorCoverage(n / 4);
  };
  c.e0 = c.x0 < c.x1 ? part(c.x0) : 1;
  c.e1 = c.x0 < c.x1 ? part(c.x1 - 1) : 1;
  return c;
}

/** Pixels in a run worth blending through a table rather than one by one. */
const BLEND_RUN = 32;

/**
 * Each channel value (red, green, blue in turn, 256 apiece) blended toward (r, g, b) at `a`, as
 * the per-pixel blend rounds it.
 */
function blendTable(r: number, g: number, b: number, a: number): Uint8Array {
  const t = new Uint8Array(768);
  for (let v = 0; v < 256; v++) {
    t[v] = (v + (r - v) * a + 0.5) | 0;
    t[256 + v] = (v + (g - v) * a + 0.5) | 0;
    t[512 + v] = (v + (b - v) * a + 0.5) | 0;
  }
  return t;
}

/** Blend `n` pixels from byte `o` through `table`. */
function blendRun(data: Uint8Array, o: number, n: number, table: Uint8Array): void {
  for (const end = o + n * 3; o < end; o += 3) {
    data[o] = table[data[o]];
    data[o + 1] = table[256 + data[o + 1]];
    data[o + 2] = table[512 + data[o + 2]];
  }
}

/** Paint `n` pixels from byte `o` one colour: a few by hand, long runs by doubling copies. */
function solid(data: Uint8Array, o: number, n: number, r: number, g: number, b: number): void {
  const first = Math.min(n, 8);
  for (let i = 0, p = o; i < first; i++, p += 3) {
    data[p] = r;
    data[p + 1] = g;
    data[p + 2] = b;
  }
  for (let done = first; done < n; ) {
    const k = Math.min(done, n - done);
    data.copyWithin(o + done * 3, o, o + k * 3);
    done += k;
  }
}
