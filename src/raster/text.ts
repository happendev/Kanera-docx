// Glyph runs drawn as the PDF writer draws them: outlines filled in the text colour, synthetic
// bold as fill and stroke, outline text as a stroke, synthetic italic as a slant. Glyphs at a
// plain scale come from a cache of coverage masks at quarter-pixel offsets.
//
// Plain fills are placed as Poppler places glyphs: the origin's x in whole quarter pixels (whole
// pixels for large fonts) and its baseline on a whole pixel. Text that is also stroked is drawn
// as vector paths at its exact position, fill then stroke, with the vector-edge gamma. In MuPDF's
// style glyphs keep their exact outlines and are placed on quarter pixels both ways (half or whole
// pixels for larger text).

import type { FontFace } from '../fonts/registry.js';
import { glyphPath } from '../fonts/outline.js';
import type { SfntFont } from '../fonts/sfnt.js';
import type { GlyphOp } from '../layout/types.js';
import { vectorCoverage, type Canvas, type RasterStyle } from './canvas.js';
import { Coverage } from './coverage.js';
import { Supersampled } from './supersample.js';
import { concat, flattenOutline, meanScale, Polylines, type Matrix } from './geometry.js';
import { strokeOutline, type Join } from './stroke.js';
import { SYNTHETIC_BOLD_STROKE } from '../profile.js';

const ITALIC_SKEW = 0.2126; // tan(12°), the slant Word uses for synthetic italics
/** Glyphs larger than this many pixels per em are drawn as paths rather than cached. */
const MAX_CACHED_EM = 256;
const MAX_CACHE_BYTES = 32 * 1024 * 1024;
const SUBPIXELS = 4;
/** Fonts whose bounding box is taller than this many pixels are placed on whole pixels. */
const FRACTION_LIMIT = 50;
/** MuPDF draws glyphs larger than this many pixels per em as paths, where they fall. */
const MUPDF_GLYPH_LIMIT = 256;

/** Masks of one face at one scale and of one kind, by glyph and quarter-pixel offset. */
type MaskSet = Map<number, GlyphMask | null>;

interface GlyphMask {
  /** Offset of the mask's top-left pixel from the glyph origin's pixel. */
  x: number;
  y: number;
  w: number;
  h: number;
  data: Uint8Array;
  /** Each row's covered columns, as [first, past the last) pairs: [0, 0) for an empty row. */
  spans: Int32Array;
}

/** Draws colour glyphs (fonts with COLR tables), where a renderer supports them. */
export interface ColorGlyphs {
  /** Whether glyph `gid` of `font` has a colour glyph. */
  has(font: SfntFont, gid: number): boolean;
  /** Draw it with `m` taking font units to device pixels, in text colour `fg` (0xRRGGBB). */
  draw(canvas: Canvas, font: SfntFont, gid: number, m: Matrix, fg: number, alpha: number): void;
}

export interface TextState {
  /** Line join for stroked text: the PDF writer leaves whatever join earlier paths set. */
  join: Join;
}

export class TextPainter {
  /** Mask sets by face, then by scale and kind. */
  private readonly cache = new Map<FontFace, Map<string, MaskSet>>();
  private bytes = 0;
  private readonly coverage = new Coverage();
  private readonly samples = new Supersampled();
  private readonly flat = new Polylines();
  private readonly outline = new Polylines();

  private canvas!: Canvas;
  private style: RasterStyle = 'poppler';

  constructor(private readonly colors?: ColorGlyphs) {}

  /** Draw `op` onto `canvas`; masks are kept for later pages too. */
  draw(canvas: Canvas, op: GlyphOp, ctm: Matrix, state: TextState, grid?: number): void {
    this.canvas = canvas;
    if (canvas.style !== this.style) {
      this.cache.clear();
      this.bytes = 0;
      this.style = canvas.style;
    }
    const mupdf = this.style === 'mupdf';
    if (!op.gids.length) return;
    const font = op.face.font;
    const snap = (v: number) => (grid ? Math.round(v / grid) * grid : v);
    const size = grid ? Math.max(grid, snap(op.size)) : op.size;
    if (!(size > 0) || !Number.isFinite(size)) return;
    const th = op.scaleX || 1;
    const skew = op.syntheticItalic ? ITALIC_SKEW : 0;
    const k = size / font.unitsPerEm;
    const colors = this.colors;
    const colored = colors !== undefined && op.gids.some((g) => colors.has(font, g));
    const fill = colored || !op.outline;
    const stroke = !colored && (op.outline || op.syntheticBold);
    const strokeWidth = op.size * (op.outline ? 0.02 : SYNTHETIC_BOLD_STROKE) * meanScale(ctm);
    const alpha = op.alpha !== undefined && op.alpha < 1 ? Math.max(0, op.alpha) : 1;
    if (alpha <= 0) return;
    const rgb = parseInt(op.color.slice(0, 6), 16) || 0;
    const x0 = op.x + op.xs[0] + (grid && op.lineX !== undefined ? snap(op.lineX) - op.lineX : 0);
    const y = grid ? snap(op.y) : op.y;
    const [a, b, c, d, e, f] = ctm;
    const base: Matrix = [k * th, 0, k * skew, -k, 0, 0];
    const device = concat(base, ctm);
    const fractional = fontHeight(font.bbox, device) <= FRACTION_LIMIT;
    // MuPDF's steps per pixel for glyph origins, by the text's size in pixels per em.
    const em = Math.sqrt(Math.abs(device[0] * device[3] - device[1] * device[2])) * font.unitsPerEm;
    const steps = em >= 48 ? 1 : em >= 24 ? 2 : 4;
    const cacheable = Math.abs(b) < 1e-9 && Math.abs(c) < 1e-9 && a > 0 && d > 0 && em <= MAX_CACHED_EM;
    // Cached masks share the run's scale, so each kind's set is looked up once.
    const scale: Matrix = [a * k * th, 0, a * k * skew, -d * k, 0, 0];
    const fills = cacheable ? this.maskSet(op.face, scale, stroke ? 'v' : 'f') : null;
    const strokes = cacheable && stroke ? this.maskSet(op.face, scale, `s${strokeWidth}${state.join}`) : null;
    // MuPDF rounds glyph origins to the nearest of `steps` positions per pixel.
    const half = 0.5 / steps;
    const quarters = SUBPIXELS / steps;
    for (let i = 0; i < op.gids.length; i++) {
      const gid = op.gids[i];
      const x = x0 + op.xs[i] - op.xs[0];
      if (colored && colors.has(font, gid)) {
        colors.draw(this.canvas, font, gid, concat([k * th, 0, k * skew, -k, x, y], ctm), rgb, alpha);
        continue;
      }
      if (fills) {
        const ox = a * x + e;
        const oy = d * y + f;
        if (!strokes) {
          if (mupdf) {
            const tx = ox + half;
            const ty = oy + half;
            const ix = Math.floor(tx);
            const iy = Math.floor(ty);
            const mask = this.mask(fills, font, gid, scale, Math.floor((tx - ix) * steps) * quarters, Math.floor((ty - iy) * steps) * quarters, false);
            if (mask) this.canvas.blit(mask.data, mask.w, mask.h, ix + mask.x, iy + mask.y, rgb, alpha, false, mask.spans);
            continue;
          }
          const ix = Math.floor(ox);
          const fx = fractional ? Math.min(SUBPIXELS - 1, Math.floor((ox - ix) * SUBPIXELS)) : 0;
          const mask = this.mask(fills, font, gid, scale, fx, 0, false);
          if (mask) this.canvas.blit(mask.data, mask.w, mask.h, ix + mask.x, Math.floor(oy) + mask.y, rgb, alpha, false, mask.spans);
          continue;
        }
        let ix = Math.floor(ox);
        let iy = Math.floor(oy);
        let fx = Math.round((ox - ix) * SUBPIXELS);
        let fy = Math.round((oy - iy) * SUBPIXELS);
        if (fx === SUBPIXELS) {
          ix++;
          fx = 0;
        }
        if (fy === SUBPIXELS) {
          iy++;
          fy = 0;
        }
        if (fill) {
          const mask = this.mask(fills, font, gid, scale, fx, fy, true);
          if (mask) this.canvas.blit(mask.data, mask.w, mask.h, ix + mask.x, iy + mask.y, rgb, alpha, true, mask.spans);
        }
        const mask = this.mask(strokes, font, gid, scale, fx, fy, true, strokeWidth, state.join);
        if (mask) this.canvas.blit(mask.data, mask.w, mask.h, ix + mask.x, iy + mask.y, rgb, alpha, true, mask.spans);
        continue;
      }
      const m = concat([k * th, 0, k * skew, -k, x, y], ctm);
      const outline = glyphPath(font, gid);
      if (!outline.cmds.length) continue;
      if (fill) {
        // Plain fills land where the glyph rasteriser puts them; stroked text stays exact.
        let at: Matrix = m;
        if (!stroke && !mupdf) {
          const ix = Math.floor(m[4]);
          const fx = fractional ? Math.floor((m[4] - ix) * SUBPIXELS) / SUBPIXELS : 0;
          at = [m[0], m[1], m[2], m[3], ix + fx, Math.floor(m[5])];
        } else if (!stroke && em <= MUPDF_GLYPH_LIMIT) {
          const [ix, fx] = mupdfPlace(m[4], steps);
          const [iy, fy] = mupdfPlace(m[5], steps);
          at = [m[0], m[1], m[2], m[3], ix + fx / SUBPIXELS, iy + fy / SUBPIXELS];
        }
        const flat = this.flat;
        flat.reset();
        flattenOutline(outline.cmds, outline.pts, at, flat);
        this.canvas.fill(flat, false, rgb, alpha, stroke);
      }
      if (stroke) {
        const flat = this.flat;
        flat.reset();
        flattenOutline(outline.cmds, outline.pts, m, flat);
        const path = this.outline;
        path.reset();
        strokeOutline(flat, { width: strokeWidth, cap: 'butt', join: state.join, miterLimit: 10 }, path);
        this.canvas.fill(path, false, rgb, alpha);
      }
    }
  }

  /** The masks of `face` scaled by `m`: glyph-rasteriser fills ('f'), vector fills ('v') or strokes. */
  private maskSet(face: FontFace, m: Matrix, kind: string): MaskSet {
    let byFace = this.cache.get(face);
    if (!byFace) this.cache.set(face, (byFace = new Map()));
    const key = `${m[0]}|${m[2]}|${m[3]}|${kind}`;
    let set = byFace.get(key);
    if (!set) byFace.set(key, (set = new Map()));
    return set;
  }

  /** A glyph's mask from `set`, rendered when missing, offset by (fx, fy) quarter pixels. */
  private mask(set: MaskSet, font: SfntFont, gid: number, m: Matrix, fx: number, fy: number, vector: boolean, strokeWidth = 0, join: Join = 'miter'): GlyphMask | null {
    const key = gid * 16 + fx * 4 + fy;
    let mask = set.get(key);
    if (mask !== undefined) return mask;
    mask = this.render(font, gid, [m[0], m[1], m[2], m[3], fx / SUBPIXELS, fy / SUBPIXELS], vector, strokeWidth, join);
    if (mask) {
      this.bytes += mask.data.length + mask.spans.length * 4 + 64;
      if (this.bytes > MAX_CACHE_BYTES) {
        // Start afresh; the sets stay, emptied, for the runs that hold them.
        for (const byFace of this.cache.values()) for (const s of byFace.values()) s.clear();
        this.bytes = mask.data.length + mask.spans.length * 4 + 64;
      }
    }
    set.set(key, mask);
    return mask;
  }

  /**
   * A glyph's fill coverage, or with `strokeWidth` its stroke's: as the glyph rasteriser draws it,
   * or as `vector` paint (with Poppler's edge gamma in its style).
   */
  private render(font: SfntFont, gid: number, m: Matrix, vector: boolean, strokeWidth: number, join: Join): GlyphMask | null {
    const outline = glyphPath(font, gid);
    if (!outline.cmds.length) return null;
    const poppler = this.style === 'poppler';
    let path = this.flat;
    path.reset();
    // Poppler's glyph fills follow FreeType's fixed-point outline at a whole pixel size; vector
    // paths, and MuPDF's glyphs (FreeType at 1,024 pixels per em), stay exact.
    if (vector || !poppler) flattenOutline(outline.cmds, outline.pts, m, path);
    else flattenOutline(outline.cmds, freetypePoints(font, outline.pts, m), IDENTITY, path);
    if (strokeWidth) {
      const stroke = this.outline;
      stroke.reset();
      strokeOutline(path, { width: strokeWidth, cap: 'butt', join, miterLimit: 10 }, stroke);
      path = stroke;
    }
    const bounds = path.bounds();
    if (!bounds) return null;
    const x0 = Math.floor(bounds[0]);
    const y0 = Math.floor(bounds[1]);
    const w = Math.ceil(bounds[2]) - x0 + 1;
    const h = Math.ceil(bounds[3]) - y0 + 1;
    if (w <= 0 || h <= 0) return null;
    const data = new Uint8Array(w * h);
    const sampled = vector && poppler;
    const cov = sampled ? this.samples : this.coverage;
    cov.begin(x0, y0, x0 + w, y0 + h);
    cov.addPolylines(path);
    cov.sweep(false, (y, from, to, row) => {
      const o = (y - y0) * w;
      // FreeType keeps coverage in 1/256ths and tops out at 255.
      if (sampled) for (let x = from; x < to; x++) data[o + x] = (vectorCoverage(row[x]) * 255 + 0.5) | 0;
      else if (vector) for (let x = from; x < to; x++) data[o + x] = (row[x] * 255 + 0.5) | 0;
      else for (let x = from; x < to; x++) data[o + x] = Math.min(255, (row[x] * 256) | 0);
    });
    const spans = new Int32Array(h * 2);
    for (let y = 0; y < h; y++) {
      const o = y * w;
      let a = 0;
      let b = w;
      while (a < w && !data[o + a]) a++;
      while (b > a && !data[o + b - 1]) b--;
      spans[y * 2] = a;
      spans[y * 2 + 1] = b;
    }
    return { x: x0, y: y0, w, h, data, spans };
  }
}

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/**
 * A glyph origin coordinate split as MuPDF splits it: rounded to the nearest of `steps` positions
 * per pixel, as the whole pixel and the offset in quarter pixels.
 */
function mupdfPlace(v: number, steps: number): [number, number] {
  const t = v + 0.5 / steps;
  const i = Math.floor(t);
  return [i, Math.floor((t - i) * steps) * (SUBPIXELS / steps)];
}

/** FreeType's fixed-point product a × b / 65536, rounded half away from zero. */
function mulFix(a: number, b: number): number {
  const p = a * b;
  return p < 0 ? -Math.floor((-p + 0x8000) / 0x10000) : Math.floor((p + 0x8000) / 0x10000);
}

/**
 * Outline points `pts` (font units) in device pixels as FreeType places them for Poppler, which
 * sets the nearest whole pixel size and stretches it to the true size with a 16.16 matrix: a
 * TrueType outline is scaled to 1/64 pixels at the whole size, then transformed in fixed point;
 * a CFF outline lands on 1/64 pixels.
 */
function freetypePoints(font: SfntFont, pts: Float32Array, m: Matrix): Float64Array {
  const upm = font.unitsPerEm;
  // Pixels per em, y up as FreeType has it.
  const xx = m[0] * upm;
  const xy = m[2] * upm;
  const yx = -m[1] * upm;
  const yy = -m[3] * upm;
  const out = new Float64Array(pts.length);
  if (font.isCFF) {
    for (let i = 0; i < pts.length; i += 2) {
      const x = pts[i];
      const y = pts[i + 1];
      out[i] = Math.round(((xx * x + xy * y) / upm) * 64) / 64 + m[4];
      out[i + 1] = -Math.round(((yx * x + yy * y) / upm) * 64) / 64 + m[5];
    }
    return out;
  }
  const size = Math.max(1, Math.round(Math.hypot(xy, yy)));
  const scale = Math.floor((size * 64 * 65536 + (upm >> 1)) / upm);
  const fxx = Math.trunc((xx / size) * 65536);
  const fxy = Math.trunc((xy / size) * 65536);
  const fyx = Math.trunc((yx / size) * 65536);
  const fyy = Math.trunc((yy / size) * 65536);
  // Implied midpoints of quadratic curves can fall on half units; they round to 1/64 too.
  const scaled = (v: number) => (Number.isInteger(v) ? mulFix(v, scale) : Math.round((v * scale) / 65536));
  for (let i = 0; i < pts.length; i += 2) {
    const sx = scaled(pts[i]);
    const sy = scaled(pts[i + 1]);
    out[i] = (mulFix(sx, fxx) + mulFix(sy, fxy)) / 64 + m[4];
    out[i + 1] = -(mulFix(sx, fyx) + mulFix(sy, fyy)) / 64 + m[5];
  }
  return out;
}

/** Height in pixels of a font's bounding box drawn through `m`, with Poppler's 3-pixel margin. */
function fontHeight(bbox: readonly number[], m: Matrix): number {
  let lo = Infinity;
  let hi = -Infinity;
  for (const [x, y] of [[bbox[0], bbox[1]], [bbox[0], bbox[3]], [bbox[2], bbox[1]], [bbox[2], bbox[3]]]) {
    const v = m[1] * x + m[3] * y;
    if (v < lo) lo = v;
    if (v > hi) hi = v;
  }
  return Math.trunc(hi) - Math.trunc(lo) + 3;
}
