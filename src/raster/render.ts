// Laid-out pages drawn to page images (PNG or JPEG), from the same drawing operations, in the same
// order and with the same graphics state, as the PDF writer emits them, anti-aliased as Poppler or
// MuPDF draws the PDF. In Poppler's style thin axis-aligned lines and rectangles snap to whole
// pixels, so rules and borders stay crisp.

import { loadOutlines, needsOutlineLoad } from '../fonts/outline.js';
import type { SfntFont } from '../fonts/sfnt.js';
import { ImageBudget, type ImageLimits } from '../image/budget.js';
import { sniffImage } from '../image/image.js';
import { JpegScanLimit } from '../image/jpeg-decoder.js';
import { encodeJpeg } from '../image/jpeg-encoder.js';
import { UnreadableJpeg, type Pixmap } from '../image/pixels.js';
import type { DrawOp, ImageOp, Page, PathSeg } from '../layout/types.js';
import type { Compression } from '../util/compression.js';
import { Canvas, type RasterStyle } from './canvas.js';
import { ColrPainter } from './color-glyph.js';
import { drawForm } from './form.js';
import { concat, flattenSegs, isAxisAligned, meanScale, Polylines, type Matrix } from './geometry.js';
import { ImagePainter } from './image.js';
import { pngFile, pngScanlines } from './png.js';
import { drawShade } from './shade.js';
import { dashCount, dashCoverage, dashPolylines, MAX_DASHES, strokeOutline, type Cap, type Join, type StrokeStyle } from './stroke.js';
import { TextPainter, type TextState } from './text.js';

type ImageRef = ImageOp['image'];

export type ImageFormat = 'png' | 'jpeg';

export interface RasterOptions {
  /** 'png' (default) or 'jpeg'. */
  format?: ImageFormat;
  /** Resolution in dots per inch (default 150). */
  dpi?: number;
  /** JPEG quality from 1 to 100 (default 90). */
  quality?: number;
  /** Indexes of the pages to draw, from 0 (default: every page). */
  pages?: readonly number[];
  imageLimits?: ImageLimits;
  /** Snap text sizes, baselines (from the page top) and line starts to this grid, in points. */
  paintGrid?: number;
  /** Whose anti-aliasing to follow (default 'poppler'). */
  style?: RasterStyle;
}

export interface PageImage {
  /** Page number, from 1. */
  page: number;
  /** Size in pixels. */
  width: number;
  height: number;
  /** The encoded image file. */
  data: Uint8Array<ArrayBuffer>;
}

/** Fast compression: three times quicker than the default for files about an eighth larger. */
const PNG_DEFLATE_LEVEL = 3;

/** Pages whose scanlines may wait for compression at once. */
const MAX_PENDING = 4;

/** The largest page image, in pixels (an A4 page at about 2,400 dpi). */
export const MAX_PAGE_PIXELS = 2 ** 29;
const MAX_DIMENSION = 65_535;

export function checkRasterOptions(opts: RasterOptions): void {
  const dpi = opts.dpi ?? 150;
  if (!Number.isFinite(dpi) || dpi <= 0 || dpi > 9600) throw new RangeError('Image resolution (dpi) must be a number above 0 and at most 9600');
  const q = opts.quality;
  if (q !== undefined && (!Number.isInteger(q) || q < 1 || q > 100)) throw new RangeError('Image quality must be an integer from 1 to 100');
  if (opts.format !== undefined && opts.format !== 'png' && opts.format !== 'jpeg') throw new RangeError("Image format must be 'png' or 'jpeg'");
}

/** Draw `pages` (or the chosen ones) to encoded page images. */
export async function renderImages(pages: Page[], compression: Compression, opts: RasterOptions = {}, warnings: string[] = []): Promise<PageImage[]> {
  checkRasterOptions(opts);
  const dpi = opts.dpi ?? 150;
  const scale = dpi / 72;
  const indexes = opts.pages ?? pages.map((_, i) => i);
  for (const i of indexes) if (!Number.isInteger(i) || i < 0 || i >= pages.length) throw new RangeError(`Page ${i + 1} is out of range (1-${pages.length})`);
  const sizes = indexes.map((i) => {
    const width = Math.max(1, Math.ceil(pages[i].width * scale - 1e-6));
    const height = Math.max(1, Math.ceil(pages[i].height * scale - 1e-6));
    if (width > MAX_DIMENSION || height > MAX_DIMENSION || width * height > MAX_PAGE_PIXELS) {
      throw new RangeError(`Page ${i + 1} would be ${width} × ${height} pixels at ${dpi} dpi, more than page images allow`);
    }
    return { width, height };
  });
  await preloadOutlines(indexes.map((i) => pages[i]));
  const style = opts.style ?? 'poppler';
  const painter = new PagePainter(compression, new ImageBudget(opts.imageLimits), warnings, opts.paintGrid);
  // The pixel limits apply to the pictures of each page: unlike the PDF writer, which embeds JPEGs
  // as they are, page images decode every picture, and a deck of photos passes the total.
  const pageBudget = () => new ImageBudget(opts.imageLimits);
  const out: PageImage[] = new Array(indexes.length);
  // Pages compressing while later ones are drawn, where compression can run alongside.
  const pending: Promise<void>[] = [];
  const spare: Uint8Array<ArrayBuffer>[] = [];
  let canvas: Canvas | null = null;
  try {
    for (const [k, i] of indexes.entries()) {
      const { width, height } = sizes[k];
      if (canvas && canvas.width === width && canvas.height === height) canvas.reset();
      else canvas = new Canvas(width, height, style);
      await painter.draw(pages[i], canvas, scale, pageBudget());
      if (opts.format === 'jpeg') {
        out[k] = { page: i + 1, width, height, data: encodeJpeg(width, height, canvas.data, opts.quality ?? 90, { lo: canvas.lo, hi: canvas.hi, dpi }) };
        continue;
      }
      const size = (width * 3 + 1) * height;
      // Scanline buffers go back to `spare` once compressed, for the next page of the same size.
      const reuse = spare.findIndex((b) => b.length === size);
      const rows = pngScanlines(width, height, canvas.data, canvas.lo, canvas.hi, reuse >= 0 ? spare.splice(reuse, 1)[0] : new Uint8Array(size));
      if (!compression.deflateZlibAsync) {
        out[k] = { page: i + 1, width, height, data: pngFile(width, height, compression.deflateZlib(rows, PNG_DEFLATE_LEVEL), dpi, compression.crc32) };
        spare.push(rows);
        continue;
      }
      const job = compression.deflateZlibAsync(rows, PNG_DEFLATE_LEVEL).then((deflated) => {
        out[k] = { page: i + 1, width, height, data: pngFile(width, height, deflated, dpi, compression.crc32) };
        if (spare.length <= MAX_PENDING) spare.push(rows);
      });
      // Rejections surface when the job is awaited below.
      job.catch(() => {});
      pending.push(job);
      if (pending.length >= MAX_PENDING) await pending.shift();
    }
  } finally {
    await Promise.allSettled(pending);
  }
  await Promise.all(pending);
  return out;
}

/** Outlines that a font keeps apart from its tables (the packaged CJK fonts), read up front. */
async function preloadOutlines(pages: Page[]): Promise<void> {
  const wanted = new Map<SfntFont, Set<number>>();
  for (const page of pages) {
    for (const layer of [page.behind, page.headerFooter, page.body, page.front]) {
      for (const op of layer) {
        if (op.kind !== 'glyphs' || !needsOutlineLoad(op.face.font)) continue;
        let gids = wanted.get(op.face.font);
        if (!gids) wanted.set(op.face.font, (gids = new Set()));
        for (const g of op.gids) gids.add(g);
      }
    }
  }
  for (const [font, gids] of wanted) await loadOutlines(font, gids);
}


const rgbOf = (hex: string): number => parseInt(hex.slice(0, 6), 16) || 0;

/** Whole pixels from a to b, rounded as PDF renderers adjust strokes: never thinner than one. */
function snap(a: number, b: number): [number, number] {
  const lo = Math.round(Math.min(a, b));
  let hi = Math.round(Math.max(a, b));
  if (hi === lo) hi = lo + 1;
  return [lo, hi];
}

interface Saved {
  ctm: Matrix;
  join: Join;
}

type Drawable = { form: string } | { pixmap: Pixmap };

class PagePainter {
  private readonly images: ImagePainter;
  private readonly forms = new Map<string, string | null>();
  /** Forms and failures by image; 'raster' for a picture the image painter decodes. */
  private readonly resolved = new Map<string, Drawable | null | 'raster'>();
  private readonly skipped = new Set<string>();
  private readonly text = new TextPainter(new ColrPainter());
  private readonly flat = new Polylines();
  private readonly outline = new Polylines();
  private readonly pieces = new Polylines();
  private svg?: typeof import('../pdf/svg.js');

  constructor(
    compression: Compression,
    budget: ImageBudget,
    private readonly warnings: string[],
    private readonly grid?: number,
  ) {
    this.images = new ImagePainter(compression, budget);
  }

  async draw(page: Page, canvas: Canvas, scale: number, budget: ImageBudget): Promise<void> {
    this.images.budget = budget;
    let ctm: Matrix = [scale, 0, 0, scale, 0, 0];
    const state: TextState = { join: 'miter' };
    const stack: Saved[] = [];
    for (const layer of [page.behind, page.headerFooter, page.body, page.front]) {
      for (const op of layer) {
        switch (op.kind) {
          case 'glyphs':
            this.text.draw(canvas, op, ctm, state, this.grid);
            break;
          case 'rect':
            this.rect(canvas, op, ctm);
            break;
          case 'line':
            this.line(canvas, op, ctm, state.join);
            break;
          case 'path':
            this.path(canvas, op, ctm, state);
            break;
          case 'shade':
            drawShade(canvas, op, ctm);
            break;
          case 'image':
            await this.image(canvas, op, ctm);
            break;
          case 'transform':
            stack.push({ ctm, join: state.join });
            canvas.save();
            ctm = concat(op.matrix, ctm);
            break;
          case 'clip':
            stack.push({ ctm, join: state.join });
            canvas.save();
            if (op.segs) this.clipSegs(canvas, op.segs, ctm, !op.nonzero);
            else this.clipRect(canvas, op.x, op.y, op.w, op.h, ctm);
            break;
          case 'restore': {
            const saved = stack.pop();
            if (saved) {
              ctm = saved.ctm;
              state.join = saved.join;
              canvas.restore();
            }
            break;
          }
          case 'link':
          case 'anchor':
          case 'float':
            break;
        }
      }
    }
  }

  private rect(canvas: Canvas, op: Extract<DrawOp, { kind: 'rect' }>, ctm: Matrix): void {
    if (!op.w || !op.h) return;
    const alpha = op.alpha !== undefined && op.alpha < 1 ? Math.max(0, op.alpha) : 1;
    if (canvas.style === 'poppler' && isAxisAligned(ctm)) {
      const [x0, x1] = snap(ctm[0] * op.x + ctm[4], ctm[0] * (op.x + op.w) + ctm[4]);
      const [y0, y1] = snap(ctm[3] * op.y + ctm[5], ctm[3] * (op.y + op.h) + ctm[5]);
      canvas.fillRect(x0, y0, x1, y1, rgbOf(op.fill), alpha);
      return;
    }
    if (isAxisAligned(ctm)) {
      const xa = ctm[0] * op.x + ctm[2] * op.y + ctm[4];
      const ya = ctm[1] * op.x + ctm[3] * op.y + ctm[5];
      const xb = ctm[0] * (op.x + op.w) + ctm[2] * (op.y + op.h) + ctm[4];
      const yb = ctm[1] * (op.x + op.w) + ctm[3] * (op.y + op.h) + ctm[5];
      canvas.fillArea(Math.min(xa, xb), Math.min(ya, yb), Math.max(xa, xb), Math.max(ya, yb), rgbOf(op.fill), alpha);
      return;
    }
    const flat = this.flat;
    flat.reset();
    flattenSegs(rectSegs(op.x, op.y, op.w, op.h), ctm, flat);
    canvas.fill(flat, false, rgbOf(op.fill), alpha);
  }

  private line(canvas: Canvas, op: Extract<DrawOp, { kind: 'line' }>, ctm: Matrix, join: Join): void {
    const scale = meanScale(ctm);
    const width = op.width > 0 ? op.width * scale : 1;
    const flat = this.flat;
    flat.reset();
    flat.moveTo(ctm[0] * op.x1 + ctm[2] * op.y1 + ctm[4], ctm[1] * op.x1 + ctm[3] * op.y1 + ctm[5]);
    flat.lineTo(ctm[0] * op.x2 + ctm[2] * op.y2 + ctm[4], ctm[1] * op.x2 + ctm[3] * op.y2 + ctm[5]);
    const cap: Cap = op.cap ?? 'butt';
    const dash = op.dash?.map((d) => d * scale);
    this.stroke(canvas, flat, { width, cap, join, miterLimit: 10, dash }, rgbOf(op.color), 1, true);
  }

  private path(canvas: Canvas, op: Extract<DrawOp, { kind: 'path' }>, ctm: Matrix, state: TextState): void {
    const flat = this.flat;
    flat.reset();
    flattenSegs(op.segs, ctm, flat);
    if (op.fill) {
      const alpha = op.fillAlpha !== undefined && op.fillAlpha < 1 ? Math.max(0, op.fillAlpha) : 1;
      const poppler = canvas.style === 'poppler';
      const rect = axisRect(op.segs, ctm, poppler);
      if (rect && poppler) {
        const [x0, x1] = snap(rect[0], rect[2]);
        const [y0, y1] = snap(rect[1], rect[3]);
        canvas.fillRect(x0, y0, x1, y1, rgbOf(op.fill), alpha);
      } else if (rect) canvas.fillArea(rect[0], rect[1], rect[2], rect[3], rgbOf(op.fill), alpha);
      else canvas.fill(flat, !op.nonzero, rgbOf(op.fill), alpha);
    }
    if (op.stroke) {
      const alpha = op.strokeAlpha !== undefined && op.strokeAlpha < 1 ? Math.max(0, op.strokeAlpha) : 1;
      const capValue = 'cap' in op ? op.cap : undefined;
      const joinValue = 'join' in op ? op.join : undefined;
      const cap: Cap = capValue === 'round' || capValue === 'square' ? capValue : 'butt';
      const join: Join = joinValue === 'miter' || joinValue === 'bevel' ? joinValue : 'round';
      const scale = meanScale(ctm);
      const sw = op.strokeWidth ?? 0.75;
      const straight = op.segs.every((s) => s.op !== 'C');
      this.stroke(canvas, flat, { width: sw > 0 ? sw * scale : 1, cap, join, miterLimit: 10, dash: op.dash?.map((d) => d * scale) }, rgbOf(op.stroke), alpha, straight);
      // The join outlives a stroke drawn outside a saved state, and stroked text inherits it.
      if (alpha === 1) state.join = join;
    }
  }

  /**
   * Stroke `flat` (device pixels). In Poppler's style thin strokes of only horizontal and vertical
   * segments snap to whole pixels, each segment a rectangle reaching over the joins.
   */
  private stroke(canvas: Canvas, flat: Polylines, style: StrokeStyle, rgb: number, alpha: number, straight: boolean): void {
    if (canvas.style === 'mupdf') {
      // MuPDF draws strokes thinner than a tenth of a pixel one pixel wide.
      if (style.width < 0.1) style = { ...style, width: 1 };
      // A straight horizontal or vertical line is a rectangle.
      if (flat.n === 2 && flat.starts.length === 1 && !flat.closed[0] && !style.dash && style.cap !== 'round') {
        const [ax, ay, bx, by] = [flat.xs[0], flat.ys[0], flat.xs[1], flat.ys[1]];
        const hw = style.width / 2;
        const ext = style.cap === 'square' ? hw : 0;
        if (ay === by && ax !== bx) return canvas.fillArea(Math.min(ax, bx) - ext, ay - hw, Math.max(ax, bx) + ext, ay + hw, rgb, alpha);
        if (ax === bx && ay !== by) return canvas.fillArea(ax - hw, Math.min(ay, by) - ext, ax + hw, Math.max(ay, by) + ext, rgb, alpha);
      }
    }
    if (style.dash && !(style.dash.some((d) => d > 0) && style.dash.every((d) => d >= 0 && Number.isFinite(d)))) style = { ...style, dash: undefined };
    if (style.dash && dashCount(flat, style.dash) > MAX_DASHES) {
      // Dashes far finer than a pixel average out to a fainter solid line.
      alpha *= dashCoverage(style.dash);
      style = { ...style, dash: undefined };
    }
    const hw = style.width / 2;
    if (canvas.style === 'poppler' && straight && style.cap !== 'round' && (style.width <= 3 || flat.n === 2) && allAxisAligned(flat)) {
      let path = flat;
      if (style.dash) {
        dashPolylines(flat, style.dash, style.phase ?? 0, this.outline);
        path = this.outline;
      }
      const pieces = this.pieces;
      pieces.reset();
      const capExt = style.cap === 'square' ? hw : 0;
      for (let s = 0; s < path.starts.length; s++) {
        const start = path.starts[s];
        const end = path.end(s);
        const closed = path.closed[s];
        const count = end - start;
        if (count < 2) continue;
        const segments = closed ? count : count - 1;
        for (let k = 0; k < segments; k++) {
          const a = start + k;
          const b = start + ((k + 1) % count);
          const [ax, ay, bx, by] = [path.xs[a], path.ys[a], path.xs[b], path.ys[b]];
          const firstSeg = !closed && k === 0;
          const lastSeg = !closed && k === segments - 1;
          const e0 = firstSeg ? capExt : hw;
          const e1 = lastSeg ? capExt : hw;
          let x0: number, x1: number, y0: number, y1: number;
          // Poppler's adjusted edges stop just short of the far pixel boundaries. Edges closer
          // than its 0.01-pixel snapping window both snap to the near one: a zero-width line.
          const across = (c: number): [number, number] => {
            if (style.width >= 0.01) {
              const [a, b] = snap(c - hw, c + hw);
              return [a, b - 0.01];
            }
            const a = Math.floor(c - hw + 0.5);
            return [a, a];
          };
          if (Math.abs(ay - by) <= Math.abs(ax - bx)) {
            const dir = bx >= ax ? 1 : -1;
            [x0, x1] = snap(ax - dir * e0, bx + dir * e1);
            x1 -= 0.01;
            [y0, y1] = across(ay);
          } else {
            const dir = by >= ay ? 1 : -1;
            [y0, y1] = snap(ay - dir * e0, by + dir * e1);
            y1 -= 0.01;
            [x0, x1] = across(ax);
          }
          if (ax === bx && ay === by && style.cap === 'butt') continue;
          pieces.moveTo(x0, y0);
          pieces.lineTo(x1, y0);
          pieces.lineTo(x1, y1);
          pieces.lineTo(x0, y1);
          pieces.close();
        }
      }
      canvas.fill(pieces, false, rgb, alpha);
      return;
    }
    const outline = this.outline;
    outline.reset();
    strokeOutline(flat, style, outline);
    canvas.fill(outline, false, rgb, alpha);
  }

  private clipRect(canvas: Canvas, x: number, y: number, w: number, h: number, ctm: Matrix): void {
    if (isAxisAligned(ctm)) {
      const xa = ctm[0] * x + ctm[4];
      const xb = ctm[0] * (x + w) + ctm[4];
      const ya = ctm[3] * y + ctm[5];
      const yb = ctm[3] * (y + h) + ctm[5];
      if (xa === xb || ya === yb) return canvas.clipRect(0, 0, -1, -1);
      canvas.clipRect(Math.min(xa, xb), Math.min(ya, yb), Math.max(xa, xb), Math.max(ya, yb));
      return;
    }
    this.clipSegs(canvas, rectSegs(x, y, w, h), ctm, false);
  }

  private clipSegs(canvas: Canvas, segs: readonly PathSeg[], ctm: Matrix, evenOdd: boolean): void {
    const rect = axisRect(segs, ctm, false);
    if (rect) {
      canvas.clipRect(rect[0], rect[1], rect[2], rect[3]);
      return;
    }
    const flat = this.flat;
    flat.reset();
    flattenSegs(segs, ctm, flat);
    canvas.clipPath(flat, evenOdd);
  }

  private form(ref: ImageRef): string | null {
    let form = this.forms.get(ref.id);
    if (form === undefined) {
      form = this.svg!.svgToPdfForm(ref.data);
      this.forms.set(ref.id, form);
    }
    return form;
  }

  /** What to draw for an image, preferring an SVG version as the PDF writer does. */
  private async resolve(ref: ImageRef): Promise<Drawable | null> {
    const key = ref.svg ? `${ref.id}\0${ref.svg.id}` : ref.id;
    const known = this.resolved.get(key);
    // Decoded pictures stay in the image painter's cache, which may drop them.
    if (known === 'raster') return this.pixmap(ref);
    if (known !== undefined) return known;
    let out: Drawable | null = null;
    const isSvg = !ref.dibBits && !ref.solid && (/\.svg$/i.test(ref.id) || sniffImage(ref.data) === 'svg');
    if (ref.svg || isSvg) this.svg ??= await import('../pdf/svg.js');
    const svgForm = ref.svg ? this.form(ref.svg) : null;
    if (svgForm !== null) out = { form: svgForm };
    else if (isSvg) {
      const form = this.form(ref);
      if (form === null) this.warnings.push(`Image ${ref.id} uses unsupported SVG features and was skipped`);
      else out = { form };
    } else {
      try {
        const pixmap = await this.images.load(ref);
        if (pixmap) {
          this.resolved.set(key, 'raster');
          return { pixmap };
        }
        this.warnings.push(`Image ${ref.id} uses an unsupported format and was skipped`);
      } catch (e) {
        if (e instanceof JpegScanLimit) this.warnings.push(`Image ${ref.id} is a JPEG with more than ${e.limit} scans and was skipped`);
        else if (e instanceof UnreadableJpeg) this.warnings.push(`Image ${ref.id} is a JPEG that could not be decoded and was skipped`);
        else if (e instanceof Error && e.message === 'Image pixel limit exceeded') {
          // Another page's budget may still have room for it.
          this.overLimit(ref);
          return null;
        } else throw e;
      }
    }
    this.resolved.set(key, out);
    return out;
  }

  /** A picture decoded before, decoded again if the cache dropped it. */
  private async pixmap(ref: ImageRef): Promise<Drawable | null> {
    try {
      const pixmap = await this.images.load(ref);
      return pixmap ? { pixmap } : null;
    } catch (e) {
      if (!(e instanceof Error && e.message === 'Image pixel limit exceeded')) throw e;
      this.overLimit(ref);
      return null;
    }
  }

  private overLimit(ref: ImageRef): void {
    if (this.skipped.has(ref.id)) return;
    this.skipped.add(ref.id);
    this.warnings.push(`Image ${ref.id} exceeds the image pixel limits and was skipped`);
  }

  private async image(canvas: Canvas, op: ImageOp, ctm: Matrix): Promise<void> {
    const drawable = await this.resolve(op.image);
    if (!drawable) return;
    const c = op.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const fw = 1 - c.l - c.r;
    const fh = 1 - c.t - c.b;
    if (fw <= 0 || fh <= 0) return;
    const W = op.w / fw;
    const H = op.h / fh;
    const m = concat([W, 0, 0, H, op.x - c.l * W, op.y - c.t * H], ctm);
    const a = 'alpha' in op ? op.alpha : undefined;
    const alpha = typeof a === 'number' && a < 1 ? Math.max(0, a) : 1;
    const cropped = c.l || c.t || c.r || c.b;
    if (cropped) {
      canvas.save();
      this.clipRect(canvas, op.x, op.y, op.w, op.h, ctm);
    }
    // A form's unit square runs up from its bottom edge; an image's rows run down from the top.
    if ('form' in drawable) drawForm(canvas, drawable.form, concat([1, 0, 0, -1, 0, 1], m));
    else this.images.draw(canvas, drawable.pixmap, m, alpha);
    if (cropped) canvas.restore();
  }
}

function rectSegs(x: number, y: number, w: number, h: number): PathSeg[] {
  return [
    { op: 'M', pts: [x, y] },
    { op: 'L', pts: [x + w, y] },
    { op: 'L', pts: [x + w, y + h] },
    { op: 'L', pts: [x, y + h] },
    { op: 'Z', pts: [] },
  ];
}

/** The device rectangle [x0, y0, x1, y1] a path traces, when it is one axis-aligned rectangle. */
function axisRect(segs: readonly PathSeg[], ctm: Matrix, fill: boolean): [number, number, number, number] | null {
  if (segs.length < 4 || segs.length > 6 || segs[0].op !== 'M' || !isAxisAligned(ctm)) return null;
  const pts: Array<[number, number]> = [];
  let closed = false;
  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    if (s.op === 'Z') {
      if (i !== segs.length - 1) return null;
      closed = true;
      continue;
    }
    if (s.op === 'C' || (i > 0 && s.op === 'M')) return null;
    pts.push([ctm[0] * s.pts[0] + ctm[4], ctm[3] * s.pts[1] + ctm[5]]);
  }
  const eq = (a: number, b: number) => Math.abs(a - b) < 1e-6;
  // Poppler snaps only the fills it hints as rectangles: four points left open, or five points
  // closed (closing adds the start point when the path does not end there). Any four-sided
  // clip counts as a rectangle.
  const last = pts[pts.length - 1];
  const returns = eq(last[0], pts[0][0]) && eq(last[1], pts[0][1]);
  const count = pts.length + (closed && !returns ? 1 : 0);
  if (fill ? (closed ? count !== 5 : count !== 4) : pts.length - (returns ? 1 : 0) !== 4) return null;
  const horizontalFirst = eq(pts[0][1], pts[1][1]) && eq(pts[1][0], pts[2][0]) && eq(pts[2][1], pts[3][1]) && eq(pts[3][0], pts[0][0]);
  const verticalFirst = eq(pts[0][0], pts[1][0]) && eq(pts[1][1], pts[2][1]) && eq(pts[2][0], pts[3][0]) && eq(pts[3][1], pts[0][1]);
  if (!horizontalFirst && !verticalFirst) return null;
  const xs = pts.slice(0, 4).map((p) => p[0]);
  const ys = pts.slice(0, 4).map((p) => p[1]);
  const x0 = Math.min(...xs);
  const x1 = Math.max(...xs);
  const y0 = Math.min(...ys);
  const y1 = Math.max(...ys);
  return x1 > x0 && y1 > y0 ? [x0, y0, x1, y1] : null;
}

function allAxisAligned(p: Polylines): boolean {
  const { xs, ys } = p;
  for (let s = 0; s < p.starts.length; s++) {
    const start = p.starts[s];
    const end = p.end(s);
    for (let i = start + 1; i <= end; i++) {
      if (i === end && !p.closed[s]) break;
      const j = i === end ? start : i;
      if (Math.abs(xs[j] - xs[i - 1]) > 1e-6 && Math.abs(ys[j] - ys[i - 1]) > 1e-6) return false;
    }
  }
  return true;
}
