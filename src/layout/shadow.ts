// Outer shadows (a:outerShdw). A blurred one is a picture of the shadow colour whose alpha is the
// shape's coverage blurred by a Gaussian of a third of the blur radius, as Word's and PowerPoint's
// PDFs draw it, reaching the blur radius past the shape on every side. How finely the picture is
// drawn, and whether it leaves out the shape itself or takes in its outline, is the application's:
// see SOFT_SHADOW in src/profile.ts.

import type { ImageRef } from '../ooxml/model.js';
import { SOFT_SHADOW } from '../profile.js';
import { Coverage } from '../raster/coverage.js';
import { flattenSegs, Polylines } from '../raster/geometry.js';
import { strokeOutline } from '../raster/stroke.js';
import type { DrawOp, PathSeg } from './types.js';

/** Larger shadows are drawn coarser, so one picture stays a few megabytes. */
const MAX_SHADOW_PIXELS = 4_000_000;
/** An outer shadow: offset (dx, dy) and blur radius in pt, colour, and scale about the shape's centre. */
export interface Shadow {
  dx: number;
  dy: number;
  blur: number;
  color: string;
  alpha: number;
  sx?: number;
  sy?: number;
}

/** The shape casting a shadow: the area it fills (null when unfilled) and its outline. */
export interface ShadowShape {
  fill: PathSeg[] | null;
  line?: { segs: PathSeg[]; width: number; dash?: readonly number[] } | null;
}

/** Shadow pictures by their shape, size and colour, so identical shadows share one image. */
export type ShadowCache = Map<string, ImageRef>;

/**
 * Ops drawing `s`, the shadow of `shape`, behind it. A sharp shadow is the shape again in the
 * shadow's colour; a blurred one a picture, whose id follows from its content so that identical
 * shadows, in `cache` or not, are one image in the output.
 */
export function shadowOps(shape: ShadowShape, s: Shadow, cache?: ShadowCache): DrawOp[] {
  if (!(s.alpha > 0)) return [];
  const sx = s.sx ?? 1;
  const sy = s.sy ?? 1;
  if (!(sx > 0 && sy > 0)) return [];
  const fill = shape.fill?.length ? shape.fill : null;
  const outline = shape.line && shape.line.width > 0 && shape.line.segs.length ? shape.line : null;
  const box = bounds(fill ?? outline?.segs ?? []);
  if (!box) return [];
  const cx = (box[0] + box[2]) / 2;
  const cy = (box[1] + box[3]) / 2;
  const move = (segs: PathSeg[]): PathSeg[] => segs.map((q) => ({ op: q.op, pts: q.pts.map((v, i) => (i % 2 === 0 ? cx + (v - cx) * sx + s.dx : cy + (v - cy) * sy + s.dy)) }));
  if (!(s.blur > 0)) {
    const ops: DrawOp[] = [];
    if (fill) ops.push({ kind: 'path', segs: move(fill), fill: s.color, ...(s.alpha < 1 && { fillAlpha: s.alpha }), stroke: null });
    if (outline) ops.push({ kind: 'path', segs: move(outline.segs), fill: null, stroke: s.color, strokeWidth: outline.width, ...(s.alpha < 1 && { strokeAlpha: s.alpha }), ...(outline.dash && { dash: [...outline.dash] }) });
    return ops;
  }
  const moved = fill && move(fill);
  // An unfilled shape's shadow is its outline's.
  const line = outline && (SOFT_SHADOW.outline || !fill) ? { ...outline, segs: move(outline.segs) } : null;
  const reach = bounds([...(moved ?? []), ...(line?.segs ?? [])]);
  if (!reach) return [];
  const pad = s.blur + (line ? line.width / 2 : 0);
  const x0 = reach[0] - pad;
  const y0 = reach[1] - pad;
  const w = reach[2] - reach[0] + 2 * pad;
  const h = reach[3] - reach[1] + 2 * pad;
  let scale = SOFT_SHADOW.dpi / 72;
  if (w * h * scale * scale > MAX_SHADOW_PIXELS) scale = Math.sqrt(MAX_SHADOW_PIXELS / (w * h));
  const pw = Math.max(1, Math.ceil(w * scale));
  const ph = Math.max(1, Math.ceil(h * scale));
  const local = (segs: PathSeg[]) => segs.map((q) => q.op + q.pts.map((v, i) => ((v - (i % 2 ? y0 : x0)) * scale).toFixed(2)).join(',')).join(' ');
  const key = `${s.color}:${s.alpha.toFixed(4)}:${s.blur.toFixed(3)}:${scale.toFixed(4)}:${pw}x${ph}:${moved ? local(moved) : ''}|${line ? `${line.width.toFixed(3)}:${line.dash?.join(',') ?? ''}:${local(line.segs)}` : ''}`;
  let image = cache?.get(key);
  if (!image) {
    image = shadowImage(moved, line, x0, y0, scale, pw, ph, s.color, s.alpha, (s.blur / 3) * scale, `shadow-${hash(key)}.bmp`);
    cache?.set(key, image);
  }
  // The picture's whole pixels, a little past the bounds on the right and below.
  const iw = pw / scale;
  const ih = ph / scale;
  const picture: DrawOp = { kind: 'image', x: x0, y: y0, w: iw, h: ih, image, rotation: 0, flipH: false, flipV: false };
  if (!SOFT_SHADOW.clip || !fill) return [picture];
  const outside: PathSeg[] = [
    ...fill,
    { op: 'M', pts: [x0, y0] },
    { op: 'L', pts: [x0 + iw, y0] },
    { op: 'L', pts: [x0 + iw, y0 + ih] },
    { op: 'L', pts: [x0, y0 + ih] },
    { op: 'Z', pts: [] },
  ];
  return [{ kind: 'clip', x: x0, y: y0, w: iw, h: ih, segs: outside }, picture, { kind: 'restore' }];
}

function bounds(segs: PathSeg[]): [number, number, number, number] | null {
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const q of segs) {
    for (let i = 0; i + 1 < q.pts.length; i += 2) {
      minX = Math.min(minX, q.pts[i]);
      maxX = Math.max(maxX, q.pts[i]);
      minY = Math.min(minY, q.pts[i + 1]);
      maxY = Math.max(maxY, q.pts[i + 1]);
    }
  }
  return maxX >= minX && maxY >= minY && (maxX > minX || maxY > minY) ? [minX, minY, maxX, maxY] : null;
}

/** Two 32-bit FNV-1a hashes of `s`, as 16 hex digits. */
function hash(s: string): string {
  let a = 0x811c9dc5;
  let b = 0x050c5d1f;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    a = Math.imul(a ^ c, 0x01000193);
    b = Math.imul(b ^ c, 0x5bd1e995);
  }
  return (a >>> 0).toString(16).padStart(8, '0') + (b >>> 0).toString(16).padStart(8, '0');
}

function shadowImage(
  fill: PathSeg[] | null,
  line: { segs: PathSeg[]; width: number; dash?: readonly number[] } | null,
  x0: number,
  y0: number,
  scale: number,
  pw: number,
  ph: number,
  color: string,
  alpha: number,
  sigma: number,
  id: string,
): ImageRef {
  const m: [number, number, number, number, number, number] = [scale, 0, 0, scale, -x0 * scale, -y0 * scale];
  const area = new Polylines();
  if (fill) flattenSegs(fill, m, area);
  const outline = new Polylines();
  if (line) {
    const poly = new Polylines();
    flattenSegs(line.segs, m, poly);
    strokeOutline(poly, { width: line.width * scale, cap: 'butt', join: 'round', miterLimit: 10, dash: line.dash?.map((d) => d * scale) }, outline);
  }

  // Between the left and right edges of a box, all columns have the same
  // coverage. Keep a full kernel window there and restore its middle pixel
  // across the omitted columns after both passes. Curves and diagonals
  // touching the middle shrink or eliminate the band.
  const middle = pw / 2;
  let left = 0, right = pw;
  for (const path of [area, outline]) {
    for (let s = 0; s < path.starts.length && left < right; s++) {
      const start = path.starts[s], end = path.end(s);
      for (let i = start, prev = end - 1; i < end; prev = i++) {
        if (path.ys[i] === path.ys[prev]) continue;
        const lo = Math.min(path.xs[i], path.xs[prev]);
        const hi = Math.max(path.xs[i], path.xs[prev]);
        if (hi < middle) left = Math.max(left, Math.ceil(hi) + 2);
        else if (lo > middle) right = Math.min(right, Math.floor(lo) - 1);
        else { right = left; break; }
      }
    }
  }
  const radius = Math.max(1, Math.ceil(3 * sigma));
  const gap = Math.max(0, right - left - (2 * radius + 1));
  const split = left + radius + 1;
  const width = pw - gap;
  const mask = new Float32Array(width * ph);
  const coverage = new Coverage();
  if (fill && coverage.begin(0, 0, pw, ph, split, gap)) {
    coverage.addPolylines(area);
    coverage.sweep(true, (y, from, to, row) => mask.set(row.subarray(from, to), y * width + from));
  }
  if (line && coverage.begin(0, 0, pw, ph, split, gap)) {
    coverage.addPolylines(outline);
    coverage.sweep(false, (y, from, to, row) => {
      for (let x = from, o = y * width + from; x < to; x++, o++) if (row[x] > mask[o]) mask[o] = row[x];
    });
  }
  if (sigma > 0) gaussianBlurInPlace(mask, width, ph, sigma);
  const opacity = new Uint8Array(pw * ph);
  for (let y = 0; y < ph; y++) {
    let dest = y * pw;
    for (let x = 0; x < width; x++) {
      if (gap && x === split) {
        opacity.fill(opacity[dest - 1], dest, dest + gap);
        dest += gap;
      }
      opacity[dest++] = Math.round(255 * alpha * Math.min(1, mask[y * width + x]));
    }
  }
  return { id, data: opacity, solid: { width: pw, height: ph, rgb: parseInt(color.slice(0, 6), 16) || 0 } };
}

/** Blur the coverage mask in place, reaching three sigmas each way with zero coverage past its edges. */
export function gaussianBlurInPlace(mask: Float32Array, w: number, h: number, sigma: number): void {
  const r = Math.max(1, Math.ceil(3 * sigma));
  const kernel = new Float32Array(2 * r + 1);
  let total = 0;
  for (let i = -r; i <= r; i++) total += kernel[i + r] = Math.exp(-(i * i) / (2 * sigma * sigma));
  for (let i = 0; i < kernel.length; i++) kernel[i] /= total;
  // Retain the two edges of an identical row run and one interior row. Every
  // omitted row has exactly the same vertical kernel window as that interior row.
  const copies = new Int32Array(h);
  const across = new Float32Array(w);
  let rows = 0;
  for (let y = 0; y < h;) {
    const start = y * w;
    let end = y + 1;
    while (end < h && equalRow(mask, start, end * w, w)) end++;
    blurLine(mask.subarray(start, start + w), across, kernel, false);
    const kept = Math.min(end - y, kernel.length);
    for (let i = 0; i < kept; i++) mask.set(across, (rows + i) * w);
    if (kept < end - y) copies[rows + r] = end - y - kept;
    rows += kept;
    y = end;
  }
  const input = new Float32Array(rows);
  const output = new Float32Array(rows);
  for (let x = 0; x < w; x++) {
    let changed = x === 0;
    for (let y = 0, pos = x; y < rows; y++, pos += w) {
      const sample = mask[pos];
      if (sample !== input[y]) {
        input[y] = sample;
        changed = true;
      }
    }
    if (changed) blurLine(input, output, kernel, true);
    for (let y = 0, pos = x; y < rows; y++, pos += w) mask[pos] = output[y];
  }
  if (rows < h) {
    let dest = h;
    for (let y = rows - 1; y >= 0; y--) {
      const row = mask.subarray(y * w, (y + 1) * w);
      for (let i = 0; i <= copies[y]; i++) mask.set(row, --dest * w);
    }
  }
}

function equalRow(mask: Float32Array, a: number, b: number, width: number): boolean {
  for (let i = 0; i < width; i++) if (mask[a + i] !== mask[b + i]) return false;
  return true;
}

function blurLine(input: Float32Array, output: Float32Array, kernel: Float32Array, roundEachTap: boolean): void {
  const radius = (kernel.length - 1) / 2;
  for (let start = 0; start < input.length;) {
    let end = start + 1;
    while (end < input.length && input[end] === input[start]) end++;
    for (let x = start; x < end;) {
      const first = Math.max(0, x - radius);
      const last = Math.min(input.length - 1, x + radius);
      let sum = 0;
      for (let i = first; i <= last; i++) {
        sum += input[i] * kernel[i - x + radius];
        // The vertical pass accumulates into Float32; the horizontal pass rounds once.
        if (roundEachTap) sum = Math.fround(sum);
      }
      // Every full kernel window inside this constant run sees the same samples.
      if (x === start + radius && x + radius + 1 < end) {
        output.fill(sum, x, end - radius);
        x = end - radius;
      } else {
        output[x++] = sum;
      }
    }
    start = end;
  }
}
