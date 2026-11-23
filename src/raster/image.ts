// Raster images on the page: decoded once, resampled to the pixels they cover and composited
// through the clip. Upright images are scaled as Poppler scales them: over every pixel they touch,
// shrunk by averaging whole source pixels, enlarged by interpolation (or by repeating pixels from
// four times up, or along an axis that shrinks the other). Rotated ones sample through the
// inverse transform inside their outline.

import type { ImageBudget } from '../image/budget.js';
import { decodePixels, type Pixmap } from '../image/pixels.js';
import type { ImageOp } from '../layout/types.js';
import type { Compression } from '../util/compression.js';
import type { Canvas } from './canvas.js';
import { invert, isAxisAligned, Polylines, type Matrix } from './geometry.js';

/** Premultiplied RGBA pixels. */
interface Premultiplied {
  width: number;
  height: number;
  data: Uint8Array;
}

type ImageRef = ImageOp['image'];

const MAX_CACHED_BYTES = 192 * 1024 * 1024;

/** Least recently used entries are dropped first once their bytes pass the limit. */
class Lru<K, V> {
  private readonly map = new Map<K, { value: V; bytes: number }>();
  private bytes = 0;
  constructor(private readonly limit: number) {}
  get(key: K): V | undefined {
    const hit = this.map.get(key);
    if (!hit) return undefined;
    this.map.delete(key);
    this.map.set(key, hit);
    return hit.value;
  }
  set(key: K, value: V, bytes: number): void {
    const old = this.map.get(key);
    if (old) {
      this.bytes -= old.bytes;
      this.map.delete(key);
    }
    this.map.set(key, { value, bytes });
    this.bytes += bytes;
    for (const [k, v] of this.map) {
      if (this.bytes <= this.limit || k === key) break;
      this.map.delete(k);
      this.bytes -= v.bytes;
    }
  }
}

export class ImagePainter {
  /** Decoded images by their bytes and colour adjustment (null: unsupported). */
  private readonly decoded = new WeakMap<Uint8Array, Map<string, Pixmap | null>>();
  private readonly sources = new Lru<Pixmap, true>(MAX_CACHED_BYTES);
  private readonly scaled = new Lru<string, Premultiplied>(MAX_CACHED_BYTES / 2);
  private readonly ids = new WeakMap<Pixmap, number>();
  private nextId = 1;
  private readonly outline = new Polylines();

  constructor(
    private readonly compression: Compression,
    private readonly budget: ImageBudget,
  ) {}

  /** Decode a raster image (or fetch it decoded); null when its format is unsupported. */
  async load(ref: ImageRef): Promise<Pixmap | null> {
    const adjust = ref.adjust ? `${ref.adjust.slope}:${ref.adjust.offset}` : '';
    let byAdjust = this.decoded.get(ref.data);
    if (!byAdjust) this.decoded.set(ref.data, (byAdjust = new Map()));
    let px = byAdjust.get(adjust);
    if (px === undefined) {
      px = await decodePixels(ref.data, this.compression, this.budget, ref.dibBits, ref.adjust);
      byAdjust.set(adjust, px);
      if (px) this.sources.set(px, true, px.data.length);
    }
    return px;
  }

  /** Draw `px` into the unit square mapped by `m` (unit square → device pixels, row 0 at v = 0). */
  draw(canvas: Canvas, px: Pixmap, m: Matrix, alpha: number): void {
    if (!(alpha > 0)) return;
    if (isAxisAligned(m)) this.drawUpright(canvas, px, m, alpha);
    else this.drawTransformed(canvas, px, m, alpha);
  }

  private id(px: Pixmap): number {
    let id = this.ids.get(px);
    if (id === undefined) this.ids.set(px, (id = this.nextId++));
    return id;
  }

  private drawUpright(canvas: Canvas, px: Pixmap, m: Matrix, alpha: number): void {
    const [a, , , d, e, f] = m;
    const x0 = Math.floor(Math.min(e, e + a));
    const x1 = Math.floor(Math.max(e, e + a)) + 1;
    const y0 = Math.floor(Math.min(f, f + d));
    const y1 = Math.floor(Math.max(f, f + d)) + 1;
    const dw = x1 - x0;
    const dh = y1 - y0;
    const c = canvas.clip;
    const vx0 = Math.max(x0, c.x0);
    const vy0 = Math.max(y0, c.y0);
    const vx1 = Math.min(x1, c.x1);
    const vy1 = Math.min(y1, c.y1);
    if (vx0 >= vx1 || vy0 >= vy1) return;
    const flipX = a < 0;
    const flipY = d < 0;
    const key = `${this.id(px)}:${dw}x${dh}:${flipX ? 1 : 0}${flipY ? 1 : 0}:${vx0 - x0},${vy0 - y0},${vx1 - x0},${vy1 - y0}`;
    let region = this.scaled.get(key);
    if (!region) {
      region = resample(px, dw, dh, vx0 - x0, vy0 - y0, vx1 - x0, vy1 - y0, flipX, flipY);
      this.scaled.set(key, region, region.data.length);
    }
    composite(canvas, region, vx0, vy0, alpha);
  }

  private drawTransformed(canvas: Canvas, px: Pixmap, m: Matrix, alpha: number): void {
    const inv = invert(m);
    if (!inv) return;
    // Sample from a copy reduced to about the image's size on the page, so detail averages out.
    const tw = Math.max(1, Math.min(px.width, Math.ceil(Math.hypot(m[0], m[1]))));
    const th = Math.max(1, Math.min(px.height, Math.ceil(Math.hypot(m[2], m[3]))));
    const key = `${this.id(px)}:${tw}x${th}:full`;
    let src = this.scaled.get(key);
    if (!src) {
      src = resample(px, tw, th, 0, 0, tw, th, false, false);
      this.scaled.set(key, src, src.data.length);
    }
    const shape = this.outline;
    shape.reset();
    shape.moveTo(m[4], m[5]);
    shape.lineTo(m[0] + m[4], m[1] + m[5]);
    shape.lineTo(m[0] + m[2] + m[4], m[1] + m[3] + m[5]);
    shape.lineTo(m[2] + m[4], m[3] + m[5]);
    shape.close();
    const { width: sw, height: sh, data } = src;
    canvas.shade(shape, false, null, (y, from, to, out) => {
      const cy = y + 0.5;
      for (let x = from, o = 0; x < to; x++, o += 4) {
        const cx = x + 0.5;
        const u = (inv[0] * cx + inv[2] * cy + inv[4]) * sw - 0.5;
        const v = (inv[1] * cx + inv[3] * cy + inv[5]) * sh - 0.5;
        const ui = Math.floor(u);
        const vi = Math.floor(v);
        const fu = u - ui;
        const fv = v - vi;
        const u0 = ui < 0 ? 0 : ui >= sw ? sw - 1 : ui;
        const u1 = ui + 1 < 0 ? 0 : ui + 1 >= sw ? sw - 1 : ui + 1;
        const v0 = vi < 0 ? 0 : vi >= sh ? sh - 1 : vi;
        const v1 = vi + 1 < 0 ? 0 : vi + 1 >= sh ? sh - 1 : vi + 1;
        const p00 = (v0 * sw + u0) * 4;
        const p10 = (v0 * sw + u1) * 4;
        const p01 = (v1 * sw + u0) * 4;
        const p11 = (v1 * sw + u1) * 4;
        const w00 = (1 - fu) * (1 - fv);
        const w10 = fu * (1 - fv);
        const w01 = (1 - fu) * fv;
        const w11 = fu * fv;
        const pa = data[p00 + 3] * w00 + data[p10 + 3] * w10 + data[p01 + 3] * w01 + data[p11 + 3] * w11;
        if (pa <= 0) {
          out[o + 3] = 0;
          continue;
        }
        const k = 1 / pa;
        // Premultiplied samples, back to straight colour for the canvas.
        out[o] = (data[p00] * w00 + data[p10] * w10 + data[p01] * w01 + data[p11] * w11) * 255 * k;
        out[o + 1] = (data[p00 + 1] * w00 + data[p10 + 1] * w10 + data[p01 + 1] * w01 + data[p11 + 1] * w11) * 255 * k;
        out[o + 2] = (data[p00 + 2] * w00 + data[p10 + 2] * w10 + data[p01 + 2] * w01 + data[p11 + 2] * w11) * 255 * k;
        out[o + 3] = pa / 255;
      }
    }, alpha);
  }
}

/** Taps of one axis: for each destination pixel, its first source pixel and their weights. */
interface Taps {
  first: Int32Array;
  count: Int32Array;
  weights: Float32Array;
  /** Where each destination pixel's weights start in `weights`. */
  at: Int32Array;
}

/**
 * Weights taking `src` pixels to destination pixels [from, to) of `dst`, mirrored when `flip`.
 * Shrinking averages whole source pixels, `src / dst` or one more, spread evenly; enlarging
 * interpolates (`smooth`) or repeats each source pixel `dst / src` or one more times.
 */
function taps(src: number, dst: number, from: number, to: number, flip: boolean, smooth: boolean): Taps {
  const n = to - from;
  const first = new Int32Array(n);
  const count = new Int32Array(n);
  const at = new Int32Array(n);
  const weights: number[] = [];
  const step = Math.floor(src / dst);
  const extra = src % dst;
  for (let i = 0; i < n; i++) {
    const di = flip ? dst - 1 - (from + i) : from + i;
    at[i] = weights.length;
    if (dst < src) {
      const a = di * step + Math.floor((di * extra) / dst);
      const b = (di + 1) * step + Math.floor(((di + 1) * extra) / dst);
      first[i] = a;
      count[i] = b - a;
      for (let s = a; s < b; s++) weights.push(1 / (b - a));
    } else if (smooth) {
      // Sampled at di · src / dst: the first destination pixel starts on the first source pixel.
      const c = (di * src) / dst;
      const lo = Math.min(src - 1, Math.floor(c));
      const t = c - lo;
      first[i] = lo;
      if (lo + 1 < src && t > 0) {
        count[i] = 2;
        weights.push(1 - t, t);
      } else {
        count[i] = 1;
        weights.push(1);
      }
    } else {
      // The source pixel whose run of repeats covers di.
      const reps = Math.floor(dst / src);
      const more = dst % src;
      let s = Math.min(src - 1, Math.floor(di / (reps + 1)));
      while (s > 0 && s * reps + Math.floor((s * more) / src) > di) s--;
      while (s + 1 < src && (s + 1) * reps + Math.floor(((s + 1) * more) / src) <= di) s++;
      first[i] = s;
      count[i] = 1;
      weights.push(1);
    }
  }
  return { first, count, weights: Float32Array.from(weights), at };
}

/**
 * `px` scaled to dw × dh, of which only pixels [rx0, rx1) × [ry0, ry1) are made, premultiplied.
 */
function resample(px: Pixmap, dw: number, dh: number, rx0: number, ry0: number, rx1: number, ry1: number, flipX: boolean, flipY: boolean): Premultiplied {
  const rw = rx1 - rx0;
  const rh = ry1 - ry0;
  // Interpolation only when both axes grow, by less than four times.
  const smooth = dw >= px.width && dh >= px.height && Math.floor(dw / px.width) < 4 && Math.floor(dh / px.height) < 4;
  const xt = taps(px.width, dw, rx0, rx1, flipX, smooth);
  const yt = taps(px.height, dh, ry0, ry1, flipY, smooth);
  if (px.opaque && (dw < px.width || dh < px.height)) return boxScale(px, xt, yt, rw, rh, dw < px.width, dh < px.height);
  const across = rowScaler(px, xt, rw);
  const out = new Uint8Array(rw * rh * 4);
  const stride = rw * 4;
  const yw = yt.weights;
  const acc = new Float64Array(stride);
  if (dh < px.height) {
    // Shrinking down, each source row adds to one destination row, so rows are scaled across one at a time.
    const line = new Float32Array(stride);
    for (let j = 0; j < rh; j++) {
      const n = yt.count[j];
      const wi0 = yt.at[j];
      for (let k = 0; k < n; k++) {
        across(yt.first[j] + k, line, 0);
        const w = yw[wi0 + k];
        if (k === 0) for (let i = 0; i < stride; i++) acc[i] = line[i] * w;
        else for (let i = 0; i < stride; i++) acc[i] += line[i] * w;
      }
      store(out, j * stride, acc, stride);
    }
    return { width: rw, height: rh, data: out };
  }
  let sy0 = px.height;
  let sy1 = 0;
  for (let i = 0; i < rh; i++) {
    sy0 = Math.min(sy0, yt.first[i]);
    sy1 = Math.max(sy1, yt.first[i] + yt.count[i]);
  }
  // Rows across first, as each feeds several destination rows.
  const rows = new Float32Array((sy1 - sy0) * stride);
  for (let sy = sy0; sy < sy1; sy++) across(sy, rows, (sy - sy0) * stride);
  for (let j = 0; j < rh; j++) {
    const n = yt.count[j];
    const base = (yt.first[j] - sy0) * stride;
    const wi0 = yt.at[j];
    for (let k = 0; k < n; k++) {
      const w = yw[wi0 + k];
      const p = base + k * stride;
      if (k === 0) for (let i = 0; i < stride; i++) acc[i] = rows[p + i] * w;
      else for (let i = 0; i < stride; i++) acc[i] += rows[p + i] * w;
    }
    store(out, j * stride, acc, stride);
  }
  return { width: rw, height: rh, data: out };
}

/**
 * Opaque pixels shrunk along at least one axis as Poppler's Splash does (scaleImageYdXd, YdXu
 * and YuXd): each box's byte sum times 2^23 / its size, rounded down to an integer, then shifted
 * back, so a box's average can come out one below its exact value. Growing axes repeat pixels.
 */
function boxScale(px: Pixmap, xt: Taps, yt: Taps, rw: number, rh: number, shrinkX: boolean, shrinkY: boolean): Premultiplied {
  const src = px.data;
  const sw = px.width;
  let cx0 = sw;
  let cx1 = 0;
  for (let i = 0; i < rw; i++) {
    cx0 = Math.min(cx0, xt.first[i]);
    cx1 = Math.max(cx1, xt.first[i] + xt.count[i]);
  }
  const cols = cx1 - cx0;
  // Each source column's sum over the destination row's source rows.
  const sums = new Int32Array(cols * 3);
  // Growing across, each source column's value, repeated into the destination.
  const values = shrinkX ? null : new Uint8Array(cols * 3);
  const out = new Uint8Array(rw * rh * 4);
  const { first, count } = xt;
  for (let j = 0, o = 0; j < rh; j++) {
    const ystep = shrinkY ? yt.count[j] : 1;
    const sy = yt.first[j];
    for (let k = 0; k < ystep; k++) {
      let p = ((sy + k) * sw + cx0) * 4;
      if (k === 0) {
        for (let c = 0; c < cols * 3; c += 3, p += 4) {
          sums[c] = src[p];
          sums[c + 1] = src[p + 1];
          sums[c + 2] = src[p + 2];
        }
      } else {
        for (let c = 0; c < cols * 3; c += 3, p += 4) {
          sums[c] += src[p];
          sums[c + 1] += src[p + 1];
          sums[c + 2] += src[p + 2];
        }
      }
    }
    if (values) {
      const d = Math.floor(0x800000 / ystep);
      for (let c = 0; c < cols * 3; c++) values[c] = (sums[c] * d) >> 23;
      for (let i = 0; i < rw; i++, o += 4) {
        const c = (first[i] - cx0) * 3;
        out[o] = values[c];
        out[o + 1] = values[c + 1];
        out[o + 2] = values[c + 2];
        out[o + 3] = 255;
      }
      continue;
    }
    for (let i = 0; i < rw; i++, o += 4) {
      const n = count[i];
      // Below 2^31: a box of n pixels sums to at most 255 · n, and d · n is at most 2^23.
      const d = Math.floor(0x800000 / (ystep * n));
      let c = (first[i] - cx0) * 3;
      let r = 0;
      let g = 0;
      let b = 0;
      for (const end = c + n * 3; c < end; c += 3) {
        r += sums[c];
        g += sums[c + 1];
        b += sums[c + 2];
      }
      out[o] = (r * d) >> 23;
      out[o + 1] = (g * d) >> 23;
      out[o + 2] = (b * d) >> 23;
      out[o + 3] = 255;
    }
  }
  return { width: rw, height: rh, data: out };
}

/** Destination bytes from summed premultiplied values, truncated as Poppler's integer arithmetic does. */
function store(out: Uint8Array, o: number, acc: Float64Array, n: number): void {
  for (let i = 0; i < n; i++) {
    const v = acc[i];
    out[o + i] = v < 0 ? 0 : v > 255 ? 255 : (v + 1e-3) | 0;
  }
}

/** Scales source row `sy` across into `dst` at `o`, premultiplied, through the taps `xt`. */
function rowScaler(px: Pixmap, xt: Taps, rw: number): (sy: number, dst: Float32Array, o: number) => void {
  const src = px.data;
  const sw = px.width;
  const { count, first, at, weights } = xt;
  const opaque = px.opaque;
  return (sy, dst, o) => {
    const so = sy * sw * 4;
    for (let i = 0; i < rw; i++, o += 4) {
      let r = 0;
      let g = 0;
      let b = 0;
      let al = 0;
      const n = count[i];
      let p = so + first[i] * 4;
      let wi = at[i];
      if (opaque) {
        for (let k = 0; k < n; k++, p += 4, wi++) {
          const w = weights[wi];
          r += src[p] * w;
          g += src[p + 1] * w;
          b += src[p + 2] * w;
          al += 255 * w;
        }
      } else {
        for (let k = 0; k < n; k++, p += 4, wi++) {
          const w = weights[wi];
          const aw = (src[p + 3] * w) / 255;
          r += src[p] * aw;
          g += src[p + 1] * aw;
          b += src[p + 2] * aw;
          al += src[p + 3] * w;
        }
      }
      dst[o] = r;
      dst[o + 1] = g;
      dst[o + 2] = b;
      dst[o + 3] = al;
    }
  };
}

/** Composite premultiplied pixels with their top-left at (x0, y0), through the canvas clip. */
function composite(canvas: Canvas, img: Premultiplied, x0: number, y0: number, alpha: number): void {
  const { data, width } = canvas;
  const c = canvas.clip;
  const mask = c.mask;
  const src = img.data;
  canvas.touch(x0, y0, x0 + img.width, y0 + img.height);
  for (let j = 0; j < img.height; j++) {
    const y = y0 + j;
    let o = (y * width + x0) * 3;
    let s = j * img.width * 4;
    const mo = mask ? (y - c.my) * c.mw - c.mx : 0;
    for (let i = 0; i < img.width; i++, o += 3, s += 4) {
      let k = alpha;
      if (x0 + i === c.x0) k *= c.e0;
      if (x0 + i === c.x1 - 1) k *= c.e1;
      if (mask) k = (k * mask[mo + x0 + i]) / 255;
      const sa = src[s + 3];
      if (!sa || k <= 0.0005) continue;
      if (sa === 255 && k >= 0.9995) {
        data[o] = src[s];
        data[o + 1] = src[s + 1];
        data[o + 2] = src[s + 2];
        continue;
      }
      const keep = 1 - (sa / 255) * k;
      data[o] = (src[s] * k + data[o] * keep + 0.5) | 0;
      data[o + 1] = (src[s + 1] * k + data[o + 1] * keep + 0.5) | 0;
      data[o + 2] = (src[s + 2] * k + data[o + 2] * keep + 0.5) | 0;
    }
  }
}
