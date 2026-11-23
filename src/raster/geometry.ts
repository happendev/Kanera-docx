// Affine matrices and path flattening for the rasteriser. Matrices are PDF-style [a b c d e f]:
// x' = a·x + c·y + e, y' = b·x + d·y + f.

import { CLOSE, CUBIC, LINE, MOVE, QUAD } from '../fonts/outline.js';
import type { PathSeg } from '../layout/types.js';

export type Matrix = [number, number, number, number, number, number];

/** The matrix that applies `a`, then `b`. */
export function concat(a: Matrix, b: Matrix): Matrix {
  return [
    a[0] * b[0] + a[1] * b[2],
    a[0] * b[1] + a[1] * b[3],
    a[2] * b[0] + a[3] * b[2],
    a[2] * b[1] + a[3] * b[3],
    a[4] * b[0] + a[5] * b[2] + b[4],
    a[4] * b[1] + a[5] * b[3] + b[5],
  ];
}

export function invert(m: Matrix): Matrix | null {
  const det = m[0] * m[3] - m[1] * m[2];
  if (!det || !Number.isFinite(det)) return null;
  const a = m[3] / det;
  const b = -m[1] / det;
  const c = -m[2] / det;
  const d = m[0] / det;
  return [a, b, c, d, -(m[4] * a + m[5] * c), -(m[4] * b + m[5] * d)];
}

/** Whether `m` keeps axis-aligned rectangles axis-aligned (no rotation or skew). */
export function isAxisAligned(m: Matrix): boolean {
  return Math.abs(m[1]) < 1e-9 && Math.abs(m[2]) < 1e-9;
}

/** How much `m` scales lengths, on average over directions (exact for uniform scales). */
export function meanScale(m: Matrix): number {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2]));
}

/** Points of flattened subpaths in device pixels. */
export class Polylines {
  xs = new Float64Array(256);
  ys = new Float64Array(256);
  n = 0;
  /** Index of each subpath's first point; a subpath runs to the next one's start. */
  starts: number[] = [];
  closed: boolean[] = [];

  reset(): void {
    this.n = 0;
    this.starts.length = 0;
    this.closed.length = 0;
  }

  moveTo(x: number, y: number): void {
    this.starts.push(this.n);
    this.closed.push(false);
    this.push(x, y);
  }

  lineTo(x: number, y: number): void {
    if (!this.starts.length) return this.moveTo(x, y);
    this.push(x, y);
  }

  close(): void {
    if (this.starts.length) this.closed[this.closed.length - 1] = true;
  }

  end(i: number): number {
    return i + 1 < this.starts.length ? this.starts[i + 1] : this.n;
  }

  private push(x: number, y: number): void {
    if (this.n === this.xs.length) {
      const xs = new Float64Array(this.n * 2);
      const ys = new Float64Array(this.n * 2);
      xs.set(this.xs);
      ys.set(this.ys);
      this.xs = xs;
      this.ys = ys;
    }
    this.xs[this.n] = x;
    this.ys[this.n] = y;
    this.n++;
  }

  /** Bounds [x0, y0, x1, y1] of all points, or null when there are none. */
  bounds(): [number, number, number, number] | null {
    if (!this.n) return null;
    let x0 = Infinity;
    let y0 = Infinity;
    let x1 = -Infinity;
    let y1 = -Infinity;
    const { xs, ys } = this;
    for (let i = 0; i < this.n; i++) {
      const x = xs[i];
      const y = ys[i];
      if (x < x0) x0 = x;
      if (x > x1) x1 = x;
      if (y < y0) y0 = y;
      if (y > y1) y1 = y;
    }
    return Number.isFinite(x0 + y0 + x1 + y1) ? [x0, y0, x1, y1] : null;
  }
}

/** Curve flatness in device pixels: how far a flattened curve may stray from the true one. */
export const TOLERANCE = 0.1;
const MAX_SEGMENTS = 512;

function quadTo(out: Polylines, x0: number, y0: number, x1: number, y1: number, x2: number, y2: number): void {
  const dx = x0 - 2 * x1 + x2;
  const dy = y0 - 2 * y1 + y2;
  const n = Math.min(MAX_SEGMENTS, Math.ceil(Math.sqrt((0.25 * Math.sqrt(dx * dx + dy * dy)) / TOLERANCE)) || 1);
  for (let i = 1; i < n; i++) {
    const t = i / n;
    const u = 1 - t;
    out.lineTo(u * u * x0 + 2 * u * t * x1 + t * t * x2, u * u * y0 + 2 * u * t * y1 + t * t * y2);
  }
  out.lineTo(x2, y2);
}

function cubicTo(out: Polylines, x0: number, y0: number, x1: number, y1: number, x2: number, y2: number, x3: number, y3: number): void {
  const ax = x0 - 2 * x1 + x2;
  const ay = y0 - 2 * y1 + y2;
  const bx = x1 - 2 * x2 + x3;
  const by = y1 - 2 * y2 + y3;
  const dd = Math.sqrt(Math.max(ax * ax + ay * ay, bx * bx + by * by));
  const n = Math.min(MAX_SEGMENTS, Math.ceil(Math.sqrt((0.75 * dd) / TOLERANCE)) || 1);
  for (let i = 1; i < n; i++) {
    const t = i / n;
    const u = 1 - t;
    const a = u * u * u;
    const b = 3 * u * u * t;
    const c = 3 * u * t * t;
    const d = t * t * t;
    out.lineTo(a * x0 + b * x1 + c * x2 + d * x3, a * y0 + b * y1 + c * y2 + d * y3);
  }
  out.lineTo(x3, y3);
}

/** Flatten an outline (commands and points as in fonts/outline.ts) through `m` into `out`. */
export function flattenOutline(cmds: ArrayLike<number>, pts: ArrayLike<number>, m: Matrix, out: Polylines): void {
  const [a, b, c, d, e, f] = m;
  let p = 0;
  let cx = 0;
  let cy = 0;
  let sx = 0;
  let sy = 0;
  let closed = true;
  const X = (i: number) => a * pts[i] + c * pts[i + 1] + e;
  const Y = (i: number) => b * pts[i] + d * pts[i + 1] + f;
  for (let i = 0; i < cmds.length; i++) {
    const cmd = cmds[i];
    if (cmd !== MOVE && cmd !== CLOSE && closed) {
      out.moveTo(sx, sy);
      closed = false;
    }
    switch (cmd) {
      case MOVE:
        cx = sx = X(p);
        cy = sy = Y(p);
        p += 2;
        out.moveTo(cx, cy);
        closed = false;
        break;
      case LINE:
        cx = X(p);
        cy = Y(p);
        p += 2;
        out.lineTo(cx, cy);
        break;
      case QUAD: {
        const x1 = X(p);
        const y1 = Y(p);
        const x2 = X(p + 2);
        const y2 = Y(p + 2);
        p += 4;
        quadTo(out, cx, cy, x1, y1, x2, y2);
        cx = x2;
        cy = y2;
        break;
      }
      case CUBIC: {
        const x1 = X(p);
        const y1 = Y(p);
        const x2 = X(p + 2);
        const y2 = Y(p + 2);
        const x3 = X(p + 4);
        const y3 = Y(p + 4);
        p += 6;
        cubicTo(out, cx, cy, x1, y1, x2, y2, x3, y3);
        cx = x3;
        cy = y3;
        break;
      }
      case CLOSE:
        if (!closed) out.close();
        closed = true;
        cx = sx;
        cy = sy;
        break;
    }
  }
}

/** Flatten laid-out path segments (page coordinates) through `m` into `out`. */
export function flattenSegs(segs: readonly PathSeg[], m: Matrix, out: Polylines): void {
  const [a, b, c, d, e, f] = m;
  let cx = 0;
  let cy = 0;
  let sx = 0;
  let sy = 0;
  let open = false;
  for (const seg of segs) {
    const p = seg.pts;
    if (seg.op === 'M') {
      cx = sx = a * p[0] + c * p[1] + e;
      cy = sy = b * p[0] + d * p[1] + f;
      out.moveTo(cx, cy);
      open = true;
      continue;
    }
    if (seg.op === 'Z') {
      if (open) out.close();
      open = false;
      cx = sx;
      cy = sy;
      continue;
    }
    if (!open) {
      out.moveTo(cx, cy);
      sx = cx;
      sy = cy;
      open = true;
    }
    if (seg.op === 'L') {
      cx = a * p[0] + c * p[1] + e;
      cy = b * p[0] + d * p[1] + f;
      out.lineTo(cx, cy);
    } else {
      const x1 = a * p[0] + c * p[1] + e;
      const y1 = b * p[0] + d * p[1] + f;
      const x2 = a * p[2] + c * p[3] + e;
      const y2 = b * p[2] + d * p[3] + f;
      const x3 = a * p[4] + c * p[5] + e;
      const y3 = b * p[4] + d * p[5] + f;
      cubicTo(out, cx, cy, x1, y1, x2, y2, x3, y3);
      cx = x3;
      cy = y3;
    }
  }
}
