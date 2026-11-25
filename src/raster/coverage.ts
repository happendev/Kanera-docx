// Anti-aliased coverage of filled paths: each edge adds its signed area to the cells of the rows it
// crosses, and a running sum along each row gives every pixel's exact covered area (nonzero) or its
// even-odd fold. Only cells an edge touched are visited, plus the span inside the shape.

import type { Polylines } from './geometry.js';

export type CoverageRow = (y: number, from: number, to: number, row: Float32Array) => void;

export class Coverage {
  /** Window origin and size in device pixels; rows and spans below are relative to the origin. */
  x0 = 0;
  y0 = 0;
  w = 0;
  h = 0;
  private stride = 2;
  private gapStart = Infinity;
  private gapWidth = 0;
  private cells = new Float32Array(0);
  private minX = new Int32Array(0);
  private maxX = new Int32Array(0);
  private top = 0;
  private bottom = 0;
  readonly row = new Float32Array(0);
  private rowBuf = new Float32Array(0);

  /** Start a window [x0, x1) × [y0, y1) of whole device pixels; false when it is empty. */
  begin(x0: number, y0: number, x1: number, y1: number, gapStart = Infinity, gapWidth = 0): boolean {
    const w = x1 - x0;
    const h = y1 - y0;
    if (w <= 0 || h <= 0) return false;
    this.x0 = x0;
    this.y0 = y0;
    this.w = w;
    this.h = h;
    // A shadow may omit an interior band crossed only by horizontal edges.
    // Keep geometry in the original coordinates and pack the cell storage only.
    this.gapStart = gapStart;
    this.gapWidth = gapWidth;
    this.stride = w - gapWidth + 2;
    const size = this.stride * h;
    if (this.cells.length < size) this.cells = new Float32Array(Math.max(size, this.cells.length * 2));
    if (this.minX.length < h) {
      const n = Math.max(h, this.minX.length * 2);
      this.minX = new Int32Array(n).fill(0x7fffffff);
      this.maxX = new Int32Array(n).fill(-1);
    }
    if (this.rowBuf.length < this.stride) this.rowBuf = new Float32Array(Math.max(this.stride, this.rowBuf.length * 2));
    this.top = h;
    this.bottom = 0;
    return true;
  }

  /** Add the edges of `p`'s subpaths, each closed back to its start. */
  addPolylines(p: Polylines): void {
    const { xs, ys } = p;
    for (let s = 0; s < p.starts.length; s++) {
      const start = p.starts[s];
      const end = p.end(s);
      if (end - start < 2) continue;
      for (let i = start + 1; i < end; i++) this.line(xs[i - 1], ys[i - 1], xs[i], ys[i]);
      this.line(xs[end - 1], ys[end - 1], xs[start], ys[start]);
    }
  }

  /** Add one edge, in device pixels. */
  line(ax: number, ay: number, bx: number, by: number): void {
    ax -= this.x0;
    bx -= this.x0;
    ay -= this.y0;
    by -= this.y0;
    if (ay === by || !(Number.isFinite(ax) && Number.isFinite(ay) && Number.isFinite(bx) && Number.isFinite(by))) return;
    const { w, h } = this;
    if ((ay <= 0 && by <= 0) || (ay >= h && by >= h)) return;
    if (ay < 0 || ay > h || by < 0 || by > h) {
      const dxdy = (bx - ax) / (by - ay);
      if (ay < 0) {
        ax -= ay * dxdy;
        ay = 0;
      } else if (ay > h) {
        ax += (h - ay) * dxdy;
        ay = h;
      }
      if (by < 0) {
        bx -= by * dxdy;
        by = 0;
      } else if (by > h) {
        bx += (h - by) * dxdy;
        by = h;
      }
      if (ay === by) return;
    }
    // Edges right of the window can't change coverage inside it; edges left of it still wind.
    if (ax >= w && bx >= w) return;
    if (ax <= 0 && bx <= 0) return this.edge(0, ay, 0, by);
    if (ax < 0 || bx < 0) {
      const ym = ay + ((0 - ax) * (by - ay)) / (bx - ax);
      if (ax < 0) {
        this.edge(0, ay, 0, ym);
        ax = 0;
        ay = ym;
      } else {
        this.edge(0, ym, 0, by);
        bx = 0;
        by = ym;
      }
    }
    if (ax > w || bx > w) {
      const ym = ay + ((w - ax) * (by - ay)) / (bx - ax);
      if (ax > w) {
        ax = w;
        ay = ym;
      } else {
        bx = w;
        by = ym;
      }
    }
    this.edge(ax, ay, bx, by);
  }

  /** Signed area of an edge inside the window (0 ≤ x ≤ w, 0 ≤ y ≤ h), cell by cell. */
  private edge(x0: number, y0: number, x1: number, y1: number): void {
    if (y0 === y1) return;
    let dir = 1;
    if (y0 > y1) {
      dir = -1;
      let t = y0;
      y0 = y1;
      y1 = t;
      t = x0;
      x0 = x1;
      x1 = t;
    }
    const { cells, stride, minX, maxX } = this;
    const shift = Math.min(x0, x1) >= this.gapStart ? this.gapWidth : 0;
    const dxdy = (x1 - x0) / (y1 - y0);
    let x = x0;
    const first = Math.floor(y0);
    const last = Math.min(this.h, Math.ceil(y1));
    if (first < this.top) this.top = first;
    if (last > this.bottom) this.bottom = last;
    for (let y = first; y < last; y++) {
      const dy = Math.min(y + 1, y1) - Math.max(y, y0);
      const xnext = x + dxdy * dy;
      const d = dy * dir;
      let xa = x;
      let xb = xnext;
      if (xa > xb) {
        xa = xnext;
        xb = x;
      }
      if (xa < 0) xa = 0;
      const xaf = Math.floor(xa);
      const xai = (xaf | 0) - shift;
      const xbc = Math.ceil(xb);
      const xbi = (xbc | 0) - shift;
      const line = y * stride;
      if (xbi <= xai + 1) {
        const xmf = 0.5 * (x + xnext) - xaf;
        cells[line + xai] += d - d * xmf;
        cells[line + xai + 1] += d * xmf;
        if (xai < minX[y]) minX[y] = xai;
        if (xai + 1 > maxX[y]) maxX[y] = xai + 1;
      } else {
        const s = 1 / (xb - xa);
        const x0f = xa - xaf;
        const a0 = 0.5 * s * (1 - x0f) * (1 - x0f);
        const x1f = xb - xbc + 1;
        const am = 0.5 * s * x1f * x1f;
        cells[line + xai] += d * a0;
        if (xbi === xai + 2) cells[line + xai + 1] += d * (1 - a0 - am);
        else {
          const a1 = s * (1.5 - x0f);
          cells[line + xai + 1] += d * (a1 - a0);
          const ds = d * s;
          for (let xi = xai + 2; xi < xbi - 1; xi++) cells[line + xi] += ds;
          const a2 = a1 + (xbi - xai - 3) * s;
          cells[line + xbi - 1] += d * (1 - a2 - am);
        }
        cells[line + xbi] += d * am;
        if (xai < minX[y]) minX[y] = xai;
        if (xbi > maxX[y]) maxX[y] = xbi;
      }
      x = xnext;
    }
  }

  /**
   * Hand each touched row's coverage (0 to 1) to `emit`, as `row[from..to)` relative to the window,
   * and clear the window for the next path.
   */
  sweep(evenOdd: boolean, emit: CoverageRow): void {
    const { cells, stride, minX, maxX } = this;
    const w = stride - 2;
    const row = this.rowBuf;
    for (let y = this.top; y < this.bottom; y++) {
      const lo = minX[y];
      const hi = maxX[y];
      if (lo > hi) continue;
      minX[y] = 0x7fffffff;
      maxX[y] = -1;
      const line = y * stride;
      const end = Math.min(hi, w - 1);
      let acc = 0;
      let v = 0;
      for (let x = lo; x <= end; x++) {
        acc += cells[line + x];
        cells[line + x] = 0;
        v = acc < 0 ? -acc : acc;
        if (evenOdd) {
          v -= 2 * Math.floor(v * 0.5);
          if (v > 1) v = 2 - v;
        } else if (v > 1) v = 1;
        row[x] = v;
      }
      for (let x = Math.max(lo, end + 1); x <= hi; x++) {
        acc += cells[line + x];
        cells[line + x] = 0;
      }
      if (lo >= w) continue;
      let to = end + 1;
      // Inside a shape that runs past the row's last edge (on beyond the window), coverage holds.
      if (v > 1e-4 && to < w) {
        row.fill(v, to, w);
        to = w;
      }
      emit(y + this.y0, lo, to, row);
    }
    this.top = this.h;
    this.bottom = 0;
  }
}
