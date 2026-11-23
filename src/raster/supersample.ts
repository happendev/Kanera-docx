// Anti-aliased coverage of filled paths as Poppler's Splash rasteriser samples it, for vector
// paint: each pixel is a 4 × 4 grid of sub-pixels, and on each sub-row every sub-pixel an edge
// passes through counts as inside, together with the spans between edges that the winding rule
// fills. Coverage is the share of the 16 sub-pixels inside. Edges therefore come out a little
// heavier than their exact area, as they do in Poppler.

import type { CoverageRow } from './coverage.js';
import type { Polylines } from './geometry.js';

const AA = 4;

export class Supersampled {
  /** Window origin and size in device pixels. */
  x0 = 0;
  y0 = 0;
  w = 0;
  h = 0;
  /** Edge crossings: sub-row, first and last sub-pixel touched, and winding (0 for none). */
  private n = 0;
  private iy = new Int32Array(1024);
  private ix0 = new Int32Array(1024);
  private ix1 = new Int32Array(1024);
  private ic = new Int8Array(1024);
  private order = new Int32Array(0);
  private rowStart = new Int32Array(0);
  private counts = new Int32Array(0);
  private middle = new Int32Array(0);
  private row = new Float32Array(0);
  /** The pixel spans [start, end) of one row's sub-row spans, in pairs. */
  private spans = new Int32Array(64);

  /** Start a window [x0, x1) × [y0, y1) of whole device pixels; false when it is empty. */
  begin(x0: number, y0: number, x1: number, y1: number): boolean {
    const w = x1 - x0;
    const h = y1 - y0;
    if (w <= 0 || h <= 0) return false;
    this.x0 = x0;
    this.y0 = y0;
    this.w = w;
    this.h = h;
    this.n = 0;
    if (this.counts.length < w + 1) {
      this.counts = new Int32Array(w + 1);
      this.middle = new Int32Array(w + 1);
      this.row = new Float32Array(w + 1);
    }
    if (this.rowStart.length < h * AA + 1) this.rowStart = new Int32Array(h * AA + 1);
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
      if (xs[end - 1] !== xs[start] || ys[end - 1] !== ys[start]) this.line(xs[end - 1], ys[end - 1], xs[start], ys[start]);
    }
  }

  /** Add one edge, in device pixels. */
  line(ax: number, ay: number, bx: number, by: number): void {
    const X0 = (ax - this.x0) * AA;
    const Y0 = (ay - this.y0) * AA;
    const X1 = (bx - this.x0) * AA;
    const Y1 = (by - this.y0) * AA;
    if (!(Number.isFinite(X0) && Number.isFinite(Y0) && Number.isFinite(X1) && Number.isFinite(Y1))) return;
    const rows = this.h * AA;
    const yMin = Math.min(Y0, Y1);
    const yMax = Math.max(Y0, Y1);
    if (yMax < 0 || yMin >= rows) return;
    if (Y0 === Y1) {
      this.add(Math.floor(Y0), Math.floor(X0), Math.floor(X1), 0);
      return;
    }
    // Winding counts where the edge crosses a sub-row's top line.
    const count = Y0 > Y1 ? 1 : -1;
    const first = Math.max(0, Math.floor(yMin));
    const last = Math.min(rows - 1, Math.floor(yMax));
    if (X0 === X1) {
      const x = Math.floor(X0);
      for (let y = first; y <= last; y++) this.add(y, x, x, yMin <= y && y < yMax ? count : 0);
      return;
    }
    const dxdy = (X1 - X0) / (Y1 - Y0);
    const xMin = Math.min(X0, X1);
    const xMax = Math.max(X0, X1);
    const clamp = (x: number) => (x < xMin ? xMin : x > xMax ? xMax : x);
    let xa = Math.floor(clamp(X0 + (first - Y0) * dxdy));
    for (let y = first; y <= last; y++) {
      const xb = Math.floor(clamp(X0 + (y + 1 - Y0) * dxdy));
      this.add(y, xa, xb, yMin <= y && y < yMax ? count : 0);
      xa = xb;
    }
  }

  private add(y: number, x0: number, x1: number, count: number): void {
    if (y < 0 || y >= this.h * AA) return;
    if (this.n === this.iy.length) {
      const grow = <T extends Int32Array | Int8Array>(a: T): T => {
        const b = new (a.constructor as new (n: number) => T)(a.length * 2);
        b.set(a);
        return b;
      };
      this.iy = grow(this.iy);
      this.ix0 = grow(this.ix0);
      this.ix1 = grow(this.ix1);
      this.ic = grow(this.ic);
    }
    // Crossings beyond the window only matter for winding and for spans reaching into it.
    const limit = this.w * AA;
    if (x0 > x1) {
      const t = x0;
      x0 = x1;
      x1 = t;
    }
    const i = this.n++;
    this.iy[i] = y;
    this.ix0[i] = x0 < -1 ? -1 : x0 > limit ? limit : x0;
    this.ix1[i] = x1 < -1 ? -1 : x1 > limit ? limit : x1;
    this.ic[i] = count;
  }

  /**
   * Hand each touched row's coverage (0 to 1) to `emit`, as `row[from..to)` relative to the window,
   * once for each run of covered pixels, and clear the window for the next path.
   */
  sweep(evenOdd: boolean, emit: CoverageRow): void {
    const n = this.n;
    if (!n) return;
    const rows = this.h * AA;
    const { iy, ix0, ix1, ic } = this;
    // Bucket the crossings by sub-row.
    const start = this.rowStart;
    start.fill(0, 0, rows + 1);
    for (let i = 0; i < n; i++) start[iy[i] + 1]++;
    for (let r = 0; r < rows; r++) start[r + 1] += start[r];
    if (this.order.length < n) this.order = new Int32Array(Math.max(n, this.order.length * 2));
    const order = this.order;
    const fill = start.slice(0, rows);
    for (let i = 0; i < n; i++) order[fill[iy[i]]++] = i;
    const limit = this.w * AA;
    const { counts, middle, row } = this;
    for (let py = 0; py < this.h; py++) {
      const r0 = py * AA;
      if (start[r0] === start[r0 + AA]) continue;
      let nSpans = 0;
      for (let r = r0; r < r0 + AA; r++) {
        const a = start[r];
        const b = start[r + 1];
        if (a === b) continue;
        // Sort the sub-row's crossings by their first sub-pixel (few per row: insertion sort).
        for (let i = a + 1; i < b; i++) {
          const v = order[i];
          const key = ix0[v];
          let j = i - 1;
          while (j >= a && ix0[order[j]] > key) {
            order[j + 1] = order[j];
            j--;
          }
          order[j + 1] = v;
        }
        let k = a;
        let winding = 0;
        while (k < b) {
          const first = order[k++];
          let s0 = ix0[first];
          let s1 = ix1[first];
          winding += ic[first];
          while (k < b && (ix0[order[k]] <= s1 || (evenOdd ? (winding & 1) !== 0 : winding !== 0))) {
            const next = order[k++];
            if (ix1[next] > s1) s1 = ix1[next];
            winding += ic[next];
          }
          if (s0 < 0) s0 = 0;
          s1++;
          if (s1 > limit) s1 = limit;
          if (s0 >= s1) continue;
          // Sub-pixels [s0, s1) into whole pixels: partial ends, whole ones between.
          const pa = s0 >> 2;
          const pb = (s1 - 1) >> 2;
          if (pa === pb) counts[pa] += s1 - s0;
          else {
            counts[pa] += (pa + 1) * AA - s0;
            counts[pb] += s1 - pb * AA;
            if (pb > pa + 1) {
              middle[pa + 1] += AA;
              middle[pb] -= AA;
            }
          }
          if (nSpans * 2 + 2 > this.spans.length) {
            const grown = new Int32Array(this.spans.length * 2);
            grown.set(this.spans);
            this.spans = grown;
          }
          this.spans[nSpans * 2] = pa;
          this.spans[nSpans * 2 + 1] = pb + 1;
          nSpans++;
        }
      }
      if (!nSpans) continue;
      // The sub-rows' spans in order, joined where they meet: runs of covered pixels.
      const spans = this.spans;
      for (let i = 1; i < nSpans; i++) {
        const a = spans[i * 2];
        const b = spans[i * 2 + 1];
        let j = i - 1;
        while (j >= 0 && spans[j * 2] > a) {
          spans[j * 2 + 2] = spans[j * 2];
          spans[j * 2 + 3] = spans[j * 2 + 1];
          j--;
        }
        spans[j * 2 + 2] = a;
        spans[j * 2 + 3] = b;
      }
      let k = 0;
      while (k < nSpans) {
        const lo = spans[k * 2];
        let hi = spans[k * 2 + 1];
        for (k++; k < nSpans && spans[k * 2] <= hi; k++) if (spans[k * 2 + 1] > hi) hi = spans[k * 2 + 1];
        let run = 0;
        for (let x = lo; x < hi; x++) {
          run += middle[x];
          row[x] = (counts[x] + run) / (AA * AA);
          counts[x] = 0;
          middle[x] = 0;
        }
        middle[hi] = 0;
        emit(py + this.y0, lo, hi, row);
      }
    }
    this.n = 0;
  }
}
