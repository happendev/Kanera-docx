// Gradients (shade ops) as the PDF writer's shadings paint them: along an axis or out from a
// centre, both extended past their ends, or across rectangular rings of triangles, through the
// current clip. Colours run linearly between evenly spaced samples.

import type { ShadeOp } from '../layout/types.js';
import type { Canvas } from './canvas.js';
import { invert, type Matrix } from './geometry.js';

export function drawShade(canvas: Canvas, op: ShadeOp, ctm: Matrix): void {
  const inv = invert(ctm);
  if (!inv || !op.colors.length) return;
  const n = op.colors.length;
  const rgb = new Float32Array(n * 3);
  op.colors.forEach((hex, i) => {
    const v = parseInt(hex.slice(0, 6), 16) || 0;
    rgb[i * 3] = (v >> 16) & 255;
    rgb[i * 3 + 1] = (v >> 8) & 255;
    rgb[i * 3 + 2] = v & 255;
  });
  // Each sample's change to the next, exact as the samples are whole numbers.
  const step = new Float32Array(n * 3);
  for (let i = 0; i < (n - 1) * 3; i++) step[i] = rgb[i + 3] - rgb[i];
  // Opacity samples are 8-bit in the PDF's soft mask too.
  const alphas = op.alphas ? Float32Array.from(op.alphas, (a) => Math.round(Math.max(0, Math.min(1, a)) * 255) / 255) : null;
  const g = op.geometry;
  let region: [number, number, number, number] | null = null;
  const deviceBox = (l: number, t: number, r: number, b: number, pad: number): [number, number, number, number] => {
    const xs: number[] = [];
    const ys: number[] = [];
    for (const [x, y] of [[l, t], [r, t], [l, b], [r, b]]) {
      xs.push(ctm[0] * x + ctm[2] * y + ctm[4]);
      ys.push(ctm[1] * x + ctm[3] * y + ctm[5]);
    }
    return [Math.floor(Math.min(...xs) - pad), Math.floor(Math.min(...ys) - pad), Math.ceil(Math.max(...xs) + pad), Math.ceil(Math.max(...ys) + pad)];
  };
  // A soft mask's form ends a point past the box; a mesh paints only its triangles.
  if (alphas) {
    const [l, t, r, b] = op.bbox;
    region = deviceBox(l - 1, t - 1, r + 1, b + 1, 0);
  }
  if (g.type === 'rect') {
    const [l, t, r, b] = g.outer;
    const box = deviceBox(l, t, r, b, 1);
    region = region ? [Math.max(region[0], box[0]), Math.max(region[1], box[1]), Math.min(region[2], box[2]), Math.min(region[3], box[3])] : box;
  }

  // Where along the gradient (0 to 1) user-space point (x, y) lies: axial and radial positions
  // are worked out inline below, rectangular rings through their triangles.
  const ring = g.type === 'rect' ? ringPosition(g.inner, g.outer) : null;
  let ax = 0, ay = 0, adx = 0, ady = 0, cx = 0, cy = 0, rx = 1, ry = 1;
  if (g.type === 'axial') {
    const dx = g.x1 - g.x0;
    const dy = g.y1 - g.y0;
    const len2 = dx * dx + dy * dy || 1;
    ax = g.x0;
    ay = g.y0;
    adx = dx / len2;
    ady = dy / len2;
  } else if (g.type === 'radial') {
    cx = g.cx;
    cy = g.cy;
    rx = g.rx || 1e-9;
    ry = g.ry || 1e-9;
  }
  const axial = g.type === 'axial';

  const alpha = op.alpha !== undefined && op.alpha < 1 ? Math.max(0, op.alpha) : 1;
  // An axial position that a step along the row (or down a column) leaves alone, as each user-space
  // coordinate either stays put or counts for nothing: one colour per row, or the same row each time.
  const flatRows = axial && (adx === 0 || inv[0] === 0) && (ady === 0 || inv[1] === 0);
  const sameRows = axial && (adx === 0 || inv[2] === 0) && (ady === 0 || inv[3] === 0);
  let last: Float32Array | null = null;
  let lastX0 = 0;
  let lastX1 = 0;
  // Upright, x in user space follows the column and y the row, so the position's x term (u² when
  // radial) is worked out once per column and its y term once per row, in the same arithmetic.
  const separable = !ring && inv[1] === 0 && inv[2] === 0;
  let cols = new Float64Array(0);
  let colX0 = 0;
  let colX1 = -1;
  canvas.shade(null, false, region, (y, x0, x1, out) => {
    if (sameRows && last && x0 === lastX0 && x1 === lastX1) {
      out.set(last.subarray(0, (x1 - x0) * 4));
      return;
    }
    // Pixel centres in user space, stepping along the row.
    const yc = y + 0.5;
    let px = inv[0] * (x0 + 0.5) + inv[2] * yc + inv[4];
    let py = inv[1] * (x0 + 0.5) + inv[3] * yc + inv[5];
    const sx = inv[0];
    const sy = inv[1];
    const end = flatRows ? Math.min(x1, x0 + 1) : x1;
    let rowTerm = 0;
    if (separable) {
      if (x0 !== colX0 || x1 !== colX1) {
        if (cols.length < x1 - x0) cols = new Float64Array(x1 - x0);
        let cx0 = px;
        for (let k = 0; k < end - x0; k++, cx0 += sx) {
          if (axial) cols[k] = (cx0 - ax) * adx;
          else {
            const u = (cx0 - cx) / rx;
            cols[k] = u * u;
          }
        }
        colX0 = x0;
        colX1 = x1;
      }
      if (axial) rowTerm = (py - ay) * ady;
      else {
        const v = (py - cy) / ry;
        rowTerm = v * v;
      }
    }
    for (let x = x0, o = 0; x < end; x++, o += 4, px += sx, py += sy) {
      let t: number;
      if (separable) t = axial ? cols[x - x0] + rowTerm : Math.sqrt(cols[x - x0] + rowTerm);
      else if (axial) t = (px - ax) * adx + (py - ay) * ady;
      else if (ring) t = ring(px, py);
      else {
        const u = (px - cx) / rx;
        const v = (py - cy) / ry;
        t = Math.sqrt(u * u + v * v);
      }
      if (t !== t) {
        out[o + 3] = 0;
        continue;
      }
      if (n === 1) {
        out[o] = rgb[0];
        out[o + 1] = rgb[1];
        out[o + 2] = rgb[2];
        out[o + 3] = alphas ? alphas[0] : 1;
        continue;
      }
      t = t < 0 ? 0 : t > 1 ? 1 : t;
      const s = t * (n - 1);
      let i = s | 0;
      if (i > n - 2) i = n - 2;
      const f = s - i;
      const p = i * 3;
      out[o] = rgb[p] + step[p] * f;
      out[o + 1] = rgb[p + 1] + step[p + 1] * f;
      out[o + 2] = rgb[p + 2] + step[p + 2] * f;
      out[o + 3] = alphas ? alphas[i] + (alphas[i + 1] - alphas[i]) * f : 1;
    }
    if (flatRows) {
      const [r, g, b, a] = out;
      for (let o = 4; o < (x1 - x0) * 4; o += 4) {
        out[o] = r;
        out[o + 1] = g;
        out[o + 2] = b;
        out[o + 3] = a;
      }
    }
    if (sameRows) {
      const n = (x1 - x0) * 4;
      if (!last || last.length < n) last = new Float32Array(n);
      last.set(out.subarray(0, n));
      lastX0 = x0;
      lastX1 = x1;
    }
  }, alpha);
}

type Box = [number, number, number, number];

/**
 * Position across rectangular rings, as the PDF writer's triangle mesh interpolates it: 0 on and
 * inside the inner rectangle, 1 on the outer one, linear within each triangle; NaN outside.
 */
function ringPosition(inner: Box, outer: Box): (x: number, y: number) => number {
  const [ol, ot, or, ob] = outer;
  const [il, it, ir, ib] = inner;
  const O = [[ol, ot], [or, ot], [or, ob], [ol, ob]];
  const I = [[il, it], [ir, it], [ir, ib], [il, ib]];
  // Triangles of (x, y, t) vertices, with each one's plane t = a·x + b·y + c and edge tests.
  const tris: Array<{ v: number[][]; a: number; b: number; c: number }> = [];
  const add = (v: number[][]) => {
    const [[x0, y0, t0], [x1, y1, t1], [x2, y2, t2]] = v;
    const det = (x1 - x0) * (y2 - y0) - (x2 - x0) * (y1 - y0);
    if (Math.abs(det) < 1e-12) return;
    const a = ((t1 - t0) * (y2 - y0) - (t2 - t0) * (y1 - y0)) / det;
    const b = ((x1 - x0) * (t2 - t0) - (x2 - x0) * (t1 - t0)) / det;
    tris.push({ v, a, b, c: t0 - a * x0 - b * y0 });
  };
  for (let k = 0; k < 4; k++) {
    const [o0, o1, i0, i1] = [O[k], O[(k + 1) % 4], I[k], I[(k + 1) % 4]];
    add([[...o0, 1], [...o1, 1], [...i1, 0]]);
    add([[...o0, 1], [...i1, 0], [...i0, 0]]);
  }
  const inside = (v: number[][], x: number, y: number) => {
    let sign = 0;
    for (let k = 0; k < 3; k++) {
      const [ax, ay] = v[k];
      const [bx, by] = v[(k + 1) % 3];
      const cross = (bx - ax) * (y - ay) - (by - ay) * (x - ax);
      if (Math.abs(cross) < 1e-9) continue;
      const s = cross > 0 ? 1 : -1;
      if (sign && s !== sign) return false;
      sign = s;
    }
    return true;
  };
  const [xl, xr] = [Math.min(il, ir), Math.max(il, ir)];
  const [yt, yb] = [Math.min(it, ib), Math.max(it, ib)];
  return (x, y) => {
    if (x >= xl && x <= xr && y >= yt && y <= yb) return 0;
    for (const t of tris) if (inside(t.v, x, y)) return t.a * x + t.b * y + t.c;
    return NaN;
  };
}
