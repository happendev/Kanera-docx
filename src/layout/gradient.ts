import type { GradientFill } from '../docx/model.js';
import type { DrawOp, PathSeg } from './types.js';

const decode = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const encode = (l: number) => (l <= 0.0031308 ? 12.92 * l : 1.055 * l ** (1 / 2.4) - 0.055);
const rgb = (hex: string) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);

/**
 * The colour `t` (0..1) of the way along a gradient. A shape's gradient is a straight blend in sRGB
 * between stops; an eased one blends in linear light along 3u² − 2u³, as the sampled shadings in
 * Word for Mac's PDFs show.
 */
function colorAt(g: GradientFill, t: number): string {
  const st = g.stops;
  const i = st.findIndex((s) => s.pos >= t);
  if (i <= 0) return st[i === 0 ? 0 : st.length - 1].color;
  const u = (t - st[i - 1].pos) / (st[i].pos - st[i - 1].pos || 1);
  const [a, b] = [rgb(st[i - 1].color), rgb(st[i].color)];
  const mix = g.eased
    ? a.map((v, k) => encode(decode(v) + (decode(b[k]) - decode(v)) * u * u * (3 - 2 * u)))
    : a.map((v, k) => v + (b[k] - v) * u);
  return mix.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** An ellipse centred on (cx, cy) as four Bézier quarters. */
function ellipse(cx: number, cy: number, rx: number, ry: number): PathSeg[] {
  const kx = 0.5522847498 * rx;
  const ky = 0.5522847498 * ry;
  return [
    { op: 'M', pts: [cx + rx, cy] },
    { op: 'C', pts: [cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry] },
    { op: 'C', pts: [cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy] },
    { op: 'C', pts: [cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry] },
    { op: 'C', pts: [cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy] },
    { op: 'Z', pts: [] },
  ];
}

/**
 * A gradient fill of the shape `segs`, clipped to it: bands across it for a linear gradient, rings
 * out from the focus for a path one, each drawn over the last.
 */
export function gradientOps(segs: PathSeg[], g: GradientFill, alpha: number): DrawOp[] {
  const xs = segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 0));
  const ys = segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 1));
  if (!xs.length) return [];
  const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
  const [w, h] = [x1 - x0, y1 - y0];
  const band = (path: PathSeg[], t: number): DrawOp => ({ kind: 'path', segs: path, fill: colorAt(g, t), ...(alpha < 1 && { fillAlpha: alpha }), stroke: null });
  const ops: DrawOp[] = [{ kind: 'clip', x: x0, y: y0, w, h, segs }];
  if (g.path) {
    // Position 0 is the focus rectangle, 1 the shape's edge: rings drawn from the outside in.
    const f = g.focus ?? { l: 0.5, t: 0.5, r: 0.5, b: 0.5 };
    const [fx0, fy0, fx1, fy1] = [x0 + f.l * w, y0 + f.t * h, x1 - f.r * w, y1 - f.b * h];
    const [cx, cy] = [(fx0 + fx1) / 2, (fy0 + fy1) / 2];
    // A circle's last ring reaches the corner farthest from the focus.
    const reach = Math.max(...[[x0, y0], [x1, y0], [x0, y1], [x1, y1]].map(([px, py]) => Math.hypot((px - cx) / (w || 1), (py - cy) / (h || 1))));
    const n = Math.max(2, Math.min(96, Math.ceil(Math.max(w, h) / 2)));
    for (let k = n - 1; k >= 0; k--) {
      const s = (k + 1) / n;
      const ring: PathSeg[] =
        g.path === 'circle'
          ? ellipse(cx, cy, w * reach * s + 0.3, h * reach * s + 0.3)
          : [
              { op: 'M', pts: [fx0 - (fx0 - x0) * s - 0.3, fy0 - (fy0 - y0) * s - 0.3] },
              { op: 'L', pts: [fx1 + (x1 - fx1) * s + 0.3, fy0 - (fy0 - y0) * s - 0.3] },
              { op: 'L', pts: [fx1 + (x1 - fx1) * s + 0.3, fy1 + (y1 - fy1) * s + 0.3] },
              { op: 'L', pts: [fx0 - (fx0 - x0) * s - 0.3, fy1 + (y1 - fy1) * s + 0.3] },
              { op: 'Z', pts: [] },
            ];
      ops.push(band(ring, (k + 0.5) / n));
    }
  } else {
    // A scaled angle is the unit square's, stretched with the shape.
    const a = (g.angle * Math.PI) / 180;
    const [ux, uy] = g.scaled && w > 0 && h > 0 ? [Math.cos(a) / w, Math.sin(a) / h] : [Math.cos(a), Math.sin(a)];
    const len = Math.hypot(ux, uy) || 1;
    const [dx, dy] = [ux / len, uy / len];
    const corners = [[x0, y0], [x1, y0], [x0, y1], [x1, y1]].map(([x, y]) => x * dx + y * dy);
    const t0 = Math.min(...corners);
    const t1 = Math.max(...corners);
    const reach = Math.hypot(w, h);
    const [cx, cy] = [(x0 + x1) / 2, (y0 + y1) / 2];
    const c0 = cx * dx + cy * dy;
    const n = Math.max(2, Math.min(128, Math.ceil((t1 - t0) / 1.5)));
    for (let k = 0; k < n; k++) {
      // Each band runs on to the far end, under the bands after it: rasterisers anti-alias each fill
      // on its own, so bands that only met or overlapped a little let the page show through as streaks.
      const s0 = t0 + ((t1 - t0) * k) / n - c0;
      const s1 = t1 - c0 + 1;
      const p = (s: number, side: number): [number, number] => [cx + dx * s - dy * side * reach, cy + dy * s + dx * side * reach];
      ops.push(band([{ op: 'M', pts: p(s0, -1) }, { op: 'L', pts: p(s1, -1) }, { op: 'L', pts: p(s1, 1) }, { op: 'L', pts: p(s0, 1) }, { op: 'Z', pts: [] }], (k + 0.5) / n));
    }
  }
  ops.push({ kind: 'restore' });
  return ops;
}
