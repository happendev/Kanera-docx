import type { GradientFill } from '../ooxml/model.js';
import type { DrawOp, PathSeg, ShadeGeometry } from './types.js';

const decode = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const encode = (l: number) => (l <= 0.0031308 ? 12.92 * l : 1.055 * l ** (1 / 2.4) - 0.055);
const rgb = (hex: string) => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);

/** Positions a gradient is sampled at: fine enough that the blend between samples can't be told from the curve. */
const SAMPLES = 256;

/** The stops either side of `t` (0..1), and how far `t` is from the first to the second. */
function between(g: GradientFill, t: number): [GradientFill['stops'][number], GradientFill['stops'][number], number] {
  const st = g.stops;
  const i = st.findIndex((s) => s.pos >= t);
  if (i <= 0) {
    const s = st[i === 0 ? 0 : st.length - 1];
    return [s, s, 0];
  }
  return [st[i - 1], st[i], (t - st[i - 1].pos) / (st[i].pos - st[i - 1].pos || 1)];
}

/**
 * The colour `t` (0..1) of the way along a gradient. A shape's gradient is a straight blend in sRGB
 * between stops; an eased one blends in linear light along 3u² − 2u³, as the sampled shadings in
 * Word for Mac's PDFs show.
 */
function colorAt(g: GradientFill, t: number): string {
  const [s0, s1, u] = between(g, t);
  const [a, b] = [rgb(s0.color), rgb(s1.color)];
  const mix = g.eased
    ? a.map((v, k) => encode(decode(v) + (decode(b[k]) - decode(v)) * u * u * (3 - 2 * u)))
    : a.map((v, k) => v + (b[k] - v) * u);
  return mix.map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** The opacity `t` of the way along a gradient, blended as its colour is. */
function alphaAt(g: GradientFill, t: number): number {
  const [s0, s1, u] = between(g, t);
  const [a, b] = [s0.alpha ?? 1, s1.alpha ?? 1];
  return a + (b - a) * (g.eased ? u * u * (3 - 2 * u) : u);
}

/**
 * A gradient fill of the shape `segs`, clipped to it, as one smooth shading: along a line for a
 * linear gradient, in rings out from the focus for a path one. `alpha` is its opacity, unless its
 * stops give their own.
 */
export function gradientOps(segs: PathSeg[], g: GradientFill, alpha: number): DrawOp[] {
  const xs = segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 0));
  const ys = segs.flatMap((s) => s.pts.filter((_, i) => i % 2 === 1));
  if (!xs.length || !g.stops.length) return [];
  const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
  const [w, h] = [x1 - x0, y1 - y0];
  if (w <= 0 || h <= 0) return [];
  let geometry: ShadeGeometry;
  if (g.path === 'circle') {
    // Position 0 is the focus, 1 the ellipse through the corner farthest from it.
    const f = g.focus ?? { l: 0.5, t: 0.5, r: 0.5, b: 0.5 };
    const [cx, cy] = [(x0 + f.l * w + x1 - f.r * w) / 2, (y0 + f.t * h + y1 - f.b * h) / 2];
    const reach = Math.max(...[[x0, y0], [x1, y0], [x0, y1], [x1, y1]].map(([px, py]) => Math.hypot((px - cx) / w, (py - cy) / h)));
    geometry = { type: 'radial', cx, cy, rx: w * reach, ry: h * reach };
  } else if (g.path) {
    // Position 0 is the focus rectangle, 1 the shape's edge.
    const f = g.focus ?? { l: 0.5, t: 0.5, r: 0.5, b: 0.5 };
    const [fl, fr, ft, fb] = [x0 + f.l * w, x1 - f.r * w, y0 + f.t * h, y1 - f.b * h];
    geometry = { type: 'rect', inner: [Math.min(fl, fr), Math.min(ft, fb), Math.max(fl, fr), Math.max(ft, fb)], outer: [x0, y0, x1, y1] };
  } else {
    // A scaled angle is the unit square's, stretched with the shape.
    const a = (g.angle * Math.PI) / 180;
    const [ux, uy] = g.scaled ? [Math.cos(a) / w, Math.sin(a) / h] : [Math.cos(a), Math.sin(a)];
    const len = Math.hypot(ux, uy) || 1;
    const [dx, dy] = [ux / len, uy / len];
    const along = [[x0, y0], [x1, y0], [x0, y1], [x1, y1]].map(([x, y]) => x * dx + y * dy);
    const [t0, t1] = [Math.min(...along), Math.max(...along)];
    const [cx, cy] = [(x0 + x1) / 2, (y0 + y1) / 2];
    const c0 = cx * dx + cy * dy;
    geometry = { type: 'axial', x0: cx + dx * (t0 - c0), y0: cy + dy * (t0 - c0), x1: cx + dx * (t1 - c0), y1: cy + dy * (t1 - c0) };
  }
  // A straight blend between two end stops is exact with its two colours.
  const st = g.stops;
  const n = !g.eased && st.length === 2 && st[0].pos === 0 && st[1].pos === 1 ? 2 : SAMPLES;
  const at = Array.from({ length: n }, (_, i) => i / (n - 1));
  const opacity = (s: GradientFill['stops'][number]) => s.alpha ?? 1;
  const varies = st.some((s) => Math.abs(opacity(s) - opacity(st[0])) > 1e-3);
  return [
    { kind: 'clip', x: x0, y: y0, w, h, segs },
    {
      kind: 'shade',
      geometry,
      bbox: [x0, y0, x1, y1],
      colors: at.map((t) => colorAt(g, t)),
      ...(varies ? { alphas: at.map((t) => alphaAt(g, t)) } : alpha < 1 && { alpha }),
    },
    { kind: 'restore' },
  ];
}
