// Strokes as outlines to fill by the nonzero rule: each subpath (after dashing) becomes its left
// side, an end cap, its right side back and a start cap, with joins on the outside of each turn.

import { Polylines, TOLERANCE } from './geometry.js';

export type Cap = 'butt' | 'round' | 'square';
export type Join = 'miter' | 'round' | 'bevel';

export interface StrokeStyle {
  /** Line width in device pixels. */
  width: number;
  cap: Cap;
  join: Join;
  miterLimit: number;
  /** Dash pattern in device pixels (on, off, ...), and how far into it the path starts. */
  dash?: readonly number[];
  phase?: number;
}

/** Beyond this many dashes, a pattern is too fine to draw dash by dash. */
export const MAX_DASHES = 100_000;

/** Roughly how many dashes `dash` cuts the subpaths of `p` into. */
export function dashCount(p: Polylines, dash: readonly number[]): number {
  const total = dash.reduce((a, b) => a + b, 0);
  if (!(total > 0)) return 0;
  let length = 0;
  const { xs, ys } = p;
  for (let s = 0; s < p.starts.length; s++) {
    const start = p.starts[s];
    const end = p.end(s);
    for (let i = start + 1; i < end; i++) length += Math.hypot(xs[i] - xs[i - 1], ys[i] - ys[i - 1]);
    if (p.closed[s] && end - start > 1) length += Math.hypot(xs[start] - xs[end - 1], ys[start] - ys[end - 1]);
  }
  return (length / total) * Math.max(1, dash.length / 2) + p.starts.length;
}

/** The share of a dash pattern that is drawn. */
export function dashCoverage(dash: readonly number[]): number {
  const pattern = dash.length % 2 ? [...dash, ...dash] : dash;
  let on = 0;
  let total = 0;
  for (let i = 0; i < pattern.length; i++) {
    total += pattern[i];
    if (i % 2 === 0) on += pattern[i];
  }
  return total > 0 ? on / total : 1;
}

const scratch = new Polylines();
const dashed = new Polylines();

/** Add the outline of `src`'s stroke to `out`. */
export function strokeOutline(src: Polylines, style: StrokeStyle, out: Polylines): void {
  const hw = style.width / 2;
  if (!(hw > 0) || !Number.isFinite(hw)) return;
  let path = src;
  const dash = style.dash;
  if (dash && dash.length && dash.some((d) => d > 0) && dash.every((d) => d >= 0 && Number.isFinite(d)) && dashCount(src, dash) <= MAX_DASHES) {
    dashPolylines(src, dash, style.phase ?? 0, dashed);
    path = dashed;
  }
  for (let s = 0; s < path.starts.length; s++) {
    const { xs, ys } = path;
    // Drop repeated points: they have no direction.
    scratch.reset();
    const start = path.starts[s];
    const end = path.end(s);
    scratch.moveTo(xs[start], ys[start]);
    for (let i = start + 1; i < end; i++) {
      const px = scratch.xs[scratch.n - 1];
      const py = scratch.ys[scratch.n - 1];
      if (Math.abs(xs[i] - px) > 1e-9 || Math.abs(ys[i] - py) > 1e-9) scratch.lineTo(xs[i], ys[i]);
    }
    let closed = path.closed[s];
    if (closed && scratch.n > 1 && Math.abs(scratch.xs[0] - scratch.xs[scratch.n - 1]) <= 1e-9 && Math.abs(scratch.ys[0] - scratch.ys[scratch.n - 1]) <= 1e-9) scratch.n--;
    if (scratch.n < 3) closed = closed && scratch.n > 2;
    if (scratch.n === 1) dot(scratch.xs[0], scratch.ys[0], hw, style.cap, out);
    else if (closed) strokeClosed(scratch, hw, style, out);
    else strokeOpen(scratch, hw, style, out);
  }
}


function dot(x: number, y: number, hw: number, cap: Cap, out: Polylines): void {
  if (cap === 'round') {
    const n = arcSteps(hw, Math.PI * 2);
    out.moveTo(x + hw, y);
    for (let i = 1; i < n; i++) out.lineTo(x + hw * Math.cos((i / n) * Math.PI * 2), y + hw * Math.sin((i / n) * Math.PI * 2));
    out.close();
  } else if (cap === 'square') {
    out.moveTo(x - hw, y - hw);
    out.lineTo(x + hw, y - hw);
    out.lineTo(x + hw, y + hw);
    out.lineTo(x - hw, y + hw);
    out.close();
  }
}

function arcSteps(r: number, angle: number): number {
  const step = r > TOLERANCE ? 2 * Math.acos(1 - TOLERANCE / r) : Math.PI / 2;
  return Math.max(2, Math.min(256, Math.ceil(Math.abs(angle) / step)));
}

/** Points of an arc around (cx, cy) from angle a0 by `sweep`, excluding the first point. */
function arc(out: Polylines, cx: number, cy: number, r: number, a0: number, sweep: number): void {
  const n = arcSteps(r, sweep);
  for (let i = 1; i <= n; i++) {
    const a = a0 + (sweep * i) / n;
    out.lineTo(cx + r * Math.cos(a), cy + r * Math.sin(a));
  }
}

/**
 * One side of the polyline from point `from` to `to` (step ±1), offset by `side` (±1) times the
 * normal, with joins at the interior vertices. The side's first point is assumed already added.
 */
function side(p: Polylines, from: number, to: number, step: number, sideSign: number, hw: number, style: StrokeStyle, out: Polylines, closed: boolean): void {
  const { xs, ys } = p;
  const n = p.n;
  const at = (i: number) => ((i % n) + n) % n;
  for (let i = from; i !== to; i += step) {
    const a = at(i);
    const b = at(i + step);
    const c = at(i + 2 * step);
    // Segment a→b's offset end, then the join at b towards segment b→c.
    let dx0 = xs[b] - xs[a];
    let dy0 = ys[b] - ys[a];
    const l0 = Math.hypot(dx0, dy0);
    dx0 /= l0;
    dy0 /= l0;
    const nx0 = -dy0 * hw * sideSign;
    const ny0 = dx0 * hw * sideSign;
    out.lineTo(xs[b] + nx0, ys[b] + ny0);
    const last = !closed && i + step === to;
    if (last) break;
    let dx1 = xs[c] - xs[b];
    let dy1 = ys[c] - ys[b];
    const l1 = Math.hypot(dx1, dy1);
    dx1 /= l1;
    dy1 /= l1;
    const nx1 = -dy1 * hw * sideSign;
    const ny1 = dx1 * hw * sideSign;
    join(out, xs[b], ys[b], dx0, dy0, nx0, ny0, dx1, dy1, nx1, ny1, sideSign, hw, style);
  }
}

function join(out: Polylines, px: number, py: number, dx0: number, dy0: number, nx0: number, ny0: number, dx1: number, dy1: number, nx1: number, ny1: number, sideSign: number, hw: number, style: StrokeStyle): void {
  const cross = (dx0 * dy1 - dy0 * dx1) * sideSign;
  const dot = dx0 * dx1 + dy0 * dy1;
  // Straight on: no join.
  if (Math.abs(cross) < 1e-9 && dot > 0) {
    out.lineTo(px + nx1, py + ny1);
    return;
  }
  if (cross > 0) {
    // This side is inside the turn: pivot through the vertex, which the other side's area covers.
    out.lineTo(px, py);
    out.lineTo(px + nx1, py + ny1);
    return;
  }
  if (style.join === 'round') {
    const a0 = Math.atan2(ny0, nx0);
    let sweep = Math.atan2(ny1, nx1) - a0;
    // The outside of the turn goes the short way round.
    while (sweep > Math.PI) sweep -= 2 * Math.PI;
    while (sweep < -Math.PI) sweep += 2 * Math.PI;
    arc(out, px, py, hw, a0, sweep);
    return;
  }
  if (style.join === 'miter') {
    const k = hw * hw + (nx0 * nx1 + ny0 * ny1);
    if (k > 1e-12) {
      const mx = ((nx0 + nx1) * hw * hw) / k;
      const my = ((ny0 + ny1) * hw * hw) / k;
      if (Math.hypot(mx, my) <= style.miterLimit * hw) {
        out.lineTo(px + mx, py + my);
        out.lineTo(px + nx1, py + ny1);
        return;
      }
    }
  }
  out.lineTo(px + nx1, py + ny1);
}

function cap(out: Polylines, px: number, py: number, dx: number, dy: number, hw: number, style: StrokeStyle): void {
  // From the left offset (px - dy·hw, py + dx·hw) round to the right one, past the end along (dx, dy).
  const nx = -dy * hw;
  const ny = dx * hw;
  if (style.cap === 'square') {
    out.lineTo(px + nx + dx * hw, py + ny + dy * hw);
    out.lineTo(px - nx + dx * hw, py - ny + dy * hw);
  } else if (style.cap === 'round') {
    arc(out, px, py, hw, Math.atan2(ny, nx), -Math.PI);
    return;
  }
  out.lineTo(px - nx, py - ny);
}

function strokeOpen(p: Polylines, hw: number, style: StrokeStyle, out: Polylines): void {
  const { xs, ys } = p;
  const n = p.n;
  // Left side forward.
  let dx = xs[1] - xs[0];
  let dy = ys[1] - ys[0];
  let l = Math.hypot(dx, dy);
  dx /= l;
  dy /= l;
  out.moveTo(xs[0] - dy * hw, ys[0] + dx * hw);
  side(p, 0, n - 1, 1, 1, hw, style, out, false);
  // End cap.
  dx = xs[n - 1] - xs[n - 2];
  dy = ys[n - 1] - ys[n - 2];
  l = Math.hypot(dx, dy);
  cap(out, xs[n - 1], ys[n - 1], dx / l, dy / l, hw, style);
  // Right side back (the left side of the reversed path).
  side(p, n - 1, 0, -1, 1, hw, style, out, false);
  dx = xs[0] - xs[1];
  dy = ys[0] - ys[1];
  l = Math.hypot(dx, dy);
  cap(out, xs[0], ys[0], dx / l, dy / l, hw, style);
  out.close();
}

function strokeClosed(p: Polylines, hw: number, style: StrokeStyle, out: Polylines): void {
  const { xs, ys } = p;
  const n = p.n;
  // Each side is a loop: start just after vertex 0 on segment 0→1, end with the join at vertex 0.
  for (const dir of [1, -1]) {
    const first = dir === 1 ? 0 : n;
    const next = first + dir;
    const a = ((first % n) + n) % n;
    const b = ((next % n) + n) % n;
    let dx = xs[b] - xs[a];
    let dy = ys[b] - ys[a];
    const l = Math.hypot(dx, dy);
    dx /= l;
    dy /= l;
    out.moveTo(xs[a] - dy * hw, ys[a] + dx * hw);
    side(p, first, first + n * dir, dir, 1, hw, style, out, true);
    out.close();
  }
}

/** Split subpaths into their dashes (open subpaths), in `out`. */
export function dashPolylines(src: Polylines, dash: readonly number[], phase: number, out: Polylines): void {
  out.reset();
  const pattern = dash.length % 2 ? [...dash, ...dash] : dash;
  const total = pattern.reduce((a, b) => a + b, 0);
  for (let s = 0; s < src.starts.length; s++) {
    const start = src.starts[s];
    const end = src.end(s);
    let index = 0;
    let left = pattern[0];
    let offset = ((phase % total) + total) % total;
    while (offset > 0) {
      if (offset >= left) {
        offset -= left;
        index = (index + 1) % pattern.length;
        left = pattern[index];
      } else {
        left -= offset;
        offset = 0;
      }
    }
    let on = index % 2 === 0;
    const { xs, ys } = src;
    const count = end - start + (src.closed[s] ? 1 : 0);
    if (on) out.moveTo(xs[start], ys[start]);
    for (let k = 1; k < count; k++) {
      const i0 = start + k - 1;
      const i1 = k === end - start ? start : start + k;
      let x0 = xs[i0];
      let y0 = ys[i0];
      const x1 = xs[i1];
      const y1 = ys[i1];
      let len = Math.hypot(x1 - x0, y1 - y0);
      while (len > 0) {
        if (left > len) {
          left -= len;
          if (on) out.lineTo(x1, y1);
          break;
        }
        const t = left / len;
        const mx = x0 + (x1 - x0) * t;
        const my = y0 + (y1 - y0) * t;
        if (on) out.lineTo(mx, my);
        else out.moveTo(mx, my);
        on = !on;
        len -= left;
        x0 = mx;
        y0 = my;
        index = (index + 1) % pattern.length;
        left = pattern[index];
        // A zero-length dash still draws its caps.
        while (left === 0) {
          if (on) out.lineTo(mx, my);
          else out.moveTo(mx, my);
          on = !on;
          index = (index + 1) % pattern.length;
          left = pattern[index];
        }
      }
    }
  }
}
