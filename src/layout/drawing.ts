import type { Drawing, LineStyle, RunProps } from '../docx/model.js';
import { DEFAULT_RUN } from '../docx/props.js';
import { arc } from '../util/arc.js';
import { layoutBlocks, type BoxEnv } from './box.js';
import { renderChart } from './chart.js';
import { gradientOps } from './gradient.js';
import { metafileOps } from './metafile.js';
import { shapeText } from './shaper.js';
import { appendOps, translateOps, type DrawOp, type GlyphOp, type PathSeg } from './types.js';
import { uprightEastAsian } from './vertical.js';

const K = 0.5522847498;

/** Transform rotating/flipping about the centre of a box (top-left coordinates, y down). */
export function boxTransform(x: number, y: number, w: number, h: number, rotation: number, flipH: boolean, flipV: boolean): [number, number, number, number, number, number] | null {
  if (!rotation && !flipH && !flipV) return null;
  const t = (rotation * Math.PI) / 180;
  const cos = Math.cos(t);
  const sin = Math.sin(t);
  const sx = flipH ? -1 : 1;
  const sy = flipV ? -1 : 1;
  const a = cos * sx;
  const b = sin * sx;
  const c = -sin * sy;
  const d = cos * sy;
  const cx = x + w / 2;
  const cy = y + h / 2;
  return [a, b, c, d, cx - (a * cx + c * cy), cy - (b * cx + d * cy)];
}

function ellipse(x: number, y: number, w: number, h: number): PathSeg[] {
  const rx = w / 2;
  const ry = h / 2;
  const cx = x + rx;
  const cy = y + ry;
  return [
    { op: 'M', pts: [cx + rx, cy] },
    { op: 'C', pts: [cx + rx, cy + ry * K, cx + rx * K, cy + ry, cx, cy + ry] },
    { op: 'C', pts: [cx - rx * K, cy + ry, cx - rx, cy + ry * K, cx - rx, cy] },
    { op: 'C', pts: [cx - rx, cy - ry * K, cx - rx * K, cy - ry, cx, cy - ry] },
    { op: 'C', pts: [cx + rx * K, cy - ry, cx + rx, cy - ry * K, cx + rx, cy] },
    { op: 'Z', pts: [] },
  ];
}

function poly(points: number[]): PathSeg[] {
  const segs: PathSeg[] = [];
  for (let i = 0; i < points.length; i += 2) segs.push({ op: i === 0 ? 'M' : 'L', pts: [points[i], points[i + 1]] });
  segs.push({ op: 'Z', pts: [] });
  return segs;
}

/** Rounded corners: radii [tl, tr, br, bl]. */
function roundedRect(x: number, y: number, w: number, h: number, r: [number, number, number, number]): PathSeg[] {
  const [tl, tr, br, bl] = r.map((v) => Math.max(0, Math.min(v, w / 2, h / 2)));
  return [
    { op: 'M', pts: [x + tl, y] },
    { op: 'L', pts: [x + w - tr, y] },
    { op: 'C', pts: [x + w - tr + tr * K, y, x + w, y + tr - tr * K, x + w, y + tr] },
    { op: 'L', pts: [x + w, y + h - br] },
    { op: 'C', pts: [x + w, y + h - br + br * K, x + w - br + br * K, y + h, x + w - br, y + h] },
    { op: 'L', pts: [x + bl, y + h] },
    { op: 'C', pts: [x + bl - bl * K, y + h, x, y + h - bl + bl * K, x, y + h - bl] },
    { op: 'L', pts: [x, y + tl] },
    { op: 'C', pts: [x, y + tl - tl * K, x + tl - tl * K, y, x + tl, y] },
    { op: 'Z', pts: [] },
  ];
}

/**
 * A (rounded) rectangle with a wedge out to (dx, dy) from its centre. As in DrawingML's presets,
 * the wedge leaves the side the tip is furthest past, its base a quarter of that side long and
 * on the tip's half.
 */
function wedgeCallout(x: number, y: number, w: number, h: number, dx: number, dy: number, r: number): PathSeg[] {
  r = Math.max(0, Math.min(r, w / 2, h / 2));
  const tip = [x + w / 2 + dx, y + h / 2 + dy];
  const across = Math.abs((dx * h) / w) > Math.abs((dy * w) / h);
  const [xa, xb] = (dx > 0 ? [7, 10] : [2, 5]).map((v) => x + (w * v) / 12);
  const [ya, yb] = (dy > 0 ? [7, 10] : [2, 5]).map((v) => y + (h * v) / 12);
  const k = r * (1 - K);
  const segs: PathSeg[] = [{ op: 'M', pts: [x + r, y] }];
  const to = (px: number, py: number) => segs.push({ op: 'L', pts: [px, py] });
  const wedge = (ax: number, ay: number, bx: number, by: number) => {
    to(ax, ay);
    to(tip[0], tip[1]);
    to(bx, by);
  };
  const corner = (c1x: number, c1y: number, c2x: number, c2y: number, ex: number, ey: number) => {
    if (r > 0) segs.push({ op: 'C', pts: [c1x, c1y, c2x, c2y, ex, ey] });
  };
  if (!across && dy < 0) wedge(xa, y, xb, y);
  to(x + w - r, y);
  corner(x + w - k, y, x + w, y + k, x + w, y + r);
  if (across && dx > 0) wedge(x + w, ya, x + w, yb);
  to(x + w, y + h - r);
  corner(x + w, y + h - k, x + w - k, y + h, x + w - r, y + h);
  if (!across && dy > 0) wedge(xb, y + h, xa, y + h);
  to(x + r, y + h);
  corner(x + k, y + h, x, y + h - k, x, y + h - r);
  if (across && dx <= 0) wedge(x, yb, x, ya);
  to(x, y + r);
  corner(x, y + k, x + k, y, x + r, y);
  segs.push({ op: 'Z', pts: [] });
  return segs;
}

const LINE_GEOMETRIES = /^(line|straightConnector1|bentConnector\d|curvedConnector\d|lineInv)$/;

/**
 * Outline of a preset or custom geometry inside the box (x, y, w, h). A custom geometry gives the
 * subpaths its fill draws, or with `part` 'stroke' the ones its outline draws.
 */
export function geometryPath(d: Drawing, x: number, y: number, w: number, h: number, part: 'fill' | 'stroke' = 'fill'): PathSeg[] {
  const adj = (name: string, dflt: number) => (d.adjust?.[name] ?? dflt) / 100000;
  const ss = Math.min(w, h);
  const g = d.geometry;
  if (g === 'custom' && d.path) {
    const segs: PathSeg[] = [];
    for (const sub of d.path) if (sub[part]) for (const s of sub.segs) segs.push({ op: s.op, pts: s.pts.map((v, i) => (i % 2 === 0 ? x + v * w : y + v * h)) });
    return segs;
  }
  if (LINE_GEOMETRIES.test(g)) return [{ op: 'M', pts: [x, y] }, { op: 'L', pts: [x + w, y + h] }];
  switch (g) {
    case 'ellipse':
    case 'flowChartConnector':
      return ellipse(x, y, w, h);
    case 'roundRect':
    case 'flowChartAlternateProcess': {
      const r = ss * adj('adj', 16667);
      return roundedRect(x, y, w, h, [r, r, r, r]);
    }
    case 'flowChartTerminator':
      return roundedRect(x, y, w, h, [h / 2, h / 2, h / 2, h / 2].map((v) => Math.min(v, w / 2)) as [number, number, number, number]);
    case 'round1Rect': {
      const r = ss * adj('adj', 16667);
      return roundedRect(x, y, w, h, [0, r, 0, 0]);
    }
    case 'round2SameRect': {
      const r1 = ss * adj('adj1', 16667);
      const r2 = ss * adj('adj2', 0);
      return roundedRect(x, y, w, h, [r1, r1, r2, r2]);
    }
    case 'round2DiagRect': {
      const r1 = ss * adj('adj1', 16667);
      const r2 = ss * adj('adj2', 0);
      return roundedRect(x, y, w, h, [r1, r2, r1, r2]);
    }
    case 'snip1Rect': {
      const s = ss * adj('adj', 16667);
      return poly([x, y, x + w - s, y, x + w, y + s, x + w, y + h, x, y + h]);
    }
    case 'snip2SameRect': {
      const s = ss * adj('adj1', 16667);
      return poly([x + s, y, x + w - s, y, x + w, y + s, x + w, y + h, x, y + h, x, y + s]);
    }
    case 'triangle':
    case 'flowChartExtract': {
      const a = g === 'triangle' ? adj('adj', 50000) : 0.5;
      return poly([x + w * a, y, x + w, y + h, x, y + h]);
    }
    case 'flowChartMerge':
      return poly([x, y, x + w, y, x + w / 2, y + h]);
    case 'rtTriangle':
      return poly([x, y, x + w, y + h, x, y + h]);
    case 'diamond':
    case 'flowChartDecision':
      return poly([x + w / 2, y, x + w, y + h / 2, x + w / 2, y + h, x, y + h / 2]);
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const o = g === 'parallelogram' ? ss * adj('adj', 25000) : w / 5;
      return poly([x + o, y, x + w, y, x + w - o, y + h, x, y + h]);
    }
    case 'trapezoid': {
      const o = ss * adj('adj', 25000);
      return poly([x + o, y, x + w - o, y, x + w, y + h, x, y + h]);
    }
    case 'pentagon':
      return poly([x + w / 2, y, x + w, y + h * 0.382, x + w * 0.809, y + h, x + w * 0.191, y + h, x, y + h * 0.382]);
    case 'hexagon': {
      const o = ss * adj('adj', 25000);
      return poly([x + o, y, x + w - o, y, x + w, y + h / 2, x + w - o, y + h, x + o, y + h, x, y + h / 2]);
    }
    case 'octagon': {
      const o = ss * adj('adj', 29289);
      return poly([x + o, y, x + w - o, y, x + w, y + o, x + w, y + h - o, x + w - o, y + h, x + o, y + h, x, y + h - o, x, y + o]);
    }
    case 'homePlate': {
      const o = ss * adj('adj', 50000);
      return poly([x, y, x + w - o, y, x + w, y + h / 2, x + w - o, y + h, x, y + h]);
    }
    case 'chevron': {
      const o = ss * adj('adj', 50000);
      return poly([x, y, x + w - o, y, x + w, y + h / 2, x + w - o, y + h, x, y + h, x + o, y + h / 2]);
    }
    case 'rightArrow':
    case 'leftArrow': {
      const shaft = h * adj('adj1', 50000);
      const head = ss * adj('adj2', 50000);
      const t = y + (h - shaft) / 2;
      const b = t + shaft;
      if (g === 'rightArrow') return poly([x, t, x + w - head, t, x + w - head, y, x + w, y + h / 2, x + w - head, y + h, x + w - head, b, x, b]);
      return poly([x + w, t, x + head, t, x + head, y, x, y + h / 2, x + head, y + h, x + head, b, x + w, b]);
    }
    case 'leftRightArrow': {
      const head = Math.min(ss * adj('adj2', 50000), w / 2);
      const t = y + (h - h * adj('adj1', 50000)) / 2;
      const b = y + h - (t - y);
      return poly([x, y + h / 2, x + head, y, x + head, t, x + w - head, t, x + w - head, y, x + w, y + h / 2, x + w - head, y + h, x + w - head, b, x + head, b, x + head, y + h]);
    }
    case 'upDownArrow': {
      const head = Math.min(ss * adj('adj2', 50000), h / 2);
      const l = x + (w - w * adj('adj1', 50000)) / 2;
      const r = x + w - (l - x);
      return poly([x, y + head, x + w / 2, y, x + w, y + head, r, y + head, r, y + h - head, x + w, y + h - head, x + w / 2, y + h, x, y + h - head, l, y + h - head, l, y + head]);
    }
    case 'upArrow':
    case 'downArrow': {
      const shaft = w * adj('adj1', 50000);
      const head = ss * adj('adj2', 50000);
      const l = x + (w - shaft) / 2;
      const r = l + shaft;
      if (g === 'downArrow') return poly([l, y, r, y, r, y + h - head, x + w, y + h - head, x + w / 2, y + h, x, y + h - head, l, y + h - head]);
      return poly([l, y + h, r, y + h, r, y + head, x + w, y + head, x + w / 2, y, x, y + head, l, y + head]);
    }
    case 'plus':
    case 'flowChartSummingJunction': {
      const o = ss * adj('adj', 25000);
      return poly([x + o, y, x + w - o, y, x + w - o, y + o, x + w, y + o, x + w, y + h - o, x + w - o, y + h - o, x + w - o, y + h, x + o, y + h, x + o, y + h - o, x, y + h - o, x, y + o, x + o, y + o]);
    }
    case 'star5': {
      const pts: number[] = [];
      for (let i = 0; i < 10; i++) {
        const ang = -Math.PI / 2 + (i * Math.PI) / 5;
        const r = i % 2 === 0 ? 0.5 : 0.19;
        pts.push(x + w / 2 + Math.cos(ang) * r * w, y + h / 2 + 0.05 * h + Math.sin(ang) * r * h);
      }
      return poly(pts);
    }
    case 'flowChartDocument': {
      return [
        { op: 'M', pts: [x, y] },
        { op: 'L', pts: [x + w, y] },
        { op: 'L', pts: [x + w, y + h * 0.83] },
        { op: 'C', pts: [x + w * 0.75, y + h * 0.7, x + w * 0.5, y + h * 1.05, x, y + h * 0.95] },
        { op: 'Z', pts: [] },
      ];
    }
    case 'can':
    case 'flowChartMagneticDisk': {
      const e = Math.min(h * 0.25, ss * adj('adj', 25000)) / 2;
      return [
        { op: 'M', pts: [x, y + e] },
        { op: 'C', pts: [x, y + e - e * K * 2, x + w, y + e - e * K * 2, x + w, y + e] },
        { op: 'L', pts: [x + w, y + h - e] },
        { op: 'C', pts: [x + w, y + h - e + e * K * 2, x, y + h - e + e * K * 2, x, y + h - e] },
        { op: 'Z', pts: [] },
        { op: 'M', pts: [x, y + e] },
        { op: 'C', pts: [x, y + e + e * K * 2, x + w, y + e + e * K * 2, x + w, y + e] },
      ];
    }
    case 'frame': {
      const t = ss * adj('adj1', 12500);
      return [...poly([x, y, x + w, y, x + w, y + h, x, y + h]), ...poly([x + t, y + t, x + t, y + h - t, x + w - t, y + h - t, x + w - t, y + t])];
    }
    case 'noSmoking': {
      // Cut two circular segments out of the ellipse. The strip between their chords
      // forms the diagonal slash and shares the shape's fill, rotation and opacity.
      const thickness = Math.max(0, Math.min(ss / 2, ss * adj('adj', 18750)));
      const rx = w / 2 - thickness, ry = h / 2 - thickness;
      const segs = ellipse(x, y, w, h);
      if (rx <= 0 || ry <= 0) return segs;
      const angle = Math.PI / 4;
      const gap = Math.asin(Math.min(1, thickness * Math.hypot(1 / rx, 1 / ry) / (2 * Math.SQRT2)));
      const cx = x + w / 2, cy = y + h / 2;
      for (const base of [angle, angle + Math.PI]) {
        let a = base + gap;
        segs.push({ op: 'M', pts: [cx + rx * Math.cos(a), cy + ry * Math.sin(a)] });
        const step = (Math.PI - 2 * gap) / 2;
        const k = 4 / 3 * Math.tan(step / 4);
        for (let i = 0; i < 2; i++) {
          const b = a + step;
          segs.push({ op: 'C', pts: [
            cx + rx * (Math.cos(a) - k * Math.sin(a)), cy + ry * (Math.sin(a) + k * Math.cos(a)),
            cx + rx * (Math.cos(b) + k * Math.sin(b)), cy + ry * (Math.sin(b) - k * Math.cos(b)),
            cx + rx * Math.cos(b), cy + ry * Math.sin(b),
          ] });
          a = b;
        }
        segs.push({ op: 'Z', pts: [] });
      }
      return segs;
    }
    case 'blockArc': {
      // A band of ring from adj1 to adj2 (60000ths of a degree, clockwise), adj3 thick.
      const st = (d.adjust?.adj1 ?? 10800000) / 60000;
      const end = (d.adjust?.adj2 ?? 0) / 60000;
      const t = ss * Math.min(0.5, Math.max(0, adj('adj3', 25000)));
      let sweep = end - st;
      if (sweep <= 0) sweep += 360;
      const outer = arc(x + w / 2, y + h / 2, w / 2, h / 2, st, sweep);
      const inner = arc(x + w / 2, y + h / 2, w / 2 - t, h / 2 - t, end, -sweep);
      return [{ op: 'M', pts: outer.from }, ...outer.segs, { op: 'L', pts: inner.from }, ...inner.segs, { op: 'Z', pts: [] }];
    }
    case 'donut': {
      const t = ss * adj('adj', 25000);
      return [...ellipse(x, y, w, h), ...ellipse(x + t, y + t, w - 2 * t, h - 2 * t)];
    }
    case 'wedgeEllipseCallout': {
      // The tail leaves the ellipse 11° either side of the direction to its tip.
      const dx = w * adj('adj1', -20833);
      const dy = h * adj('adj2', 62500);
      const cx = x + w / 2, cy = y + h / 2;
      const pang = Math.atan2(dy * w, dx * h);
      const edge = (t: number) => [(w / 2) * Math.cos(t), (h / 2) * Math.sin(t)];
      const [x1, y1] = edge(pang + (11 * Math.PI) / 180);
      const [x2, y2] = edge(pang - (11 * Math.PI) / 180);
      const st = (Math.atan2(y1, x1) * 180) / Math.PI;
      let sweep = (Math.atan2(y2, x2) * 180) / Math.PI - st;
      if (sweep <= 0) sweep += 360;
      const a = arc(cx, cy, w / 2, h / 2, st, sweep);
      return [{ op: 'M', pts: [cx + dx, cy + dy] }, { op: 'L', pts: a.from }, ...a.segs, { op: 'Z', pts: [] }];
    }
    case 'wedgeRectCallout':
    case 'wedgeRoundRectCallout':
      return wedgeCallout(x, y, w, h, w * adj('adj1', -20833), h * adj('adj2', 62500), g === 'wedgeRoundRectCallout' ? ss * adj('adj3', 16667) : 0);
    default:
      return poly([x, y, x + w, y, x + w, y + h, x, y + h]);
  }
}

function lineDash(line: LineStyle): number[] | undefined {
  const w = Math.max(line.width, 0.5);
  switch (line.dash) {
    case 'dot':
    case 'sysDot':
    case 'shortdot':
      return [w, w];
    case 'dash':
    case 'sysDash':
    case 'shortdash':
      return [w * 4, w * 3];
    case 'lgDash':
    case 'longdash':
      return [w * 8, w * 3];
    case 'dashDot':
    case 'sysDashDot':
      return [w * 4, w * 3, w, w * 3];
    case 'lgDashDot':
      return [w * 8, w * 3, w, w * 3];
    case 'lgDashDotDot':
    case 'sysDashDotDot':
      return [w * 8, w * 3, w, w * 3, w, w * 3];
    default:
      return undefined;
  }
}

/**
 * Draw a drawing whose extent's top-left corner is at (x, y). Effect extents (shadows, glow) sit
 * outside this box and are not drawn.
 */
export function renderDrawing(d: Drawing, x: number, y: number, env: BoxEnv): DrawOp[] {
  if (d.hidden) return [];
  const w = d.width;
  const h = d.height;
  const ops: DrawOp[] = [];
  const isLine = LINE_GEOMETRIES.test(d.geometry);
  const m = boxTransform(x, y, w, h, d.rotation, d.flipH, d.flipV);
  if (m) ops.push({ kind: 'transform', matrix: m });

  if (d.kind === 'group' && d.children) {
    for (const kid of d.children) appendOps(ops, renderDrawing(kid, x + (kid.x ?? 0), y + (kid.y ?? 0), env));
  } else if (d.chart) {
    appendOps(ops, renderChart(d.chart, x, y, w, h, env));
  } else {
    const shape = d.kind === 'image' && d.geometry === 'rect' ? null : geometryPath(d, x, y, w, h);
    // A custom shape whose subpaths are all unfilled has nothing to fill or clip to.
    const path = shape?.length ? shape : null;
    const fill = isLine ? null : d.fill;
    if (fill?.gradient && path) ops.push(...gradientOps(path, fill.gradient, fill.alpha));
    else if (fill?.color && path) ops.push({ kind: 'path', segs: path, fill: fill.color, fillAlpha: fill.alpha < 1 ? fill.alpha : undefined, stroke: null });
    // Metafiles play back as vector drawing; anything they cannot express stays a picture.
    const meta = d.image ? metafileOps(d.image, x, y, w, h, d.crop, env.fonts) : null;
    if (meta) {
      if (path && d.geometry !== 'rect') ops.push({ kind: 'clip', x, y, w, h, segs: path }, ...meta, { kind: 'restore' });
      else appendOps(ops, meta);
    } else if (d.image) {
      if (path && d.geometry !== 'rect') {
        // Pictures cropped to a shape (rounded corners, circles).
        ops.push({ kind: 'clip', x, y, w, h, segs: path });
        ops.push({ kind: 'image', x, y, w, h, image: d.image, crop: d.crop, rotation: 0, flipH: false, flipV: false });
        ops.push({ kind: 'restore' });
      } else ops.push({ kind: 'image', x, y, w, h, image: d.image, crop: d.crop, rotation: 0, flipH: false, flipV: false });
    }
    const upright = d.textBox?.upright && m;
    if (d.textBox && !upright) appendOps(ops, renderTextBox(d, x, y, w, h, env));
    if (d.textPath) appendOps(ops, renderTextPath(d, x, y, w, h, env));
    const line = d.line;
    if (line?.color && line.width > 0) {
      const segs = d.geometry === 'custom' ? geometryPath(d, x, y, w, h, 'stroke') : path ?? geometryPath(d, x, y, w, h);
      if (segs.length) ops.push({ kind: 'path', segs, fill: null, stroke: line.color, strokeWidth: line.width, strokeAlpha: line.alpha < 1 ? line.alpha : undefined, dash: lineDash(line) });
    }
    if (upright) {
      ops.push({ kind: 'restore' }, ...renderTextBox(d, x, y, w, h, env));
      return ops;
    }
  }
  if (m) ops.push({ kind: 'restore' });
  return ops;
}

/**
 * The height of a text box set to resize to fit its text (a:spAutoFit, VML mso-fit-shape-to-text):
 * Word refits it to its text and insets as it lays the page out, shrinking it as well as growing it,
 * whatever size the file stored. Undefined for a box that keeps its size: one whose text area
 * depends on its height (rounded, elliptical, legacy text rectangles), or with vertical text.
 */
export function fittedHeight(d: Drawing, env: BoxEnv): number | undefined {
  const tb = d.textBox;
  if (!tb?.autoFit || tb.rect || tb.vertical === 'vert' || tb.vertical === 'vert270' || tb.vertical === 'eaVert') return undefined;
  if (!textRectInsets(d, d.width, d.height).every((v) => v === 0)) return undefined;
  return layoutBlocks(tb.blocks, Math.max(1, d.width - tb.insetL - tb.insetR), env).height + tb.insetT + tb.insetB;
}

/**
 * How far a preset's text rectangle sits inside its box on each side, before the text insets: text
 * in a rounded rectangle keeps clear of the corners, text in an ellipse stays in its inscribed box.
 */
function textRectInsets(d: Drawing, w: number, h: number): [number, number, number, number] {
  if (d.geometry === 'roundRect' || d.geometry === 'wedgeRoundRectCallout') {
    const a = d.geometry === 'roundRect' ? d.adjust?.adj : d.adjust?.adj3;
    const i = Math.min(w, h) * Math.min(0.5, Math.max(0, (a ?? 16667) / 100000)) * 0.29289;
    return [i, i, i, i];
  }
  if (d.geometry === 'ellipse' || d.geometry === 'wedgeEllipseCallout') {
    const ix = w * (0.5 - Math.SQRT1_2 / 2);
    const iy = h * (0.5 - Math.SQRT1_2 / 2);
    return [ix, iy, ix, iy];
  }
  return [0, 0, 0, 0];
}

function renderTextBox(d: Drawing, x: number, y: number, w: number, h: number, env: BoxEnv): DrawOp[] {
  const vertical = d.textBox!.vertical === 'vert' || d.textBox!.vertical === 'vert270' || d.textBox!.vertical === 'eaVert';
  const rect = d.textBox!.rect;
  const [rl, rt, rr, rb] = vertical ? [0, 0, 0, 0] : rect ?? textRectInsets(d, w, h);
  const tb = { ...d.textBox!, insetL: d.textBox!.insetL + rl, insetT: d.textBox!.insetT + rt, insetR: d.textBox!.insetR + rr, insetB: d.textBox!.insetB + rb };
  // Vertical text boxes lay text out along the box's height, then rotate it into place.
  const innerW = Math.max(1, (vertical ? h : w) - tb.insetL - tb.insetR);
  const innerH = (vertical ? w : h) - tb.insetT - tb.insetB;
  const box = layoutBlocks(tb.blocks, innerW, env);
  let dy = 0;
  if (tb.vAlign === 'center') dy = (innerH - box.height) / 2;
  else if (tb.vAlign === 'bottom') dy = innerH - box.height;
  const content: DrawOp[] = [];
  for (const it of box.items) appendOps(content, it.ops);
  const ops: DrawOp[] = [];
  // A legacy shape's text area also bounds its text; Word cuts off what overflows it.
  if (!tb.unclipped) ops.push(rect && !vertical ? { kind: 'clip', x: x + rl, y: y + rt, w: w - rl - rr, h: h - rt - rb } : { kind: 'clip', x, y, w, h });
  if (vertical) {
    const ox = x + (w - h) / 2;
    const oy = y + (h - w) / 2;
    const m = boxTransform(ox, oy, h, w, tb.vertical === 'vert270' ? 270 : 90, false, false)!;
    ops.push({ kind: 'transform', matrix: m });
    const turned = resolveBoxFloats(translateOps(content, ox + tb.insetL, oy + tb.insetT + dy), env);
    // East Asian vertical text keeps its ideographs and kana upright.
    appendOps(ops, tb.vertical === 'eaVert' ? uprightEastAsian(turned) : turned);
    ops.push({ kind: 'restore' });
  } else appendOps(ops, resolveBoxFloats(translateOps(content, x + tb.insetL, y + tb.insetT + dy), env));
  if (!tb.unclipped) ops.push({ kind: 'restore' });
  return ops;
}

/** Floats inside text boxes are positioned relative to the anchoring paragraph. */
function resolveBoxFloats(ops: DrawOp[], env: BoxEnv): DrawOp[] {
  const out: DrawOp[] = [];
  for (const op of ops) {
    if (op.kind !== 'float') {
      out.push(op);
      continue;
    }
    const d = op.drawing;
    const a = d.anchor;
    const fx = op.x + (a?.posH.offset ?? 0);
    const fy = op.y + (a?.posV.offset ?? 0);
    appendOps(out, renderDrawing(d, fx, fy, env));
  }
  return out;
}

function renderTextPath(d: Drawing, x: number, y: number, w: number, h: number, env: BoxEnv): DrawOp[] {
  const tp = d.textPath!;
  const props: RunProps = { ...DEFAULT_RUN, fontAscii: tp.font, fontHAnsi: tp.font, fontEastAsia: tp.font, fontCs: tp.font, bold: tp.bold, italic: tp.italic, size: 100, sizeCs: 100 };
  const runs = shapeText(tp.text, props, env.fonts);
  if (!runs.length) return [];
  let width = 0;
  let asc = 0;
  let desc = 0;
  for (const r of runs) {
    for (const g of r.glyphs) width += g.advance;
    asc = Math.max(asc, r.face.face.font.capHeight > 0 ? (r.face.face.font.capHeight / r.face.face.font.unitsPerEm) * 100 : r.ascent * 0.7);
    desc = Math.max(desc, 0);
  }
  if (width <= 0 || asc <= 0) return [];
  // The text is stretched to fill the shape (VML fitshape), measured on cap height.
  const scale = h / (asc + desc);
  const scaleX = w / (width * scale);
  const ops: DrawOp[] = [];
  let px = x;
  const baseline = y + asc * scale;
  for (const r of runs) {
    const op: GlyphOp = {
      kind: 'glyphs',
      x: px,
      y: baseline,
      face: r.face.face,
      size: r.size * scale,
      gids: [],
      xs: [],
      texts: [],
      color: tp.color ?? 'C0C0C0',
      syntheticBold: r.face.syntheticBold,
      syntheticItalic: r.face.syntheticItalic,
      scaleX,
      alpha: tp.alpha < 1 ? tp.alpha : undefined,
    };
    let gx = 0;
    for (const g of r.glyphs) {
      if (g.gid >= 0) {
        op.gids.push(g.gid);
        op.xs.push(gx);
        op.texts.push(g.text);
      }
      gx += g.advance * scale * scaleX;
    }
    px += gx;
    ops.push(op);
  }
  return ops;
}
