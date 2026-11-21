import type { Chart, ChartAxis, ChartErrorBars, ChartLabels, ChartLine, ChartPlot, ChartSeries, ChartShape, ChartTextStyle, ChartTitle, ChartTrendline } from '../docx/model.js';
import { DEFAULT_RUN } from '../docx/props.js';
import type { BoxEnv } from './box.js';
import { gradientOps } from './gradient.js';
import { shapeText, type ShapedRun } from './shaper.js';
import type { DrawOp, GlyphOp, PathSeg } from './types.js';

/** Space Word leaves between the chart's edge and its title, legend and plot. */
// Word's automatic chart layout, measured on its PDFs: labels and legend keep 6.5pt (the legend 9pt)
// from the chart's edge, the plot 11pt from a bare right edge; the title's text starts 10.4pt down
// and leaves 9.6pt above the plot; a bottom legend sits 7.7pt below the axis labels.
const PAD = 9;
const EDGE = 6.5;
const TITLE_TOP = 7.6;
const TITLE_GAP = 12.4;
/** Room between an axis title and its labels, and the extra margin Word leaves outside the title. */
const AXIS_TITLE_GAP = 2.6;
const AXIS_TITLE_EDGE = 6;
/** Where a manually placed title's text starts inside the box Word positions. */
const MANUAL_INSET = [3, 1.5] as const;
const LEGEND_GAP = 9;
const TICK = 3;

interface Text {
  runs: ShapedRun[];
  width: number;
  ascent: number;
  descent: number;
  height: number;
  /** How far the baseline sits below the centre of the em box. */
  mid: number;
}

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function measure(text: string, s: ChartTextStyle, env: BoxEnv): Text {
  const props = { ...DEFAULT_RUN, fontAscii: s.font, fontHAnsi: s.font, fontEastAsia: s.font, fontCs: s.font, size: s.size, sizeCs: s.size, bold: s.bold, boldCs: s.bold, italic: s.italic, italicCs: s.italic };
  const runs = text ? shapeText(text, props, env.fonts) : [];
  let width = 0;
  let ascent = 0;
  let descent = 0;
  let lead = 0;
  for (const r of runs) {
    for (const g of r.glyphs) width += g.advance;
    ascent = Math.max(ascent, r.ascent);
    descent = Math.max(descent, r.descent);
    lead = Math.max(lead, r.extLeading);
  }
  if (!runs.length) {
    ascent = s.size * 0.75;
    descent = s.size * 0.25;
    lead = s.size * 0.2;
  }
  const font = runs[0]?.face.face.font;
  const mid = font ? ((font.typoAscender + font.typoDescender) / 2 / font.unitsPerEm) * runs[0].size : s.size * 0.25;
  return { runs, width, ascent, descent, height: ascent + descent + lead, mid };
}

function glyphOps(t: Text, x: number, baseline: number, color: string): GlyphOp[] {
  const ops: GlyphOp[] = [];
  let px = x;
  for (const r of t.runs) {
    const op: GlyphOp = { kind: 'glyphs', x: px, y: baseline, face: r.face.face, size: r.size, gids: [], xs: [], texts: [], color, syntheticBold: r.face.syntheticBold, syntheticItalic: r.face.syntheticItalic, scaleX: 1 };
    let gx = 0;
    for (const g of r.glyphs) {
      if (g.gid >= 0) {
        op.gids.push(g.gid);
        op.xs.push(gx);
        op.texts.push(g.text);
      }
      gx += g.advance;
    }
    px += gx;
    ops.push(op);
  }
  return ops;
}

/**
 * Draw text so the point (px, py) sits at fraction (ax, ay) of its box (0 = left/top, 1 =
 * right/bottom), the box turned `rotation` degrees clockwise about that point.
 */
function place(t: Text, color: string, px: number, py: number, ax: number, ay: number, rotation = 0): DrawOp[] {
  // Edges align the line box, but Word centres a label on its em box (Calibri's baseline sits 0.25em below the
  // centre, not the line box's 0.34em).
  const baseline = ay <= 0.5 ? t.ascent + (t.mid - t.ascent) * ay * 2 : t.mid - (t.mid + t.descent) * (ay - 0.5) * 2;
  const body = glyphOps(t, px - ax * t.width, py + baseline, color);
  if (!rotation) return body;
  const a = (rotation * Math.PI) / 180;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  return [{ kind: 'transform', matrix: [cos, sin, -sin, cos, px - (cos * px - sin * py), py - (sin * px + cos * py)] }, ...body, { kind: 'restore' }];
}

/** The extent of a box of size (w, h) turned `rotation` degrees: [width, height]. */
function turned(w: number, h: number, rotation: number): [number, number] {
  const a = (Math.abs(rotation) * Math.PI) / 180;
  return [Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a)), Math.abs(w * Math.sin(a)) + Math.abs(h * Math.cos(a))];
}

function dashOf(l: ChartLine): number[] | undefined {
  const w = Math.max(l.width, 0.5);
  switch (l.dash) {
    case 'dot':
    case 'sysDot':
      return [w, w];
    case 'dash':
    case 'sysDash':
      return [w * 4, w * 3];
    case 'lgDash':
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

function stroke(segs: PathSeg[], l: ChartLine | null): DrawOp[] {
  if (!l || l.width <= 0) return [];
  return [{ kind: 'path', segs, fill: null, stroke: l.color, strokeWidth: l.width, ...(l.alpha !== undefined && { strokeAlpha: l.alpha }), ...(dashOf(l) && { dash: dashOf(l) }) }];
}

/**
 * A shape's outer shadow: its fill and outline again in the shadow colour, offset. A blur fades the
 * shadow out over its radius: four translucent layers, each outlined wider, stand in for it.
 */
function shadowOps(segs: PathSeg[], s: ChartShape): DrawOp[] {
  const sh = s.shadow;
  if (!sh || (!s.fill && !s.line)) return [];
  const moved = segs.map((g) => ({ ...g, pts: g.pts.map((v, i) => v + (i % 2 ? sh.dy : sh.dx)) }));
  const lw = s.line?.width ?? 0;
  const layer = (alpha: number, width: number): DrawOp => ({
    kind: 'path',
    segs: moved,
    fill: s.fill ? sh.color : null,
    ...(s.fill && alpha < 1 && { fillAlpha: alpha }),
    stroke: width > 0 ? sh.color : null,
    ...(width > 0 && { strokeWidth: width, ...(alpha < 1 && { strokeAlpha: alpha }) }),
  });
  if (sh.blur < 0.5) return [layer(sh.alpha, lw)];
  const a = 1 - Math.pow(1 - sh.alpha, 1 / 4);
  return [0, 1, 2, 3].map((k) => layer(a, lw + (sh.blur * 2 * k) / 3));
}

function paint(segs: PathSeg[], s: ChartShape): DrawOp[] {
  const ops: DrawOp[] = shadowOps(segs, s);
  if (s.fill?.gradient) ops.push(...gradientOps(segs, s.fill.gradient, s.fill.alpha));
  else if (s.fill) ops.push({ kind: 'path', segs, fill: s.fill.color, ...(s.fill.alpha < 1 && { fillAlpha: s.fill.alpha }), stroke: null });
  ops.push(...stroke(segs, s.line));
  return ops;
}

const rectPath = (x: number, y: number, w: number, h: number): PathSeg[] => [
  { op: 'M', pts: [x, y] },
  { op: 'L', pts: [x + w, y] },
  { op: 'L', pts: [x + w, y + h] },
  { op: 'L', pts: [x, y + h] },
  { op: 'Z', pts: [] },
];
const seg = (x1: number, y1: number, x2: number, y2: number): PathSeg[] => [
  { op: 'M', pts: [x1, y1] },
  { op: 'L', pts: [x2, y2] },
];

// ---------------------------------------------------------------------------------------------
// Number formats

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function general(v: number): string {
  if (v === 0) return '0';
  const a = Math.abs(v);
  if (a >= 1e11 || a < 1e-9) {
    const [m, e] = v.toExponential(5).split('e');
    const exp = Number(e);
    return `${Number(m)}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  // Excel shows up to about ten significant digits.
  return String(Number(v.toPrecision(10)));
}

function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') quoted = !quoted;
    if (c === '\\' && i + 1 < code.length) {
      cur += c + code[++i];
      continue;
    }
    if (c === ';' && !quoted) {
      out.push(cur);
      cur = '';
    } else cur += c;
  }
  out.push(cur);
  return out;
}

/** Literal text of a format fragment: quoted strings and escapes kept, codes like _) and *  dropped. */
function literal(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (c === '"') {
      const j = s.indexOf('"', i + 1);
      out += s.slice(i + 1, j < 0 ? s.length : j);
      i = j < 0 ? s.length : j;
    } else if (c === '\\') out += s[++i] ?? '';
    else if (c === '_' || c === '*') i++;
    else out += c;
  }
  return out;
}

function formatDate(serial: number, code: string): string {
  // Excel's dates run from serial 0 to the end of 9999; it shows hashes for values outside them.
  if (!(serial >= 0 && serial < 2958466)) return '#####';
  // Shown to the whole second, rounded (Excel's 0.0014 days is 12:02:01 AM).
  const ms = Math.round((serial - 25569) * 86400) * 1000;
  const d = new Date(ms);
  const tokens = code.match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|A\/P|a\/p|[^"\\ymdhsAa]+|./g) ?? [];
  const hour12 = tokens.some((t) => /^(AM\/PM|am\/pm|A\/P|a\/p)$/.test(t));
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    // m and mm are minutes after an hour or before seconds.
    const prev = tokens.slice(0, i).reverse().find((x) => /^[hsdmy]/.test(x));
    const next = tokens.slice(i + 1).find((x) => /^[hsdmy]/.test(x));
    const minutes = (t === 'm' || t === 'mm') && (prev?.startsWith('h') || next?.startsWith('s'));
    const h = d.getUTCHours();
    switch (t) {
      case 'yyyy':
        out += d.getUTCFullYear();
        break;
      case 'yy':
        out += String(d.getUTCFullYear() % 100).padStart(2, '0');
        break;
      case 'mmmmm':
        out += MONTHS[d.getUTCMonth()][0];
        break;
      case 'mmmm':
        out += MONTHS[d.getUTCMonth()];
        break;
      case 'mmm':
        out += MONTHS[d.getUTCMonth()].slice(0, 3);
        break;
      case 'mm':
        out += String(minutes ? d.getUTCMinutes() : d.getUTCMonth() + 1).padStart(2, '0');
        break;
      case 'm':
        out += minutes ? d.getUTCMinutes() : d.getUTCMonth() + 1;
        break;
      case 'dddd':
        out += DAYS[d.getUTCDay()];
        break;
      case 'ddd':
        out += DAYS[d.getUTCDay()].slice(0, 3);
        break;
      case 'dd':
        out += String(d.getUTCDate()).padStart(2, '0');
        break;
      case 'd':
        out += d.getUTCDate();
        break;
      case 'hh':
        out += String(hour12 ? h % 12 || 12 : h).padStart(2, '0');
        break;
      case 'h':
        out += hour12 ? h % 12 || 12 : h;
        break;
      case 'ss':
        out += String(d.getUTCSeconds()).padStart(2, '0');
        break;
      case 's':
        out += d.getUTCSeconds();
        break;
      case 'AM/PM':
      case 'am/pm':
        out += h < 12 ? 'AM' : 'PM';
        break;
      case 'A/P':
      case 'a/p':
        out += h < 12 ? 'A' : 'P';
        break;
      default:
        out += literal(t);
    }
  }
  return out;
}

/** Format a number with an Excel number format (the common codes: decimals, grouping, percent, dates). */
export function formatNumber(v: number, code: string | undefined): string {
  if (!Number.isFinite(v)) return '';
  const sections = splitSections((code ?? 'General').replace(/\[(\$[^\]]*|[A-Za-z]+\d*|<[^\]]*|>[^\]]*|=[^\]]*)\]/g, (m) => (/^\[\$([^-\]]+)/.test(m) ? /^\[\$([^-\]]+)/.exec(m)![1] : '')));
  let fmt = sections[0];
  let neg = false;
  if (v < 0 && sections.length > 1 && sections[1].trim() !== '') {
    fmt = sections[1];
    neg = true;
  } else if (v === 0 && sections.length > 2 && sections[2].trim() !== '') fmt = sections[2];
  const value = neg ? -v : v;
  if (/^\s*general\s*$/i.test(fmt) || fmt.trim() === '') return general(value);
  const bare = fmt.replace(/"[^"]*"|\\./g, '');
  if (/[ymdhs]/i.test(bare.replace(/General/gi, '')) && !/[0#?]/.test(bare)) return formatDate(value, fmt);
  const first = bare.search(/[0#?]/);
  if (first < 0) return literal(fmt);
  // The numeric core runs from the first to the last digit placeholder (with its point and commas).
  let start = -1;
  let end = -1;
  let quoted = false;
  for (let i = 0; i < fmt.length; i++) {
    const c = fmt[i];
    if (c === '"') quoted = !quoted;
    else if (c === '\\') i++;
    else if (!quoted && /[0#?]/.test(c)) {
      if (start < 0) start = i;
      end = i;
    }
  }
  const core = fmt.slice(start, end + 1);
  const prefix = literal(fmt.slice(0, start));
  const suffix = literal(fmt.slice(end + 1));
  let x = value;
  for (const c of prefix + suffix) if (c === '%') x *= 100;
  const exp = /E[+-]/i.test(suffix) || /E[+-]/i.test(core);
  if (exp) return prefix + general(x) + suffix;
  const dot = core.indexOf('.');
  const intPart = dot < 0 ? core : core.slice(0, dot);
  const fracPart = dot < 0 ? '' : core.slice(dot + 1);
  const decimals = (fracPart.match(/[0#?]/g) ?? []).length;
  const minFrac = (fracPart.match(/0/g) ?? []).length;
  const minInt = (intPart.match(/0/g) ?? []).length;
  const group = intPart.includes(',');
  let s = Math.abs(x).toFixed(decimals);
  let [ip, fp = ''] = s.split('.');
  while (fp.length > minFrac && fp.endsWith('0')) fp = fp.slice(0, -1);
  if (ip === '0' && minInt === 0) ip = '';
  ip = ip.padStart(minInt, '0');
  if (group) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  s = ip + (fp ? '.' + fp : '');
  if (!s) s = '0';
  const sign = x < 0 && !neg && Number(s.replace(/[^\d]/g, '')) !== 0 ? '-' : '';
  return sign + prefix + s + suffix;
}

// ---------------------------------------------------------------------------------------------
// Scales

interface Scale {
  min: number;
  max: number;
  unit: number;
  log?: number;
}

const NICE = [1, 2, 5, 10];

/**
 * Excel's automatic value axis: 5% headroom beyond the data, zero kept in unless the data sits far
 * from it, then the smallest 1-2-5 step that leaves no more than ten gridlines and keeps them about
 * one and a half lines of label text apart, with the ends rounded out to whole steps.
 */
function autoScale(dmin: number, dmax: number, axis: ChartAxis, length: number, spacing: number): Scale {
  if (axis.logBase) {
    const b = axis.logBase;
    const lo = axis.min !== undefined && axis.min > 0 ? axis.min : b ** Math.floor(Math.log(Math.max(dmin, 1e-300)) / Math.log(b) + 1e-9);
    const hi = axis.max ?? b ** Math.ceil(Math.log(Math.max(dmax, lo * b)) / Math.log(b) - 1e-9);
    const max = Math.max(hi, lo * b);
    if (!(lo > 0 && max > lo && Number.isFinite(max))) throw new Error('Invalid logarithmic chart range');
    return { min: lo, max, unit: axis.majorUnit ?? b, log: b };
  }
  if (!Number.isFinite(dmin) || !Number.isFinite(dmax)) {
    dmin = 0;
    dmax = 1;
  }
  if (dmin === dmax) {
    if (dmax === 0) dmax = 1;
    else if (dmax > 0) dmin = 0;
    else dmax = 0;
  }
  let lo: number;
  let hi: number;
  if (dmin >= 0) {
    lo = dmin > (dmax * 5) / 6 ? dmin - (dmax - dmin) / 2 : 0;
    hi = dmax + 0.05 * (dmax - lo);
  } else if (dmax <= 0) {
    hi = dmax < (dmin * 5) / 6 ? dmax + (dmax - dmin) / 2 : 0;
    lo = dmin - 0.05 * (hi - dmin);
  } else {
    hi = dmax + 0.05 * (dmax - dmin);
    lo = dmin - 0.05 * (dmax - dmin);
  }
  if (axis.min !== undefined) lo = axis.min;
  if (axis.max !== undefined) hi = axis.max;
  if (hi <= lo) hi = lo + 1;
  const range = hi - lo;
  if (!(range > 0 && Number.isFinite(range))) throw new Error('Chart axis range exceeds numeric limits');
  let unit = axis.majorUnit;
  if (!unit) {
    const most = Math.max(1, Math.min(10, length / spacing));
    const power = 10 ** Math.max(-323, Math.floor(Math.log10(range) - Math.log10(most)));
    unit = NICE.map((m) => m * power).find((c) => Number.isFinite(c) && range / c <= most + 1e-9) ?? range;
  }
  const snap = (n: number) => Math.round(n / unit! / 1e-9) * 1e-9 * unit!;
  const min = axis.min ?? snap(Math.floor(lo / unit + 1e-9) * unit);
  const max = axis.max ?? snap(Math.ceil(hi / unit - 1e-9) * unit);
  const upper = max > min ? max : min + unit;
  if (!(upper > min && Number.isFinite(upper - min))) throw new Error('Chart axis range exceeds numeric limits');
  return { min, max: upper, unit };
}

function ticks(s: Scale): number[] {
  const out: number[] = [];
  if (s.log) {
    for (let v = s.min; Number.isFinite(v) && v / s.max <= 1 + 1e-9; v *= s.log) {
      if (out.length >= 1000 || v * s.log <= v) return [s.min, s.max];
      out.push(v);
    }
    return out;
  }
  const n = Math.round((s.max - s.min) / s.unit);
  if (!Number.isFinite(n) || n > 1000) return [s.min, s.max];
  for (let i = 0; i <= n; i++) out.push(Number((s.min + i * s.unit).toPrecision(12)));
  return out;
}

// ---------------------------------------------------------------------------------------------
// Axes

interface AxisState {
  axis: ChartAxis;
  horizontal: boolean;
  /** Value scale (value axes) or category count (category axes). */
  scale?: Scale;
  count: number;
  labels: string[];
  format: string;
  /** Category label turn in degrees and wrapped lines, decided by fitting. */
  rotation: number;
  lines: string[][];
  skip: number;
  plots: ChartPlot[];
  /** Where the axis line sits (page coordinate across the axis direction). */
  at: number;
}

/** Fraction along an axis of a value (value axes) or a category index (category axes, centre of the band). */
function fraction(a: AxisState, v: number): number {
  let f: number;
  if (a.scale) {
    const s = a.scale;
    f = s.log ? (Math.log(Math.max(s.min, v)) - Math.log(s.min)) / (Math.log(s.max) - Math.log(s.min)) : (v - s.min) / (s.max - s.min);
  } else f = a.axis.between ? (v + 0.5) / a.count : a.count > 1 ? v / (a.count - 1) : 0.5;
  return a.axis.reverse ? 1 - f : f;
}

function coord(a: AxisState, plot: Rect, v: number): number {
  const f = fraction(a, v);
  return a.horizontal ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h;
}

/** The value or category edge at which an axis crosses the other one. */
function crossing(a: AxisState, other: AxisState, plot: Rect): number {
  const c = a.axis.crosses;
  if (other.scale) {
    const s = other.scale;
    const v = c === 'min' ? s.min : c === 'max' ? s.max : c === 'autoZero' ? Math.min(Math.max(0, s.min), s.max) : c;
    return coord(other, plot, s.log && v <= 0 ? s.min : v);
  }
  // Across categories: the first edge, the last, or a category's middle.
  const atEnd = c === 'max';
  if (typeof c === 'number') return coord(other, plot, c - 1);
  let f = atEnd ? 1 : 0;
  if (other.axis.reverse) f = 1 - f;
  return other.horizontal ? plot.x + f * plot.w : plot.y + plot.h - f * plot.h;
}

function wrapWords(text: string, width: number, style: ChartTextStyle, env: BoxEnv): string[] | null {
  const words = text.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let cur = '';
  for (const w of words) {
    if (measure(w, style, env).width > width) return null;
    const next = cur ? `${cur} ${w}` : w;
    if (cur && measure(next, style, env).width > width) {
      lines.push(cur);
      cur = w;
    } else cur = next;
  }
  if (cur) lines.push(cur);
  return lines;
}

/** Decide how a horizontal category axis shows its labels in bands `band` wide: as they are, wrapped, or turned. */
function fitCategoryLabels(a: AxisState, band: number, env: BoxEnv): void {
  const st = a.axis.style;
  a.lines = a.labels.map((l) => [l]);
  a.rotation = st.rotation ?? 0;
  a.skip = a.axis.labelSkip ?? 1;
  if (!a.horizontal || st.rotation !== undefined) return;
  const widths = a.labels.map((l) => measure(l, st, env).width);
  if (widths.every((w) => w <= band - 2)) return;
  // Excel wraps labels at spaces when every word fits the band, keeping them to a few lines. A wrapped word may
  // fill the whole band (Word wraps a 48.7pt word in a 50.6pt band rather than turning the labels), but labels
  // with no spaces at all (dates, codes) are turned instead.
  const wrapped = a.labels.map((l) => wrapWords(l, band, st, env));
  if (a.labels.some((l) => /\s/.test(l.trim())) && wrapped.every((w) => w && w.length <= 3)) {
    a.lines = wrapped as string[][];
    return;
  }
  a.rotation = -45;
  const lineH = measure('0', st, env).height;
  if (a.axis.labelSkip === undefined) a.skip = Math.max(1, Math.ceil((lineH * 1.1) / (band * Math.SQRT1_2)));
}

/** Width (vertical axes) or height (horizontal axes) the labels need beside the axis. */
function labelExtent(a: AxisState, env: BoxEnv): number {
  if (a.axis.deleted || a.axis.labelPosition === 'none') return 0;
  const st = a.axis.style;
  const lineH = measure('0', st, env).height;
  let most = 0;
  for (const lines of a.lines) {
    const w = Math.max(0, ...lines.map((l) => measure(l, st, env).width));
    const h = lines.length * lineH;
    // A turned label hangs from the centre of its end, so only half its line height reaches past that point.
    const [tw, th] = a.rotation ? turned(w, lineH / 2, a.rotation) : turned(w, h, 0);
    most = Math.max(most, a.horizontal ? th : tw);
  }
  return most;
}

function titleExtent(t: ChartTitle | undefined, env: BoxEnv): [number, number, number] {
  if (!t) return [0, 0, 0];
  let w = 0;
  let h = 0;
  let rot = 0;
  for (const p of t.paragraphs) {
    let pw = 0;
    let ph = 0;
    for (const r of p) {
      const m = measure(r.text, r.style, env);
      pw += m.width;
      ph = Math.max(ph, m.height);
      rot = r.style.rotation ?? rot;
    }
    const first = t.paragraphs.find((q) => q.length)?.[0].style;
    if (!p.length && first) ph = measure('', first, env).height;
    w = Math.max(w, pw);
    h += ph;
  }
  return [w, h, rot];
}

/** Draw a title's paragraphs centred about (cx, top), turned `rot` about its centre. */
function drawTitle(t: ChartTitle, cx: number, cy: number, env: BoxEnv): DrawOp[] {
  const [w, h, rot] = titleExtent(t, env);
  const ops: DrawOp[] = [];
  let y = cy - h / 2;
  const body: DrawOp[] = [];
  const first = t.paragraphs.find((p) => p.length)?.[0].style;
  for (const p of t.paragraphs) {
    const ms = p.map((r) => measure(r.text, r.style, env));
    const pw = ms.reduce((s, m) => s + m.width, 0);
    const asc = Math.max(0, ...ms.map((m) => m.ascent));
    // An empty paragraph keeps a line of the title's text.
    const ph = ms.length ? Math.max(...ms.map((m) => m.height)) : first ? measure('', first, env).height : 0;
    // Word centres a title line without its trailing spaces.
    const last = p[p.length - 1];
    const trail = last && last.text !== last.text.trimEnd() ? ms[ms.length - 1].width - measure(last.text.trimEnd(), last.style, env).width : 0;
    let x = cx - (pw - trail) / 2;
    for (const [i, m] of ms.entries()) {
      body.push(...glyphOps(m, x, y + asc, p[i].style.color));
      x += m.width;
    }
    y += ph;
  }
  if (t.shape.fill || t.shape.line) ops.push(...paint(rectPath(cx - w / 2 - 3, cy - h / 2 - 2, w + 6, h + 4), t.shape));
  if (!rot) return [...ops, ...body];
  const a = (rot * Math.PI) / 180;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  return [...ops, { kind: 'transform', matrix: [cos, sin, -sin, cos, cx - (cos * cx - sin * cy), cy - (sin * cx + cos * cy)] }, ...body, { kind: 'restore' }];
}

// ---------------------------------------------------------------------------------------------
// Series geometry

const isPie = (p: ChartPlot) => p.type === 'pie' || p.type === 'doughnut';
const isCartesian = (p: ChartPlot) => !isPie(p) && p.type !== 'radar';

/** Catmull-Rom style Bézier through the points, as Excel smooths lines. */
function smoothPath(pts: Array<[number, number]>): PathSeg[] {
  const segs: PathSeg[] = [{ op: 'M', pts: pts[0] }];
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] ?? p2;
    segs.push({ op: 'C', pts: [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6, p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6, p2[0], p2[1]] });
  }
  return segs;
}

/** Highest and lowest points a smoothed line reaches between its data points (Excel scales to them). */
function smoothRange(values: Array<number | null>): [number, number] {
  let lo = Infinity;
  let hi = -Infinity;
  const pts = values.map((v, i) => [i, v] as const).filter((p): p is readonly [number, number] => p[1] !== null);
  for (let i = 0; i < pts.length - 1; i++) {
    const y0 = (pts[i - 1] ?? pts[i])[1];
    const y1 = pts[i][1];
    const y2 = pts[i + 1][1];
    const y3 = (pts[i + 2] ?? pts[i + 1])[1];
    const c1 = y1 + (y2 - y0) / 6;
    const c2 = y2 - (y3 - y1) / 6;
    for (let k = 1; k < 16; k++) {
      const t = k / 16;
      const y = (1 - t) ** 3 * y1 + 3 * (1 - t) ** 2 * t * c1 + 3 * (1 - t) * t ** 2 * c2 + t ** 3 * y2;
      lo = Math.min(lo, y);
      hi = Math.max(hi, y);
    }
  }
  return [lo, hi];
}

const TREND_NAMES: Record<string, string> = { linear: 'Linear', exp: 'Expon.', log: 'Log.', poly: 'Poly.', power: 'Power' };

/** Solve the square system A·x = b by Gaussian elimination with partial pivoting; null when singular. */
function solve(A: number[][], b: number[]): number[] | null {
  const n = b.length;
  const m = A.map((row, i) => [...row, b[i]]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(m[r][c]) > Math.abs(m[p][c])) p = r;
    if (Math.abs(m[p][c]) < 1e-12) return null;
    [m[c], m[p]] = [m[p], m[c]];
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = m[r][c] / m[c][c];
      for (let k = c; k <= n; k++) m[r][k] -= f * m[c][k];
    }
  }
  return m.map((row, i) => row[n] / row[i]);
}

/**
 * Points along a trendline through a series' (x, y) points, in data space. Fits are least squares, as
 * Excel's: exponential and power ones on ln y, logarithmic and power ones on ln x, so they skip points
 * those can't take. The curve runs from the first x to the last, plus the forecast either side.
 */
function trendPoints(t: ChartTrendline, data: Array<[number, number]>): Array<[number, number]> {
  if (t.type === 'movingAvg') {
    const p = Math.max(2, Math.round(t.order));
    const out: Array<[number, number]> = [];
    for (let i = p - 1; i < data.length; i++) out.push([data[i][0], data.slice(i - p + 1, i + 1).reduce((sum, d) => sum + d[1], 0) / p]);
    return out;
  }
  const logX = t.type === 'log' || t.type === 'power';
  const logY = t.type === 'exp' || t.type === 'power';
  const pts = data.filter(([x, y]) => (!logX || x > 0) && (!logY || y > 0));
  const degree = t.type === 'poly' ? Math.max(2, Math.min(6, Math.round(t.order))) : 1;
  if (pts.length <= degree) return [];
  // A set intercept fixes the constant term (ln of it for an exponential fit).
  const fixed = t.intercept === undefined || logX ? undefined : logY ? (t.intercept > 0 ? Math.log(t.intercept) : undefined) : t.intercept;
  const us = pts.map(([x]) => (logX ? Math.log(x) : x));
  // Without a fixed constant, centre and scale u: the same curve, but a high-order fit stays well conditioned.
  const c = fixed === undefined ? us.reduce((a, u) => a + u, 0) / us.length : 0;
  const k = fixed === undefined ? Math.max(...us.map((u) => Math.abs(u - c))) || 1 : 1;
  const first = fixed === undefined ? 0 : 1;
  const n = degree + 1 - first;
  const A = Array.from({ length: n }, () => new Array<number>(n).fill(0));
  const b = new Array<number>(n).fill(0);
  pts.forEach(([, y], i) => {
    const u = (us[i] - c) / k;
    const v = (logY ? Math.log(y) : y) - (fixed ?? 0);
    const basis = Array.from({ length: n }, (_, j) => u ** (j + first));
    for (let r = 0; r < n; r++) {
      b[r] += basis[r] * v;
      for (let q = 0; q < n; q++) A[r][q] += basis[r] * basis[q];
    }
  });
  const coef = solve(A, b);
  if (!coef) return [];
  const f = (x: number) => {
    const u = ((logX ? Math.log(x) : x) - c) / k;
    const y = coef.reduce((sum, a, j) => sum + a * u ** (j + first), fixed ?? 0);
    return logY ? Math.exp(y) : y;
  };
  const xs = pts.map(([x]) => x);
  let lo = Math.min(...xs) - t.backward;
  const hi = Math.max(...xs) + t.forward;
  if (logX && lo <= 0) lo = Math.min(...xs);
  const steps = degree === 1 && !logX && !logY ? 1 : 96;
  return Array.from({ length: steps + 1 }, (_, i) => {
    const x = lo + ((hi - lo) * i) / steps;
    return [x, f(x)] as [number, number];
  }).filter(([, y]) => Number.isFinite(y));
}

/**
 * Each point's error bar as its low and high ends, in the units of the values it was given (null
 * where the point has no value).
 */
function errorEnds(eb: ChartErrorBars, values: Array<number | null>): Array<[number, number] | null> {
  const vals = values.filter((v): v is number => v !== null);
  const n = vals.length;
  const mean = n ? vals.reduce((t, v) => t + v, 0) / n : 0;
  const sd = n > 1 ? Math.sqrt(vals.reduce((t, v) => t + (v - mean) ** 2, 0) / (n - 1)) : 0;
  // Excel's documented standard error is of the values themselves, not of their spread from the mean.
  const se = n > 1 ? Math.sqrt(vals.reduce((t, v) => t + v * v, 0) / (n * (n - 1))) : 0;
  return values.map((v, i) => {
    if (v === null) return null;
    let [minus, plus, centre] = [eb.value, eb.value, v];
    if (eb.valType === 'percentage') minus = plus = (Math.abs(v) * eb.value) / 100;
    else if (eb.valType === 'stdErr') minus = plus = se;
    else if (eb.valType === 'cust') [minus, plus] = [eb.minus?.[i] ?? 0, eb.plus?.[i] ?? 0];
    else if (eb.valType === 'stdDev') {
      // Standard deviation bars stand on the series' mean rather than on each point.
      minus = plus = sd * eb.value;
      centre = mean;
    }
    return [centre - (eb.type === 'plus' ? 0 : minus), centre + (eb.type === 'minus' ? 0 : plus)];
  });
}

/** Half the length of an error bar's end cap. */
const CAP = 3;

/** An error bar between page points `lo` and `hi`, running up the page or across it, capped at its ends. */
function errorBarOps(eb: ChartErrorBars, lo: [number, number], hi: [number, number], vertical: boolean): DrawOp[] {
  if (!eb.line) return [];
  const segs: PathSeg[] = seg(lo[0], lo[1], hi[0], hi[1]);
  // A one-sided bar has no cap where it leaves the point.
  const caps = eb.endCap ? [...(eb.type === 'plus' ? [] : [lo]), ...(eb.type === 'minus' ? [] : [hi])] : [];
  for (const [x, y] of caps) segs.push(...(vertical ? seg(x - CAP, y, x + CAP, y) : seg(x, y - CAP, x, y + CAP)));
  return stroke(segs, eb.line);
}

function markerPath(symbol: string, cx: number, cy: number, size: number): PathSeg[] {
  const r = size / 2;
  switch (symbol) {
    case 'square':
      return rectPath(cx - r, cy - r, size, size);
    case 'diamond':
      return [{ op: 'M', pts: [cx, cy - r] }, { op: 'L', pts: [cx + r, cy] }, { op: 'L', pts: [cx, cy + r] }, { op: 'L', pts: [cx - r, cy] }, { op: 'Z', pts: [] }];
    case 'triangle':
      return [{ op: 'M', pts: [cx, cy - r] }, { op: 'L', pts: [cx + r, cy + r] }, { op: 'L', pts: [cx - r, cy + r] }, { op: 'Z', pts: [] }];
    case 'x':
      return [...seg(cx - r, cy - r, cx + r, cy + r), ...seg(cx - r, cy + r, cx + r, cy - r)];
    case 'plus':
      return [...seg(cx - r, cy, cx + r, cy), ...seg(cx, cy - r, cx, cy + r)];
    case 'star':
      return [...seg(cx - r, cy - r, cx + r, cy + r), ...seg(cx - r, cy + r, cx + r, cy - r), ...seg(cx, cy - r, cx, cy + r)];
    case 'dash':
      return rectPath(cx - r, cy - size / 10, size, size / 5);
    case 'dot':
      return rectPath(cx - size / 10, cy - size / 10, size / 5, size / 5);
    default: {
      const k = 0.5522847498 * r;
      return [
        { op: 'M', pts: [cx + r, cy] },
        { op: 'C', pts: [cx + r, cy + k, cx + k, cy + r, cx, cy + r] },
        { op: 'C', pts: [cx - k, cy + r, cx - r, cy + k, cx - r, cy] },
        { op: 'C', pts: [cx - r, cy - k, cx - k, cy - r, cx, cy - r] },
        { op: 'C', pts: [cx + k, cy - r, cx + r, cy - k, cx + r, cy] },
        { op: 'Z', pts: [] },
      ];
    }
  }
}

function drawMarker(m: NonNullable<ChartSeries['marker']>, x: number, y: number): DrawOp[] {
  const segs = markerPath(m.symbol, x, y, m.size);
  // Crosses and pluses are strokes only.
  if (m.symbol === 'x' || m.symbol === 'plus' || m.symbol === 'star') return stroke(segs, m.shape.line ?? (m.shape.fill && { color: m.shape.fill.color, width: 0.75 }));
  return paint(segs, m.shape);
}

function labelText(l: ChartLabels, s: ChartSeries, i: number, value: number, percent: number | undefined, category: string): string {
  const parts: string[] = [];
  if (l.series) parts.push(s.name);
  if (l.category) parts.push(category);
  if (l.value) parts.push(formatNumber(value, l.format ?? s.format));
  if (l.percent && percent !== undefined) parts.push(formatNumber(percent, '0%'));
  return parts.join(l.separator ?? ', ');
}

function pointShape(s: ChartSeries, i: number): ChartShape {
  return s.points.get(i) ?? s.shape;
}

// ---------------------------------------------------------------------------------------------
// Chart

/** Draw a chart in the box (x, y, w, h). */
export function renderChart(chart: Chart, X: number, Y: number, W: number, H: number, env: BoxEnv): DrawOp[] {
  const ops: DrawOp[] = [];
  ops.push({ kind: 'clip', x: X - 1, y: Y - 1, w: W + 2, h: H + 2 });
  ops.push(...paint(rectPath(X, Y, W, H), chart.space));
  let top = Y + (chart.title ? TITLE_TOP : 11);
  let bottom = Y + H - EDGE;
  let left = X + EDGE;
  let right = X + W - 11;

  // Title.
  const titleOps: DrawOp[] = [];
  if (chart.title) {
    const [tw, th] = titleExtent(chart.title, env);
    const tx = chart.title.x !== undefined ? X + chart.title.x * W + MANUAL_INSET[0] + tw / 2 : X + W / 2;
    const ty = chart.title.y !== undefined ? Y + chart.title.y * H + MANUAL_INSET[1] + th / 2 : top + th / 2;
    titleOps.push(...drawTitle(chart.title, tx, ty, env));
    if (!chart.title.overlay && chart.title.y === undefined) top += th + TITLE_GAP;
  }

  // Legend entries.
  interface Entry {
    text: Text;
    shape: ChartShape;
    line: boolean;
    marker?: ChartSeries['marker'];
  }
  const entries: Entry[] = [];
  const legend = chart.legend;
  if (legend) {
    const trended = chart.plots.some((p) => isCartesian(p) && p.series.some((s) => s.trendlines?.length));
    for (const p of chart.plots) {
      // Pies, and a lone series in varied colours, list their categories rather than their series,
      // unless the chart has a trendline to list too.
      if (isPie(p) || (p.series.length === 1 && p.varyColors && !trended)) {
        const s = p.series[0];
        if (!s) continue;
        s.categories.forEach((c, i) => {
          if (!legend.hidden.includes(i)) entries.push({ text: measure(c, legend.style, env), shape: pointShape(s, i), line: false });
        });
        continue;
      }
      for (const s of p.series) {
        const n = entries.length;
        if (legend.hidden.includes(n)) {
          entries.push({ text: measure('', legend.style, env), shape: s.shape, line: false });
          continue;
        }
        // A series whose line is hidden shows only its marker in the legend.
        const line = (p.type === 'line' || (p.type === 'scatter' && p.lines) || (p.type === 'radar' && p.radarStyle !== 'filled')) && s.shape.line !== null;
        entries.push({ text: measure(s.name, legend.style, env), shape: s.shape, line, marker: s.marker });
      }
    }
    // Trendlines follow every series, each named for its fit and series unless given a name.
    for (const p of chart.plots)
      for (const s of isCartesian(p) && p.grouping !== 'stacked' && p.grouping !== 'percentStacked' ? p.series : [])
        for (const t of s.trendlines ?? []) {
          const name = t.name ?? `${t.type === 'movingAvg' ? `${Math.round(t.order)} per. Mov. Avg.` : (Object.hasOwn(TREND_NAMES, t.type) ? TREND_NAMES[t.type] : 'Linear')} (${s.name})`;
          entries.push({ text: measure(legend.hidden.includes(entries.length) ? '' : name, legend.style, env), shape: { fill: null, line: t.line }, line: true });
        }
  }
  const visible = entries.filter((e, i) => !legend?.hidden.includes(i) || e.text.runs.length);
  const legendOps: DrawOp[] = [];
  if (legend && visible.length) {
    const ls = legend.style;
    const lineH = measure('0', ls, env).height;
    // Word's keys are 0.54em squares, but once a line series is listed every key takes the line key's width.
    const keyH = ls.size * 0.54;
    const wide = visible.some((e) => e.line);
    const keyW = (e: Entry) => (e.line || wide ? ls.size * 2.1 : e.marker && e.marker.symbol !== 'none' ? e.marker.size : keyH);
    const entryW = (e: Entry) => keyW(e) + 2.05 + e.text.width;
    const gap = ls.size * 0.92;
    const drawEntry = (e: Entry, x: number, cy: number) => {
      const kw = keyW(e);
      if (e.line) {
        legendOps.push(...stroke(seg(x, cy, x + kw, cy), e.shape.line));
        if (e.marker) legendOps.push(...drawMarker(e.marker, x + kw / 2, cy));
      } else if (e.marker && e.marker.symbol !== 'none' && !wide) legendOps.push(...drawMarker(e.marker, x + kw / 2, cy));
      else legendOps.push(...paint(rectPath(x, cy - keyH / 2, kw, keyH), e.shape.fill ? e.shape : { fill: e.shape.line && { color: e.shape.line.color, alpha: 1 }, line: null }));
      legendOps.push(...place(e.text, ls.color, x + kw + 2.05, cy, 0, 0.5));
    };
    const horizontal = legend.position === 'b' || legend.position === 't';
    if (horizontal) {
      // Rows of entries, centred, wrapping when wider than the chart.
      const rows: Entry[][] = [[]];
      let rowW = 0;
      for (const e of visible) {
        const w = entryW(e);
        if (rows[rows.length - 1].length && rowW + gap + w > right - left) {
          rows.push([]);
          rowW = 0;
        }
        rowW += (rows[rows.length - 1].length ? gap : 0) + w;
        rows[rows.length - 1].push(e);
      }
      const blockH = rows.length * lineH;
      let y0 = legend.position === 'b' ? Y + H - PAD - blockH : top;
      if (legend.layout) y0 = Y + legend.layout.y * H;
      for (const row of rows) {
        const w = row.reduce((s, e, i) => s + entryW(e) + (i ? gap : 0), 0);
        let x = legend.layout ? X + legend.layout.x * W + (legend.layout.w * W - w) / 2 : X + (W - w) / 2;
        for (const e of row) {
          drawEntry(e, x, y0 + lineH / 2);
          x += entryW(e) + gap;
        }
        y0 += lineH;
      }
      if (!legend.overlay && !legend.layout) {
        if (legend.position === 'b') bottom = Y + H - PAD - blockH - LEGEND_GAP;
        else top += blockH + LEGEND_GAP;
      }
    } else {
      const w = Math.max(...visible.map(entryW));
      const blockH = visible.length * lineH;
      let x0 = legend.position === 'l' ? left : right - w;
      // A side legend is centred between 11pt in from the top (or below the title) and 11pt up from the bottom.
      let y0 = legend.position === 'tr' ? top : (top + Y + H - 11) / 2 - blockH / 2;
      if (legend.layout) {
        x0 = X + legend.layout.x * W;
        y0 = Y + legend.layout.y * H;
      }
      for (const e of visible) {
        drawEntry(e, x0, y0 + lineH / 2);
        y0 += lineH;
      }
      if (!legend.overlay && !legend.layout) {
        if (legend.position === 'l') left += w + 6;
        else right -= w + 6;
      }
    }
  }

  // Plot area and axes.
  const cartesian = chart.plots.filter(isCartesian);
  const axes = new Map<string, AxisState>();
  for (const ax of chart.axes) {
    const plots = cartesian.filter((p) => p.axes.includes(ax.id));
    if (!plots.length) continue;
    axes.set(ax.id, { axis: ax, horizontal: ax.position === 'b' || ax.position === 't', count: 0, labels: [], format: ax.format, rotation: 0, lines: [], skip: 1, plots, at: 0 });
  }
  // Which axis of a plot is its category (or x) axis: the one its partner crosses.
  const pairs: Array<{ plot: ChartPlot; cat: AxisState; val: AxisState }> = [];
  for (const p of cartesian) {
    const [a, b] = p.axes.map((id) => axes.get(id));
    if (!a || !b) continue;
    const scatter = p.type === 'scatter';
    const cat = scatter ? (a.horizontal ? a : b) : a.axis.kind === 'val' ? b : a;
    const valAx = cat === a ? b : a;
    pairs.push({ plot: p, cat, val: valAx });
  }
  for (const { plot: p, cat } of pairs) {
    if (p.type === 'scatter') continue;
    const n = Math.max(0, ...p.series.map((s) => Math.max(s.values.length, s.categories.length)));
    cat.count = Math.max(cat.count, n);
    if (!cat.labels.length) {
      const s = p.series.find((x) => x.categories.length) ?? p.series[0];
      const fmt = cat.axis.sourceLinked ? (s?.categoryFormat ?? 'General') : cat.axis.format;
      cat.labels = Array.from({ length: n }, (_, i) => {
        const num = s?.categoryNumbers?.[i];
        if (num !== undefined && num !== null) return formatNumber(num, fmt);
        return s?.categories[i] ?? String(i + 1);
      });
    }
  }
  // Data ranges of the value axes (and scatter x axes).
  const ranges = new Map<AxisState, [number, number]>();
  const trends: Array<{ t: ChartTrendline; pts: Array<[number, number]>; cat: AxisState; val: AxisState; scatter: boolean }> = [];
  const widen = (a: AxisState, lo: number, hi: number) => {
    const r = ranges.get(a) ?? [Infinity, -Infinity];
    ranges.set(a, [Math.min(r[0], lo), Math.max(r[1], hi)]);
  };
  for (const { plot: p, cat, val: v } of pairs) {
    if (p.grouping === 'percentStacked') {
      widen(v, 0, 1);
      const anyNeg = p.series.some((s) => s.values.some((x) => x !== null && x < 0));
      if (anyNeg) widen(v, -1, 1);
      v.format = '0%';
      continue;
    }
    if (p.grouping === 'stacked') {
      const n = cat.count;
      for (let i = 0; i < n; i++) {
        let pos = 0;
        let neg = 0;
        for (const s of p.series) {
          const x = s.values[i];
          if (x === null || x === undefined) continue;
          if (x >= 0) pos += x;
          else neg += x;
        }
        widen(v, neg, pos);
      }
    } else
      for (const s of p.series) {
        const vals = s.values.filter((x): x is number => x !== null);
        if (vals.length) widen(v, Math.min(...vals), Math.max(...vals));
        if (s.smooth && p.type !== 'scatter') {
          const [lo, hi] = smoothRange(s.values);
          if (Number.isFinite(lo)) widen(v, lo, hi);
        }
        if (p.type === 'scatter' && s.xValues) {
          const xs = s.xValues.filter((x): x is number => x !== null);
          if (xs.length) widen(cat, Math.min(...xs), Math.max(...xs));
        }
        // Excel scales the axes to take in the series' error bars and trendlines too.
        for (const eb of s.errorBars ?? []) {
          const x = eb.dir === 'x';
          const ends = errorEnds(eb, x ? (s.xValues ?? []) : s.values).filter((e): e is [number, number] => e !== null);
          if (ends.length) widen(x ? cat : v, Math.min(...ends.map((e) => e[0])), Math.max(...ends.map((e) => e[1])));
        }
        const data: Array<[number, number]> = [];
        s.values.forEach((y, i) => {
          const x = p.type === 'scatter' ? s.xValues?.[i] : i + 1;
          if (y !== null && x !== null && x !== undefined) data.push([x, y]);
        });
        for (const t of s.trendlines ?? []) {
          const pts = trendPoints(t, data);
          if (!pts.length) continue;
          trends.push({ t, pts, cat, val: v, scatter: p.type === 'scatter' });
          widen(v, Math.min(...pts.map((q) => q[1])), Math.max(...pts.map((q) => q[1])));
          if (p.type === 'scatter') widen(cat, pts[0][0], pts[pts.length - 1][0]);
        }
      }
    if (v.axis.sourceLinked && v.format === 'General') v.format = p.series[0]?.format ?? 'General';
    if (p.type === 'scatter' && cat.axis.sourceLinked && cat.format === 'General') cat.format = p.series[0]?.categoryFormat ?? 'General';
  }

  // Series with no values leave their axis no data range. It still gets a scale, the one a lone 1 would
  // give (from 0, or from 1 when logarithmic), so the chart draws, empty.
  for (const { val: v } of pairs) if (!ranges.has(v)) ranges.set(v, [1, 1]);
  const valueAxes = [...axes.values()].filter((a) => ranges.has(a));
  const scaleAll = (plot: Rect) => {
    for (const a of valueAxes) {
      const [lo, hi] = ranges.get(a)!;
      const len = a.horizontal ? plot.w : plot.h;
      const probe = measure(formatNumber(hi, a.format), a.axis.style, env).width;
      const spacing = a.horizontal ? Math.max(1.5 * a.axis.style.size, probe + a.axis.style.size) : 1.5 * a.axis.style.size;
      a.scale = autoScale(lo, hi, a.axis, len, spacing);
      a.labels = ticks(a.scale).map((t) => formatNumber(t, a.format));
      a.lines = a.labels.map((l) => [l]);
      a.rotation = a.axis.style.rotation ?? 0;
    }
  };
  // Labels keep a share of their size from the axis line; turned labels also keep a tick's length further off,
  // whether or not ticks are drawn (measured against Word with and without).
  const gapOf = (a: AxisState) => (a.horizontal ? (a.rotation ? TICK + a.axis.style.size * 0.85 : a.axis.style.size * 0.62) : a.axis.style.size * 0.9);
  const titleSize = (a: AxisState) => {
    if (!a.axis.title || a.axis.deleted) return 0;
    const [tw, th, rot] = titleExtent(a.axis.title, env);
    const [w, h] = turned(tw, th, rot);
    // A title Word was told where to put takes no room from the plot.
    if (a.axis.title.x !== undefined && a.axis.title.y !== undefined) return 0;
    return (a.horizontal ? h : w) + AXIS_TITLE_GAP + AXIS_TITLE_EDGE;
  };
  const shown = (a: AxisState) => !a.axis.deleted;

  let plot: Rect = { x: left, y: top, w: Math.max(1, right - left), h: Math.max(1, bottom - top) };
  if (pairs.length) {
    const fixed = chart.plotLayout;
    const outer: Rect = fixed ? { x: X + fixed.x * W, y: Y + fixed.y * H, w: fixed.w * W, h: fixed.h * H } : plot;
    let inner = { ...outer };
    for (let pass = 0; pass < 3; pass++) {
      scaleAll(inner);
      for (const a of axes.values()) if (!a.scale && a.count) fitCategoryLabels(a, (a.horizontal ? inner.w : inner.h) / Math.max(1, a.count), env);
      if (fixed?.inner) break;
      const side = { l: 0, r: 0, t: 0, b: 0 };
      let overhang = 0;
      for (const a of axes.values()) {
        if (!shown(a)) continue;
        // Labels beside an axis that crosses inside the plot (a scatter's x axis at y = 0) keep only half a line
        // of room at its edge (measured on one Word chart).
        const o = axes.get(a.axis.crossAxis);
        const at = o ? crossing(a, o, inner) : undefined;
        const edge = { l: inner.x, r: inner.x + inner.w, t: inner.y, b: inner.y + inner.h }[a.axis.position];
        const inside = a.axis.labelPosition === 'nextTo' && at !== undefined && Math.abs(at - edge) > 0.5;
        const ext = labelExtent(a, env);
        const need = (inside ? (ext ? a.axis.style.size * 0.5 : 0) : ext + (ext ? gapOf(a) : 0)) + titleSize(a);
        side[a.axis.position] = Math.max(side[a.axis.position], need);
        // A horizontal value axis's last label is centred on the plot's right edge; Word keeps it clear of the margin.
        if (a.scale && a.horizontal && ext && a.labels.length) overhang = Math.max(overhang, measure(a.labels[a.labels.length - 1], a.axis.style, env).width / 2);
      }
      // The plot keeps 11pt from a bare right edge, but labels and titles there run out to the usual margin.
      const reach = side.r && !fixed ? right - (X + W - EDGE) : side.r ? 0 : overhang;
      inner = { x: outer.x + side.l, y: outer.y + side.t, w: Math.max(1, outer.w - side.l - side.r - reach), h: Math.max(1, outer.h - side.t - side.b) };
    }
    if (fixed?.inner) {
      // Word keeps a manually placed plot's labels inside the chart, narrowing the plot if it must.
      for (const a of axes.values()) {
        if (!shown(a) || a.horizontal) continue;
        const need = labelExtent(a, env) + gapOf(a) + titleSize(a);
        if (a.axis.position === 'l' && inner.x - need < X + EDGE) {
          const nx = X + EDGE + need;
          inner.w -= nx - inner.x;
          inner.x = nx;
        } else if (a.axis.position === 'r' && inner.x + inner.w + need > X + W - EDGE) inner.w = X + W - EDGE - need - inner.x;
      }
      scaleAll(inner);
    }
    plot = inner;
  } else if (chart.plotLayout) {
    const f = chart.plotLayout;
    plot = { x: X + f.x * W, y: Y + f.y * H, w: f.w * W, h: f.h * H };
  }

  ops.push(...paint(rectPath(plot.x, plot.y, plot.w, plot.h), chart.plotShape));

  // Where each axis line runs.
  const partner = (a: AxisState) => axes.get(a.axis.crossAxis);
  for (const a of axes.values()) {
    const o = partner(a);
    a.at = o ? crossing(a, o, plot) : a.horizontal ? plot.y + plot.h : plot.x;
  }

  // Gridlines, behind the series.
  for (const a of axes.values()) {
    const grid = a.axis.gridlines;
    if (!grid) continue;
    const positions = a.scale ? ticks(a.scale).map((t) => coord(a, plot, t)) : Array.from({ length: a.count + (a.axis.between ? 1 : 0) }, (_, i) => (a.axis.between ? (a.horizontal ? plot.x + (i / Math.max(1, a.count)) * plot.w : plot.y + plot.h - (i / Math.max(1, a.count)) * plot.h) : coord(a, plot, i)));
    for (const p of positions) ops.push(...stroke(a.horizontal ? seg(p, plot.y, p, plot.y + plot.h) : seg(plot.x, p, plot.x + plot.w, p), grid));
  }

  // Series.
  const labelOps: DrawOp[] = [];
  const seriesOps: DrawOp[] = [];
  const errorOps: DrawOp[] = [];
  const markerOps: DrawOp[] = [];
  for (const { plot: p, cat, val: v } of pairs) {
    const n = cat.count;
    const pos = (a: AxisState, x: number) => coord(a, plot, x);
    const base = Math.min(Math.max(0, v.scale!.min), v.scale!.max);
    const totals = p.grouping === 'percentStacked' ? new Float64Array(n) : undefined;
    if (totals) for (const series of p.series) series.values.forEach((value, i) => { totals[i] += Math.abs(value ?? 0); });
    if (p.type === 'bar') {
      const band = (cat.horizontal ? plot.w : plot.h) / Math.max(1, n);
      const stacked = p.grouping !== 'clustered';
      const k = stacked ? 1 : p.series.length;
      const o = Math.max(-1, Math.min(1, p.overlap / 100));
      const w = band / (k - (k - 1) * o + p.gapWidth / 100);
      const cluster = w * (k - (k - 1) * o);
      const posStack = new Array(n).fill(0);
      const negStack = new Array(n).fill(0);
      p.series.forEach((s, si) => {
        const errors = (totals ? [] : s.errorBars ?? []).map((eb) => ({ eb, ends: errorEnds(eb, s.values) }));
        for (let i = 0; i < n; i++) {
          const given = s.values[i];
          if (given === null || given === undefined) continue;
          let value: number = given;
          let from = base;
          if (stacked) {
            const total = totals?.[i] || 1;
            value /= total;
            if (value >= 0) {
              from = posStack[i];
              posStack[i] += value;
              value = posStack[i];
            } else {
              from = negStack[i];
              negStack[i] += value;
              value = negStack[i];
            }
          }
          const c0 = pos(cat, i) - (cat.horizontal ? 1 : -1) * (band / 2) * (cat.axis.reverse ? -1 : 1);
          const dir = (cat.horizontal ? 1 : -1) * (cat.axis.reverse ? -1 : 1);
          const start = c0 + dir * ((band - cluster) / 2 + (stacked ? 0 : si * w * (1 - o)));
          const a0 = pos(v, from);
          const a1 = pos(v, value);
          const r: Rect = cat.horizontal ? { x: Math.min(start, start + dir * w), y: Math.min(a0, a1), w, h: Math.abs(a1 - a0) } : { x: Math.min(a0, a1), y: Math.min(start, start + dir * w), w: Math.abs(a1 - a0), h: w };
          seriesOps.push(...paint(rectPath(r.x, r.y, r.w, r.h), pointShape(s, i)));
          // Error bars stand on the bar's end (a stack's top); a percentage stack has no scale for them.
          for (const { eb, ends } of errors) {
            const e = ends[i];
            if (!e) continue;
            const c = start + (dir * w) / 2;
            const at = (x: number): [number, number] => (cat.horizontal ? [c, pos(v, x + value - given)] : [pos(v, x + value - given), c]);
            errorOps.push(...errorBarOps(eb, at(e[0]), at(e[1]), cat.horizontal));
          }
          const l = s.labels;
          if (l && !l.hidden?.includes(i)) {
            const raw = s.values[i]!;
            const t = measure(labelText(l, s, i, raw, undefined, cat.labels[i] ?? ''), l.style, env);
            const where = l.position ?? (stacked ? 'ctr' : 'outEnd');
            const g = l.style.size * 0.4;
            const end = a1;
            // Which way the bar grows on the page (+1 down or right): a negative value or a reversed axis turns it.
            const out = (raw >= 0 ? 1 : -1) * (v.axis.reverse ? -1 : 1) * (cat.horizontal ? -1 : 1);
            const mid = cat.horizontal ? r.x + r.w / 2 : r.y + r.h / 2;
            const along = where === 'ctr' ? (a0 + a1) / 2 : where === 'inBase' ? a0 : end;
            const off = where === 'outEnd' ? g : where === 'inEnd' ? -g : where === 'inBase' ? g : 0;
            if (cat.horizontal) {
              const y = along + out * off;
              const ay = where === 'ctr' ? 0.5 : (out < 0) === (where === 'outEnd') ? 1 : 0;
              labelOps.push(...place(t, l.style.color, mid, y, 0.5, where === 'inBase' ? 1 - ay : ay));
            } else {
              const x = along + out * off;
              const ax = where === 'ctr' ? 0.5 : (out > 0) === (where === 'outEnd') ? 0 : 1;
              labelOps.push(...place(t, l.style.color, x, mid, where === 'inBase' ? 1 - ax : ax, 0.5));
            }
          }
        }
      });
    } else {
      const stacked = p.grouping === 'stacked' || p.grouping === 'percentStacked';
      const acc = new Array(n).fill(0);
      const prevLine: Array<[number, number]> = [];
      p.series.forEach((s, si) => {
        const pts: Array<[number, number] | null> = [];
        // Each point's value on the axis (a stack's running total) and x, for its error bars.
        const tops: Array<[number, number] | null> = [];
        const count = p.type === 'scatter' ? s.values.length : n;
        for (let i = 0; i < count; i++) {
          const given = s.values[i];
          if (given === null || given === undefined) {
            pts.push(stacked ? [pos(cat, i), pos(v, acc[i])] : null);
            continue;
          }
          let value: number = given;
          if (stacked) {
            const total = totals?.[i] || 1;
            acc[i] += value / total;
            value = acc[i];
          }
          const xv = p.type === 'scatter' ? s.xValues?.[i] : i;
          if (xv === null || xv === undefined) {
            pts.push(null);
            continue;
          }
          const px = pos(cat, xv);
          const py = pos(v, value);
          pts.push(cat.horizontal ? [px, py] : [py, px]);
          tops[i] = [xv, value];
        }
        if (p.grouping !== 'percentStacked')
          for (const eb of s.errorBars ?? []) {
            const alongX = eb.dir === 'x';
            errorEnds(eb, alongX ? (s.xValues ?? []) : s.values).forEach((e, i) => {
              const t = tops[i];
              if (!e || !t) return;
              const shift = alongX ? 0 : t[1] - s.values[i]!;
              const at = (x: number): [number, number] => {
                const [cx, vy] = alongX ? [pos(cat, x), pos(v, t[1])] : [pos(cat, t[0]), pos(v, x + shift)];
                return cat.horizontal ? [cx, vy] : [vy, cx];
              };
              errorOps.push(...errorBarOps(eb, at(e[0]), at(e[1]), cat.horizontal !== alongX));
            });
          }
        // Gaps split a line into runs.
        const runs: Array<Array<[number, number]>> = [[]];
        for (const pt of pts) {
          if (pt) runs[runs.length - 1].push(pt);
          else if (runs[runs.length - 1].length) runs.push([]);
        }
        if (p.type === 'area') {
          const top = pts.map((pt, i) => pt ?? [pos(cat, i), pos(v, base)]) as Array<[number, number]>;
          if (top.length) {
            const bottomLine = stacked && si > 0 ? [...prevLine].reverse() : [
              [top[top.length - 1][0], pos(v, base)] as [number, number],
              [top[0][0], pos(v, base)] as [number, number],
            ];
            const segs: PathSeg[] = [{ op: 'M', pts: top[0] }, ...top.slice(1).map((q): PathSeg => ({ op: 'L', pts: q })), ...bottomLine.map((q): PathSeg => ({ op: 'L', pts: q })), { op: 'Z', pts: [] }];
            seriesOps.push(...paint(segs, s.shape));
          }
          prevLine.length = 0;
          prevLine.push(...top);
        } else if (s.shape.line && (p.type !== 'scatter' || p.lines)) {
          for (const run of runs) {
            if (run.length < 2) continue;
            const segs = s.smooth ? smoothPath(run) : [{ op: 'M', pts: run[0] } as PathSeg, ...run.slice(1).map((q): PathSeg => ({ op: 'L', pts: q }))];
            seriesOps.push(...stroke(segs, s.shape.line));
          }
        }
        pts.forEach((pt, i) => {
          if (!pt) return;
          const own = s.points.get(i);
          if (s.marker) markerOps.push(...drawMarker(own ? { ...s.marker, shape: { fill: own.fill ?? s.marker.shape.fill, line: own.line ?? s.marker.shape.line } } : s.marker, pt[0], pt[1]));
          const l = s.labels;
          if (l && !l.hidden?.includes(i) && s.values[i] !== null) {
            const t = measure(labelText(l, s, i, s.values[i]!, undefined, cat.labels[i] ?? ''), l.style, env);
            const g = (s.marker ? s.marker.size / 2 : 0) + 2;
            const where = l.position ?? 'r';
            if (where === 't') labelOps.push(...place(t, l.style.color, pt[0], pt[1] - g, 0.5, 1));
            else if (where === 'b') labelOps.push(...place(t, l.style.color, pt[0], pt[1] + g, 0.5, 0));
            else if (where === 'l') labelOps.push(...place(t, l.style.color, pt[0] - g, pt[1], 1, 0.5));
            else if (where === 'ctr') labelOps.push(...place(t, l.style.color, pt[0], pt[1], 0.5, 0.5));
            else labelOps.push(...place(t, l.style.color, pt[0] + g, pt[1], 0, 0.5));
          }
        });
      });
    }
  }
  // Trendlines over the series; a category's x is its 1-based number.
  for (const { t, pts, cat, val: v, scatter } of trends) {
    const xy = pts.map(([x, y]): [number, number] => {
      const px = coord(cat, plot, scatter ? x : x - 1);
      const py = coord(v, plot, y);
      return cat.horizontal ? [px, py] : [py, px];
    });
    seriesOps.push(...stroke([{ op: 'M', pts: xy[0] }, ...xy.slice(1).map((q): PathSeg => ({ op: 'L', pts: q }))], t.line));
  }
  if (seriesOps.length || errorOps.length) {
    ops.push({ kind: 'clip', x: plot.x, y: plot.y, w: plot.w, h: plot.h });
    for (const op of seriesOps) ops.push(op);
    for (const op of errorOps) ops.push(op);
    ops.push({ kind: 'restore' });
  }

  // Pie and doughnut plots fill the plot area's largest centred circle.
  for (const p of chart.plots) {
    if (!isPie(p)) continue;
    const s = p.series[0];
    if (!s) continue;
    const total = s.values.reduce<number>((t, x) => t + Math.abs(x ?? 0), 0);
    if (!total) continue;
    const r = Math.min(plot.w, plot.h) / 2;
    const cx = plot.x + plot.w / 2;
    const cy = plot.y + plot.h / 2;
    const inner = p.type === 'doughnut' ? (r * p.holeSize) / 100 : 0;
    let angle = p.firstSliceAngle - 90;
    s.values.forEach((x, i) => {
      if (x === null || !x) return;
      const sweep = (Math.abs(x) / total) * 360;
      const shape = pointShape(s, i);
      const mid = ((angle + sweep / 2) * Math.PI) / 180;
      const ex = (s.explosion / 100) * r;
      const ox = cx + Math.cos(mid) * ex;
      const oy = cy + Math.sin(mid) * ex;
      ops.push(...paint(slice(ox, oy, r, inner, angle, sweep), shape));
      const l = s.labels;
      if (l && !l.hidden?.includes(i)) {
        const t = measure(labelText(l, s, i, x, Math.abs(x) / total, s.categories[i] ?? ''), l.style, env);
        const where = l.position ?? 'bestFit';
        const rr = where === 'outEnd' ? r + 4 : where === 'inEnd' ? r * 0.8 : where === 'ctr' ? (r + inner) / 2 : inner ? (r + inner) / 2 : r * 0.62;
        const lx = ox + Math.cos(mid) * rr;
        const ly = oy + Math.sin(mid) * rr;
        if (where === 'outEnd') labelOps.push(...place(t, l.style.color, lx, ly, Math.cos(mid) >= 0 ? 0 : 1, 0.5));
        else labelOps.push(...place(t, l.style.color, lx, ly, 0.5, 0.5));
      }
      angle += sweep;
    });
  }

  // Radar plots: categories on spokes clockwise from the top, values out from the centre.
  for (const p of chart.plots) {
    if (p.type !== 'radar') continue;
    const n = Math.max(0, ...p.series.map((s) => Math.max(s.values.length, s.categories.length)));
    const catAx = chart.axes.find((a) => p.axes.includes(a.id) && a.kind !== 'val');
    const valAx = chart.axes.find((a) => p.axes.includes(a.id) && a.kind === 'val');
    if (!n || !valAx) continue;
    const named = p.series.find((s) => s.categories.length) ?? p.series[0];
    const fmt = catAx && !catAx.sourceLinked ? catAx.format : (named?.categoryFormat ?? 'General');
    const names = Array.from({ length: n }, (_, i) => {
      const num = named?.categoryNumbers?.[i];
      return num !== undefined && num !== null ? formatNumber(num, fmt) : (named?.categories[i] ?? String(i + 1));
    });
    const catText = catAx && !catAx.deleted && catAx.labelPosition !== 'none' ? names.map((t) => measure(t, catAx.style, env)) : [];
    const gap = catAx ? catAx.style.size * 0.5 : 0;
    const lw = Math.max(0, ...catText.map((t) => t.width));
    const lh = Math.max(0, ...catText.map((t) => t.height));
    const r = Math.max(1, Math.min(plot.w / 2 - (lw && lw + gap), plot.h / 2 - (lh && lh + gap)));
    const [cx, cy] = [plot.x + plot.w / 2, plot.y + plot.h / 2];
    let low = Infinity;
    let high = -Infinity;
    for (const s of p.series) for (const value of s.values) {
      if (value !== null) {
        low = Math.min(low, value);
        high = Math.max(high, value);
      }
    }
    const scale = autoScale(low, high, valAx, r, 1.5 * valAx.style.size);
    const angle = (i: number) => ((-90 + (360 * i) / n) * Math.PI) / 180;
    const at = (i: number, v: number): [number, number] => {
      const d = r * Math.max(0, (v - scale.min) / (scale.max - scale.min));
      return [cx + Math.cos(angle(i)) * d, cy + Math.sin(angle(i)) * d];
    };
    const ring = (pts: Array<[number, number]>): PathSeg[] => [...pts.map((q, k): PathSeg => ({ op: k ? 'L' : 'M', pts: q })), { op: 'Z', pts: [] }];
    // The web: a ring through the spokes at each major value, and the spokes (the category axis).
    const steps = ticks(scale);
    if (valAx.gridlines) for (const t of steps) ops.push(...stroke(ring(names.map((_, i) => at(i, t))), valAx.gridlines));
    const spoke = (catAx && !catAx.deleted ? catAx.line : null) ?? catAx?.gridlines ?? null;
    if (spoke) for (let i = 0; i < n; i++) ops.push(...stroke(seg(cx, cy, ...at(i, scale.max)), spoke));
    for (const s of p.series) {
      const pts = names.map((_, i) => (s.values[i] === null || s.values[i] === undefined ? null : at(i, s.values[i]!)));
      const shown = pts.filter((q): q is [number, number] => q !== null);
      if (p.radarStyle === 'filled') ops.push(...paint(ring(shown), s.shape));
      else if (s.shape.line && shown.length > 1) {
        const line = s.shape.line;
        const closed = pts.every((q) => q);
        // Points in their own colours colour the stretch of line that runs into them. Points that
        // repeat the series' colour (as varied fills leave them) keep the line unbroken at its corners.
        const own = (i: number) => s.points.get(i)?.line;
        if (pts.some((_, i) => own(i) && own(i)!.color !== line.color))
          pts.forEach((q, i) => {
            const from = i ? pts[i - 1] : closed ? pts[pts.length - 1] : null;
            if (q && from) ops.push(...stroke([{ op: 'M', pts: from }, { op: 'L', pts: q }], own(i) ?? line));
          });
        else ops.push(...stroke(closed ? ring(shown) : shown.map((q, k): PathSeg => ({ op: k ? 'L' : 'M', pts: q })), line));
      }
      pts.forEach((q, i) => {
        if (!q) return;
        const own = s.points.get(i);
        if (s.marker) markerOps.push(...drawMarker(own ? { ...s.marker, shape: { fill: own.fill ?? s.marker.shape.fill, line: own.line ?? s.marker.shape.line } } : s.marker, q[0], q[1]));
        const l = s.labels;
        if (l && !l.hidden?.includes(i)) labelOps.push(...place(measure(labelText(l, s, i, s.values[i]!, undefined, names[i]), l.style, env), l.style.color, q[0] + (s.marker ? s.marker.size / 2 : 0) + 2, q[1], 0, 0.5));
      });
    }
    // The value axis runs up the first spoke, labelled on its left.
    if (!valAx.deleted) {
      if (valAx.line) ops.push(...stroke(seg(cx, cy, cx, cy - r), valAx.line));
      if (valAx.labelPosition !== 'none')
        for (const t of steps) labelOps.push(...place(measure(formatNumber(t, valAx.sourceLinked ? (p.series[0]?.format ?? 'General') : valAx.format), valAx.style, env), valAx.style.color, cx - valAx.style.size * 0.9, at(0, t)[1], 1, 0.5));
    }
    // Category names sit beyond the spokes' ends, turned away from the centre.
    catText.forEach((t, i) => {
      const [ux, uy] = [Math.cos(angle(i)), Math.sin(angle(i))];
      const ax = Math.abs(ux) < 0.05 ? 0.5 : ux > 0 ? 0 : 1;
      const ay = Math.abs(uy) < 0.05 ? 0.5 : uy > 0 ? 0 : 1;
      labelOps.push(...place(t, catAx!.style.color, cx + ux * (r + gap), cy + uy * (r + gap), ax, ay));
    });
  }

  // Axis lines, tick marks and labels.
  for (const a of axes.values()) {
    if (!shown(a)) continue;
    const st = a.axis.style;
    const along = a.horizontal ? [plot.x, plot.x + plot.w] : [plot.y, plot.y + plot.h];
    if (a.axis.line) ops.push(...stroke(a.horizontal ? seg(along[0], a.at, along[1], a.at) : seg(a.at, along[0], a.at, along[1]), a.axis.line));
    // Labels and outside ticks face away from the plot: down and left for bottom and left axes.
    const outward = a.axis.position === 'b' || a.axis.position === 'r' ? 1 : -1;
    const tickPos = a.scale ? ticks(a.scale).map((t) => coord(a, plot, t)) : Array.from({ length: a.count + (a.axis.between ? 1 : 0) }, (_, i) => (a.axis.between ? (a.horizontal ? plot.x + (i / Math.max(1, a.count)) * plot.w : plot.y + plot.h - (i / Math.max(1, a.count)) * plot.h) : coord(a, plot, i)));
    const t0 = a.axis.majorTick === 'in' ? -TICK : a.axis.majorTick === 'cross' ? -TICK : 0;
    const t1 = a.axis.majorTick === 'out' || a.axis.majorTick === 'cross' ? TICK : 0;
    if (a.axis.line && a.axis.majorTick !== 'none')
      for (const p of tickPos) ops.push(...stroke(a.horizontal ? seg(p, a.at + outward * t0, p, a.at + outward * t1) : seg(a.at + outward * t0, p, a.at + outward * t1, p), a.axis.line));
    if (a.axis.labelPosition === 'none') continue;
    // Next to the axis, or at the plot's low or high edge.
    let line = a.at;
    if (a.axis.labelPosition === 'low') line = a.horizontal ? plot.y + plot.h : plot.x;
    else if (a.axis.labelPosition === 'high') line = a.horizontal ? plot.y : plot.x + plot.w;
    const out = a.axis.labelPosition === 'high' ? -outward : outward;
    const gap = gapOf(a);
    const lineH = measure('0', st, env).height;
    const centres = a.scale ? ticks(a.scale).map((t) => coord(a, plot, t)) : a.labels.map((_, i) => coord(a, plot, i));
    a.lines.forEach((lines, i) => {
      if (i % a.skip) return;
      const c = centres[i];
      if (c === undefined) return;
      if (a.horizontal) {
        const y = line + out * gap;
        if (a.rotation) {
          const t = measure(lines.join(' '), st, env);
          // Turned labels end at their category, reading up towards it.
          labelOps.push(...place(t, st.color, c, y, a.rotation < 0 ? 1 : 0, 0.5, a.rotation));
        } else {
          lines.forEach((l, k) => {
            const t = measure(l, st, env);
            const yy = out > 0 ? y + k * lineH : y - (lines.length - k) * lineH;
            labelOps.push(...place(t, st.color, c, yy, 0.5, 0));
          });
        }
      } else {
        const t = measure(lines.join(' '), st, env);
        labelOps.push(...place(t, st.color, line + out * gap, c, out > 0 ? 0 : 1, 0.5, a.rotation));
      }
    });
    if (a.axis.title) {
      const [tw, th, rot] = titleExtent(a.axis.title, env);
      const [w, h] = turned(tw, th, rot);
      const ext = labelExtent(a, env) + gap + AXIS_TITLE_GAP;
      const t = a.axis.title;
      if (t.x !== undefined && t.y !== undefined) ops.push(...drawTitle(t, X + t.x * W + MANUAL_INSET[0] + w / 2, Y + t.y * H + MANUAL_INSET[1] + h / 2, env));
      else if (a.horizontal) ops.push(...drawTitle(t, plot.x + plot.w / 2, line + out * (ext + h / 2), env));
      else ops.push(...drawTitle(t, line + out * (ext + w / 2), plot.y + plot.h / 2, env));
    }
  }
  for (const layer of [markerOps, labelOps, titleOps, legendOps]) for (const op of layer) ops.push(op);
  ops.push({ kind: 'restore' });
  return ops;
}

/** A pie slice (or doughnut segment) from `start` through `sweep` degrees, clockwise from east. */
function slice(cx: number, cy: number, r: number, inner: number, start: number, sweep: number): PathSeg[] {
  const arc = (rad: number, a0: number, a1: number): PathSeg[] => {
    const segs: PathSeg[] = [];
    const n = Math.max(1, Math.ceil(Math.abs(a1 - a0) / 90));
    const step = ((a1 - a0) / n) * (Math.PI / 180);
    const k = (4 / 3) * Math.tan(step / 4);
    for (let i = 0; i < n; i++) {
      const t0 = (a0 * Math.PI) / 180 + i * step;
      const t1 = t0 + step;
      segs.push({ op: 'C', pts: [cx + rad * (Math.cos(t0) - k * Math.sin(t0)), cy + rad * (Math.sin(t0) + k * Math.cos(t0)), cx + rad * (Math.cos(t1) + k * Math.sin(t1)), cy + rad * (Math.sin(t1) - k * Math.cos(t1)), cx + rad * Math.cos(t1), cy + rad * Math.sin(t1)] });
    }
    return segs;
  };
  const rad = (d: number) => (d * Math.PI) / 180;
  const end = start + sweep;
  if (inner > 0)
    return [
      { op: 'M', pts: [cx + r * Math.cos(rad(start)), cy + r * Math.sin(rad(start))] },
      ...arc(r, start, end),
      { op: 'L', pts: [cx + inner * Math.cos(rad(end)), cy + inner * Math.sin(rad(end))] },
      ...arc(inner, end, start),
      { op: 'Z', pts: [] },
    ];
  if (sweep >= 359.999) return [{ op: 'M', pts: [cx + r, cy] }, ...arc(r, 0, 360), { op: 'Z', pts: [] }];
  return [{ op: 'M', pts: [cx, cy] }, { op: 'L', pts: [cx + r * Math.cos(rad(start)), cy + r * Math.sin(rad(start))] }, ...arc(r, start, end), { op: 'Z', pts: [] }];
}
