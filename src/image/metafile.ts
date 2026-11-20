/**
 * Windows metafiles (EMF and WMF) played back as vector drawing commands, so the PDF keeps them as
 * paths, text and embedded bitmaps. Covers the GDI records Word documents use: mapping modes and
 * world transforms, pens and brushes, paths and Bézier curves, shapes, clipping, DIB blits and
 * text. EMF+ records are skipped; dual EMF+ files carry the same picture as plain EMF records.
 */

export interface MetaSeg {
  op: 'M' | 'L' | 'C' | 'Z';
  pts: number[];
}

export interface MetaFont {
  family: string;
  /** Em size in page points, or the cell height (ascent plus descent) when `cell` is set. */
  size: number;
  cell: boolean;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
}

interface Dib {
  header: Uint8Array;
  bits: Uint8Array;
}

export type MetaOp =
  | { kind: 'path'; segs: MetaSeg[]; fill?: string; stroke?: string; width: number; evenOdd: boolean; clip: MetaSeg[][] }
  | { kind: 'image'; x: number; y: number; w: number; h: number; dib: Dib; crop?: { l: number; t: number; r: number; b: number }; clip: MetaSeg[][] }
  | {
      kind: 'text';
      x: number;
      y: number;
      text: string;
      font: MetaFont;
      color: string;
      /** 0 left, 1 centre, 2 right of the reference point; vertical: 'top' | 'baseline' | 'bottom'. */
      align: number;
      valign: 'top' | 'baseline' | 'bottom';
      /** Character advances in page points, when the record lists them. */
      dx?: number[];
      /** Baseline angle in degrees, counter-clockwise. */
      angle: number;
      clip: MetaSeg[][];
    };

export interface MetaTarget {
  x: number;
  y: number;
  w: number;
  h: number;
}

type Mat = [number, number, number, number, number, number];
const IDENTITY: Mat = [1, 0, 0, 1, 0, 0];

interface Pen {
  kind: 'pen';
  style: number;
  width: number;
  color: string;
}
interface Brush {
  kind: 'brush';
  style: number;
  color: string;
}
interface Font {
  kind: 'font';
  height: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  escapement: number;
  face: string;
}
type GdiObject = Pen | Brush | Font | { kind: 'other' };

interface DcState {
  pen: Pen;
  brush: Brush;
  font: Font;
  textColor: string;
  bkColor: string;
  bkMode: number;
  fillMode: number;
  textAlign: number;
  mapMode: number;
  wOrg: [number, number];
  wExt: [number, number];
  vOrg: [number, number];
  vExt: [number, number];
  world: Mat;
  clip: MetaSeg[][];
}

const BLACK_PEN: Pen = { kind: 'pen', style: 0, width: 0, color: '000000' };
const NULL_PEN: Pen = { kind: 'pen', style: 5, width: 0, color: '000000' };
const WHITE_BRUSH: Brush = { kind: 'brush', style: 0, color: 'FFFFFF' };
const NULL_BRUSH: Brush = { kind: 'brush', style: 1, color: 'FFFFFF' };
const DEFAULT_FONT: Font = { kind: 'font', height: -12, weight: 400, italic: false, underline: false, strike: false, escapement: 0, face: 'Arial' };

function stockObject(index: number): GdiObject | undefined {
  switch (index & 0x7fffffff) {
    case 0:
      return WHITE_BRUSH;
    case 1:
      return { kind: 'brush', style: 0, color: 'C0C0C0' };
    case 2:
      return { kind: 'brush', style: 0, color: '808080' };
    case 3:
      return { kind: 'brush', style: 0, color: '404040' };
    case 4:
      return { kind: 'brush', style: 0, color: '000000' };
    case 5:
      return NULL_BRUSH;
    case 6:
      return { kind: 'pen', style: 0, width: 0, color: 'FFFFFF' };
    case 7:
      return BLACK_PEN;
    case 8:
      return NULL_PEN;
    case 10:
    case 11:
    case 12:
    case 13:
    case 14:
    case 16:
    case 17:
      return DEFAULT_FONT;
    default:
      return undefined;
  }
}

function colorRef(d: DataView, o: number): string {
  const hex = (n: number) => n.toString(16).padStart(2, '0').toUpperCase();
  return hex(d.getUint8(o)) + hex(d.getUint8(o + 1)) + hex(d.getUint8(o + 2));
}

function mul(a: Mat, b: Mat): Mat {
  // Row-vector convention (GDI): p' = p * a * b.
  return [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
}

/** Bézier control points for an elliptical arc from angle a0 to a1 (radians, y down). */
function arcSegs(cx: number, cy: number, rx: number, ry: number, a0: number, a1: number, move: boolean): MetaSeg[] {
  const segs: MetaSeg[] = [];
  const n = Math.max(1, Math.ceil(Math.abs(a1 - a0) / (Math.PI / 2)));
  const step = (a1 - a0) / n;
  const k = (4 / 3) * Math.tan(step / 4);
  let a = a0;
  if (move) segs.push({ op: 'M', pts: [cx + rx * Math.cos(a), cy + ry * Math.sin(a)] });
  for (let i = 0; i < n; i++) {
    const b = a + step;
    const [c0, s0, c1, s1] = [Math.cos(a), Math.sin(a), Math.cos(b), Math.sin(b)];
    segs.push({ op: 'C', pts: [cx + rx * (c0 - k * s0), cy + ry * (s0 + k * c0), cx + rx * (c1 + k * s1), cy + ry * (s1 - k * c1), cx + rx * c1, cy + ry * s1] });
    a = b;
  }
  return segs;
}

// Preserve views into the record. The raster decoder checks its pixel budget before allocating.
function dibParts(bytes: Uint8Array, recordEnd: number, bmiOff: number, bmiLen: number, bitsOff: number, bitsLen: number): Dib | null {
  if (bmiLen < 12 || bitsLen <= 0 || bmiOff + bmiLen > recordEnd || bitsOff + bitsLen > recordEnd) return null;
  return { header: bytes.subarray(bmiOff, bmiOff + bmiLen), bits: bytes.subarray(bitsOff, bitsOff + bitsLen) };
}

/** Length of a packed DIB's header and colour table (WMF bitmaps carry the bits right after). */
function dibHeaderLength(d: DataView, o: number): number {
  const size = d.getUint32(o, true);
  if (size === 12) {
    const bits = d.getUint16(o + 10, true);
    return 12 + (bits <= 8 ? (1 << bits) * 3 : 0);
  }
  const bits = d.getUint16(o + 14, true);
  const compression = d.getUint32(o + 16, true);
  const used = d.getUint32(o + 32, true);
  const colors = used || (bits <= 8 ? 1 << bits : 0);
  return size + (compression === 3 && size === 40 ? 12 : 0) + colors * 4;
}

class Player {
  readonly ops: MetaOp[] = [];
  dc: DcState;
  private stack: DcState[] = [];
  path: MetaSeg[] | null = null;
  /** A path closed by ENDPATH, waiting for a fill, stroke or clip record. */
  pendingPath: MetaSeg[] | null = null;
  /** WMF: logical units map straight from the window onto the target. */
  wmf = false;
  cur: [number, number] = [0, 0];
  /** Device rectangle mapped onto the target. */
  frame = { x: 0, y: 0, w: 1, h: 1 };
  /** Device units per millimetre (EMF reference device). */
  dpmm: [number, number] = [96 / 25.4, 96 / 25.4];

  constructor(private target: MetaTarget) {
    this.dc = {
      pen: BLACK_PEN,
      brush: WHITE_BRUSH,
      font: DEFAULT_FONT,
      textColor: '000000',
      bkColor: 'FFFFFF',
      bkMode: 2,
      fillMode: 1,
      textAlign: 0,
      mapMode: 1,
      wOrg: [0, 0],
      wExt: [1, 1],
      vOrg: [0, 0],
      vExt: [1, 1],
      world: IDENTITY,
      clip: [],
    };
  }

  save(): void {
    if (this.stack.length >= 1024) throw new Error('Metafile saved state limit exceeded');
    this.stack.push({ ...this.dc });
  }
  restore(n: number): void {
    // A positive argument names the save (1 = first), a negative one counts back from the latest.
    const count = n < 0 ? -n : this.stack.length - n + 1;
    for (let i = 0; i < count && this.stack.length; i++) this.dc = this.stack.pop()!;
  }

  /** Logical units to device units, before the world transform. */
  private pageScale(): [number, number] {
    const { mapMode: m, wExt, vExt } = this.dc;
    const [dx, dy] = this.dpmm;
    switch (m) {
      case 2:
        return [dx * 0.1, -dy * 0.1];
      case 3:
        return [dx * 0.01, -dy * 0.01];
      case 4:
        return [dx * 0.254, -dy * 0.254];
      case 5:
        return [dx * 0.0254, -dy * 0.0254];
      case 6:
        return [(dx * 25.4) / 1440, (-dy * 25.4) / 1440];
      case 7: {
        const sx = vExt[0] / (wExt[0] || 1);
        const sy = vExt[1] / (wExt[1] || 1);
        const s = Math.min(Math.abs(sx), Math.abs(sy));
        return [Math.sign(sx) * s, Math.sign(sy) * s];
      }
      case 8:
        return [vExt[0] / (wExt[0] || 1), vExt[1] / (wExt[1] || 1)];
      default:
        return [1, 1];
    }
  }

  /** Logical point to page coordinates. */
  pt(x: number, y: number): [number, number] {
    if (this.wmf) {
      const { wOrg, wExt } = this.dc;
      return [this.target.x + ((x - wOrg[0]) * this.target.w) / (wExt[0] || 1), this.target.y + ((y - wOrg[1]) * this.target.h) / (wExt[1] || 1)];
    }
    const w = this.dc.world;
    const px = x * w[0] + y * w[2] + w[4];
    const py = x * w[1] + y * w[3] + w[5];
    const [sx, sy] = this.pageScale();
    const dx = (px - this.dc.wOrg[0]) * sx + this.dc.vOrg[0];
    const dy = (py - this.dc.wOrg[1]) * sy + this.dc.vOrg[1];
    const f = this.frame;
    return [this.target.x + ((dx - f.x) * this.target.w) / f.w, this.target.y + ((dy - f.y) * this.target.h) / f.h];
  }

  /** Average page length of a logical length (pen widths, font heights). */
  len(v: number): number {
    const a = this.pt(0, 0);
    const b = this.pt(v, 0);
    const c = this.pt(0, v);
    return (Math.hypot(b[0] - a[0], b[1] - a[1]) + Math.hypot(c[0] - a[0], c[1] - a[1])) / 2;
  }

  /** Page size of one device pixel. */
  private pixel(): number {
    if (this.wmf) return Math.min(this.len(1), 0.75);
    return Math.abs(this.target.w / this.frame.w + this.target.h / this.frame.h) / 2;
  }

  penWidth(): number {
    const p = this.dc.pen;
    const w = p.width > 0 ? this.len(p.width) : 0;
    return Math.max(w, this.pixel());
  }

  draw(segs: MetaSeg[], fill: boolean, stroke: boolean): void {
    if (!segs.length) return;
    if (this.path) {
      this.path.push(...segs);
      return;
    }
    const brush = this.dc.brush;
    const pen = this.dc.pen;
    const f = fill && brush.style !== 1 ? brush.color : undefined;
    const s = stroke && (pen.style & 15) !== 5 ? pen.color : undefined;
    if (!f && !s) return;
    this.ops.push({ kind: 'path', segs, fill: f, stroke: s, width: this.penWidth(), evenOdd: this.dc.fillMode !== 2, clip: this.dc.clip });
  }

  line(from: [number, number], to: [number, number]): void {
    if (this.path) {
      if (!this.path.length) this.path.push({ op: 'M', pts: from });
      this.path.push({ op: 'L', pts: to });
      return;
    }
    this.draw([{ op: 'M', pts: from }, { op: 'L', pts: to }], false, true);
  }

  poly(points: [number, number][], closed: boolean, continueFromCurrent = false): void {
    if (!points.length) return;
    const segs: MetaSeg[] = [];
    if (continueFromCurrent) {
      if (!this.path || !this.path.length) segs.push({ op: 'M', pts: this.pt(...this.cur) });
    } else segs.push({ op: 'M', pts: this.pt(...points[0]) });
    for (const p of continueFromCurrent ? points : points.slice(1)) segs.push({ op: 'L', pts: this.pt(...p) });
    if (closed) segs.push({ op: 'Z', pts: [] });
    if (this.path && continueFromCurrent) this.path.push(...segs);
    else this.draw(segs, closed, true);
  }

  bezier(points: [number, number][], continueFromCurrent: boolean): void {
    const segs: MetaSeg[] = [];
    let i = 0;
    if (continueFromCurrent) {
      if (!this.path || !this.path.length) segs.push({ op: 'M', pts: this.pt(...this.cur) });
    } else if (points.length) {
      segs.push({ op: 'M', pts: this.pt(...points[0]) });
      i = 1;
    }
    for (; i + 2 < points.length; i += 3) segs.push({ op: 'C', pts: [...this.pt(...points[i]), ...this.pt(...points[i + 1]), ...this.pt(...points[i + 2])] });
    if (this.path && continueFromCurrent) this.path.push(...segs);
    else this.draw(segs, false, true);
  }

  rect(l: number, t: number, r: number, b: number, fill = true, stroke = true): void {
    const segs: MetaSeg[] = [{ op: 'M', pts: this.pt(l, t) }, { op: 'L', pts: this.pt(r, t) }, { op: 'L', pts: this.pt(r, b) }, { op: 'L', pts: this.pt(l, b) }, { op: 'Z', pts: [] }];
    this.draw(segs, fill, stroke);
  }

  ellipseSegs(l: number, t: number, r: number, b: number): MetaSeg[] {
    // Built in logical space, then mapped point by point.
    const raw = arcSegs((l + r) / 2, (t + b) / 2, Math.abs(r - l) / 2, Math.abs(b - t) / 2, 0, Math.PI * 2, true);
    const segs = raw.map((s) => ({ op: s.op, pts: this.mapPts(s.pts) }));
    segs.push({ op: 'Z', pts: [] });
    return segs;
  }

  roundRectSegs(l: number, t: number, r: number, b: number, cw: number, ch: number): MetaSeg[] {
    const rx = Math.min(Math.abs(cw) / 2, Math.abs(r - l) / 2);
    const ry = Math.min(Math.abs(ch) / 2, Math.abs(b - t) / 2);
    const [x0, x1, y0, y1] = [Math.min(l, r), Math.max(l, r), Math.min(t, b), Math.max(t, b)];
    const raw: MetaSeg[] = [
      { op: 'M', pts: [x0 + rx, y0] },
      { op: 'L', pts: [x1 - rx, y0] },
      ...arcSegs(x1 - rx, y0 + ry, rx, ry, -Math.PI / 2, 0, false),
      { op: 'L', pts: [x1, y1 - ry] },
      ...arcSegs(x1 - rx, y1 - ry, rx, ry, 0, Math.PI / 2, false),
      { op: 'L', pts: [x0 + rx, y1] },
      ...arcSegs(x0 + rx, y1 - ry, rx, ry, Math.PI / 2, Math.PI, false),
      { op: 'L', pts: [x0, y0 + ry] },
      ...arcSegs(x0 + rx, y0 + ry, rx, ry, Math.PI, Math.PI * 1.5, false),
    ];
    const segs = raw.map((s) => ({ op: s.op, pts: this.mapPts(s.pts) }));
    segs.push({ op: 'Z', pts: [] });
    return segs;
  }

  /** GDI arcs run counter-clockwise (as seen on screen) from the start radial to the end radial. */
  arcSegs(l: number, t: number, r: number, b: number, xs: number, ys: number, xe: number, ye: number, kind: 'arc' | 'chord' | 'pie', move: boolean): MetaSeg[] {
    const cx = (l + r) / 2;
    const cy = (t + b) / 2;
    const rx = Math.abs(r - l) / 2;
    const ry = Math.abs(b - t) / 2;
    const a0 = Math.atan2((ys - cy) / (ry || 1), (xs - cx) / (rx || 1));
    let a1 = Math.atan2((ye - cy) / (ry || 1), (xe - cx) / (rx || 1));
    // Counter-clockwise on screen is decreasing angle in y-down coordinates.
    while (a1 >= a0) a1 -= Math.PI * 2;
    const raw = arcSegs(cx, cy, rx, ry, a0, a1, true);
    const segs = raw.map((s) => ({ op: s.op, pts: this.mapPts(s.pts) }));
    if (!move) segs[0] = { op: 'L', pts: segs[0].pts };
    if (kind === 'pie') segs.push({ op: 'L', pts: this.pt(cx, cy) }, { op: 'Z', pts: [] });
    if (kind === 'chord') segs.push({ op: 'Z', pts: [] });
    return segs;
  }

  mapPts(pts: number[]): number[] {
    const out: number[] = [];
    for (let i = 0; i + 1 < pts.length; i += 2) out.push(...this.pt(pts[i], pts[i + 1]));
    return out;
  }

  /** Clip to a polygon list in page coordinates (intersecting with the current clip). */
  clipTo(segs: MetaSeg[], replace: boolean): void {
    if (!replace && this.dc.clip.length >= 128) throw new Error('Metafile clip depth limit exceeded');
    this.dc.clip = replace ? [segs] : [...this.dc.clip, segs];
  }

  image(dib: Dib | null, l: number, t: number, w: number, h: number, crop?: { l: number; t: number; r: number; b: number }): void {
    if (!dib) return;
    const a = this.pt(l, t);
    const b = this.pt(l + w, t + h);
    const x = Math.min(a[0], b[0]);
    const y = Math.min(a[1], b[1]);
    const W = Math.abs(b[0] - a[0]);
    const H = Math.abs(b[1] - a[1]);
    if (W <= 0 || H <= 0) return;
    this.ops.push({ kind: 'image', x, y, w: W, h: H, dib, crop, clip: this.dc.clip });
  }

  text(x: number, y: number, s: string, dx: number[] | undefined): void {
    if (!s) return;
    const f = this.dc.font;
    const align = this.dc.textAlign;
    if (align & 1) [x, y] = this.cur;
    const p = this.pt(x, y);
    const size = this.len(Math.abs(f.height) || 12);
    const h = (align & 6) === 6 ? 1 : align & 2 ? 2 : 0;
    const v = (align & 24) === 24 ? 'baseline' : align & 8 ? 'bottom' : 'top';
    // Escapement is counter-clockwise in device space; a y-up mapping mirrors it.
    const flip = this.wmf ? this.dc.wExt[1] < 0 : this.pageScale()[1] < 0 !== this.target.h / this.frame.h < 0;
    const angle = (f.escapement / 10) * (flip ? -1 : 1);
    const advances = dx?.map((d) => this.len(d));
    this.ops.push({
      kind: 'text',
      x: p[0],
      y: p[1],
      text: s,
      font: { family: f.face, size, cell: f.height > 0, bold: f.weight >= 600, italic: f.italic, underline: f.underline, strike: f.strike },
      color: this.dc.textColor,
      align: h,
      valign: v,
      dx: advances,
      angle,
      clip: this.dc.clip,
    });
  }

  select(obj: GdiObject | undefined): void {
    if (!obj) return;
    if (obj.kind === 'pen') this.dc.pen = obj;
    else if (obj.kind === 'brush') this.dc.brush = obj;
    else if (obj.kind === 'font') this.dc.font = obj;
  }
}

function readFace(bytes: Uint8Array, o: number, wide: boolean): string {
  let s = '';
  for (let i = 0; i < 32; i++) {
    const c = wide ? bytes[o + i * 2] | (bytes[o + i * 2 + 1] << 8) : bytes[o + i];
    if (!c) break;
    s += String.fromCharCode(c);
  }
  return s.trim() || 'Arial';
}

/** True when the data is an EMF or WMF file. */
export function isMetafile(d: Uint8Array): boolean {
  if (d.length < 44) return false;
  if (d[0] === 1 && d[1] === 0 && d[2] === 0 && d[3] === 0 && d[40] === 0x20 && d[41] === 0x45 && d[42] === 0x4d && d[43] === 0x46) return true;
  return (d[0] === 0xd7 && d[1] === 0xcd && d[2] === 0xc6 && d[3] === 0x9a) || (d[0] === 1 && d[1] === 0 && d[2] === 9 && d[3] === 0) || (d[0] === 2 && d[1] === 0 && d[2] === 9 && d[3] === 0);
}

/** Play a metafile into the target rectangle; null when the data cannot be read. */
export function playMetafile(bytes: Uint8Array, target: MetaTarget): MetaOp[] | null {
  try {
    if (bytes[0] === 1 && bytes[40] === 0x20 && bytes[41] === 0x45) return playEmf(bytes, target);
    return playWmf(bytes, target);
  } catch {
    return null;
  }
}

/** Fonts and text a metafile draws, so they can be loaded before layout. */
export function metafileText(bytes: Uint8Array): Array<{ family: string; bold: boolean; italic: boolean; text: string }> {
  const ops = playMetafile(bytes, { x: 0, y: 0, w: 100, h: 100 }) ?? [];
  const out = new Map<string, { family: string; bold: boolean; italic: boolean; text: string }>();
  for (const op of ops) {
    if (op.kind !== 'text') continue;
    const key = `${op.font.family}|${op.font.bold}|${op.font.italic}`;
    const e = out.get(key) ?? { family: op.font.family, bold: op.font.bold, italic: op.font.italic, text: '' };
    e.text += op.text;
    out.set(key, e);
  }
  return [...out.values()];
}

// ---------------------------------------------------------------------------------------------
// EMF

function playEmf(bytes: Uint8Array, target: MetaTarget): MetaOp[] | null {
  const d = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const p = new Player(target);
  const i32 = (o: number) => d.getInt32(o, true);
  const u32 = (o: number) => d.getUint32(o, true);
  const f32 = (o: number) => d.getFloat32(o, true);
  // A file whose header record is broken cannot be played; it is reported, not drawn as nothing.
  if (bytes.length < 88 || u32(4) < 88 || u32(4) > bytes.length) return null;
  // Reference device and picture frame.
  const bounds = [i32(8), i32(12), i32(16), i32(20)];
  const frame = [i32(24), i32(28), i32(32), i32(36)];
  const dev = [i32(72), i32(76)];
  const mm = [i32(80), i32(84)];
  if (dev[0] > 0 && mm[0] > 0 && dev[1] > 0 && mm[1] > 0) p.dpmm = [dev[0] / mm[0], dev[1] / mm[1]];
  const fw = ((frame[2] - frame[0]) / 100) * p.dpmm[0];
  const fh = ((frame[3] - frame[1]) / 100) * p.dpmm[1];
  if (fw > 0 && fh > 0) p.frame = { x: (frame[0] / 100) * p.dpmm[0], y: (frame[1] / 100) * p.dpmm[1], w: fw, h: fh };
  else p.frame = { x: bounds[0], y: bounds[1], w: Math.max(1, bounds[2] - bounds[0] + 1), h: Math.max(1, bounds[3] - bounds[1] + 1) };

  const points = (o: number, n: number, small: boolean): [number, number][] => {
    const out: [number, number][] = [];
    for (let i = 0; i < n; i++) out.push(small ? [d.getInt16(o + i * 4, true), d.getInt16(o + i * 4 + 2, true)] : [i32(o + i * 8), i32(o + i * 8 + 4)]);
    return out;
  };
  const objects = new Map<number, GdiObject>();
  const obj = (index: number) => (index & 0x80000000 ? stockObject(index) : objects.get(index));

  let off = 0;
  let records = 0;
  while (off + 8 <= bytes.length && records++ < 2_000_000) {
    const type = u32(off);
    const size = u32(off + 4);
    if (size < 8 || off + size > bytes.length) break;
    const r = off;
    off += size;
    switch (type) {
      case 14: // EOF
        return p.ops;
      case 2: // POLYBEZIER
      case 85:
        p.bezier(points(r + 28, u32(r + 24), type === 85), false);
        break;
      case 3: // POLYGON
      case 86:
        p.poly(points(r + 28, u32(r + 24), type === 86), true);
        break;
      case 4: // POLYLINE
      case 87:
        p.poly(points(r + 28, u32(r + 24), type === 87), false);
        break;
      case 5: // POLYBEZIERTO
      case 88: {
        const pts = points(r + 28, u32(r + 24), type === 88);
        p.bezier(pts, true);
        if (pts.length) p.cur = pts[pts.length - 1];
        break;
      }
      case 6: // POLYLINETO
      case 89: {
        const pts = points(r + 28, u32(r + 24), type === 89);
        p.poly(pts, false, true);
        if (pts.length) p.cur = pts[pts.length - 1];
        break;
      }
      case 7: // POLYPOLYLINE
      case 8: // POLYPOLYGON
      case 90:
      case 91: {
        const small = type >= 90;
        const closed = type === 8 || type === 91;
        const n = u32(r + 24);
        let at = r + 32 + n * 4;
        const segs: MetaSeg[] = [];
        for (let i = 0; i < n; i++) {
          const count = u32(r + 32 + i * 4);
          const pts = points(at, count, small);
          at += count * (small ? 4 : 8);
          if (!pts.length) continue;
          segs.push({ op: 'M', pts: p.pt(...pts[0]) });
          for (const q of pts.slice(1)) segs.push({ op: 'L', pts: p.pt(...q) });
          if (closed) segs.push({ op: 'Z', pts: [] });
        }
        p.draw(segs, closed, true);
        break;
      }
      case 56: // POLYDRAW
      case 92: {
        const small = type === 92;
        const n = u32(r + 24);
        const pts = points(r + 28, n, small);
        const types = r + 28 + n * (small ? 4 : 8);
        const segs: MetaSeg[] = [];
        let start = p.pt(...p.cur);
        if (!p.path?.length) segs.push({ op: 'M', pts: start });
        for (let i = 0; i < n; i++) {
          const t = bytes[types + i];
          if ((t & 6) === 6) {
            start = p.pt(...pts[i]);
            segs.push({ op: 'M', pts: start });
          } else if ((t & 6) === 2) segs.push({ op: 'L', pts: p.pt(...pts[i]) });
          else if ((t & 6) === 4 && i + 2 < n) {
            segs.push({ op: 'C', pts: [...p.pt(...pts[i]), ...p.pt(...pts[i + 1]), ...p.pt(...pts[i + 2])] });
            i += 2;
          }
          if (t & 1) segs.push({ op: 'Z', pts: [] });
          p.cur = pts[i];
        }
        if (p.path) p.path.push(...segs);
        else p.draw(segs, false, true);
        break;
      }
      case 9: // SETWINDOWEXTEX
        p.dc.wExt = [i32(r + 8), i32(r + 12)];
        break;
      case 10: // SETWINDOWORGEX
        p.dc.wOrg = [i32(r + 8), i32(r + 12)];
        break;
      case 11: // SETVIEWPORTEXTEX
        p.dc.vExt = [i32(r + 8), i32(r + 12)];
        break;
      case 12: // SETVIEWPORTORGEX
        p.dc.vOrg = [i32(r + 8), i32(r + 12)];
        break;
      case 17: // SETMAPMODE
        p.dc.mapMode = u32(r + 8);
        break;
      case 18: // SETBKMODE
        p.dc.bkMode = u32(r + 8);
        break;
      case 19: // SETPOLYFILLMODE
        p.dc.fillMode = u32(r + 8);
        break;
      case 22: // SETTEXTALIGN
        p.dc.textAlign = u32(r + 8);
        break;
      case 24: // SETTEXTCOLOR
        p.dc.textColor = colorRef(d, r + 8);
        break;
      case 25: // SETBKCOLOR
        p.dc.bkColor = colorRef(d, r + 8);
        break;
      case 27: // MOVETOEX
        p.cur = [i32(r + 8), i32(r + 12)];
        if (p.path) p.path.push({ op: 'M', pts: p.pt(...p.cur) });
        break;
      case 54: {
        // LINETO
        const to: [number, number] = [i32(r + 8), i32(r + 12)];
        p.line(p.pt(...p.cur), p.pt(...to));
        p.cur = to;
        break;
      }
      case 30: {
        // INTERSECTCLIPRECT
        const [l, t, rr, b] = [i32(r + 8), i32(r + 12), i32(r + 16), i32(r + 20)];
        p.clipTo([{ op: 'M', pts: p.pt(l, t) }, { op: 'L', pts: p.pt(rr, t) }, { op: 'L', pts: p.pt(rr, b) }, { op: 'L', pts: p.pt(l, b) }, { op: 'Z', pts: [] }], false);
        break;
      }
      case 75: {
        // EXTSELECTCLIPRGN: region rectangles are in device units.
        const cb = u32(r + 8);
        const mode = u32(r + 12);
        if (mode === 5 && cb === 0) {
          p.dc.clip = [];
          break;
        }
        if (cb < 32 || (mode !== 1 && mode !== 5)) break;
        const count = u32(r + 16 + 8);
        const segs: MetaSeg[] = [];
        const f = p.frame;
        const map = (x: number, y: number): number[] => [target.x + ((x - f.x) * target.w) / f.w, target.y + ((y - f.y) * target.h) / f.h];
        for (let i = 0; i < count && r + 48 + i * 16 + 16 <= r + size; i++) {
          const q = r + 48 + i * 16;
          const [l, t, rr, b] = [i32(q), i32(q + 4), i32(q + 8), i32(q + 12)];
          segs.push({ op: 'M', pts: map(l, t) }, { op: 'L', pts: map(rr, t) }, { op: 'L', pts: map(rr, b) }, { op: 'L', pts: map(l, b) }, { op: 'Z', pts: [] });
        }
        if (segs.length) p.clipTo(segs, mode === 5);
        break;
      }
      case 33: // SAVEDC
        p.save();
        break;
      case 34: // RESTOREDC
        p.restore(i32(r + 8));
        break;
      case 35: // SETWORLDTRANSFORM
        p.dc.world = [f32(r + 8), f32(r + 12), f32(r + 16), f32(r + 20), f32(r + 24), f32(r + 28)];
        break;
      case 36: {
        // MODIFYWORLDTRANSFORM
        const x: Mat = [f32(r + 8), f32(r + 12), f32(r + 16), f32(r + 20), f32(r + 24), f32(r + 28)];
        const mode = u32(r + 32);
        if (mode === 1) p.dc.world = IDENTITY;
        else if (mode === 2) p.dc.world = mul(x, p.dc.world);
        else if (mode === 3) p.dc.world = mul(p.dc.world, x);
        else if (mode === 4) p.dc.world = x;
        break;
      }
      case 37: // SELECTOBJECT
        p.select(obj(u32(r + 8)));
        break;
      case 38: // CREATEPEN
        objects.set(u32(r + 8), { kind: 'pen', style: u32(r + 12), width: i32(r + 16), color: colorRef(d, r + 24) });
        break;
      case 95: // EXTCREATEPEN
        objects.set(u32(r + 8), { kind: 'pen', style: u32(r + 36) === 1 ? 5 : u32(r + 28), width: u32(r + 32), color: u32(r + 36) === 1 ? '000000' : colorRef(d, r + 40) });
        break;
      case 39: // CREATEBRUSHINDIRECT
        objects.set(u32(r + 8), { kind: 'brush', style: u32(r + 12), color: colorRef(d, r + 16) });
        break;
      case 93: // CREATEMONOBRUSH
      case 94: // CREATEDIBPATTERNBRUSHPT
        objects.set(u32(r + 8), { kind: 'brush', style: 0, color: '808080' });
        break;
      case 40: // DELETEOBJECT
        objects.delete(u32(r + 8));
        break;
      case 82: {
        // EXTCREATEFONTINDIRECTW
        const o = r + 12;
        objects.set(u32(r + 8), {
          kind: 'font',
          height: i32(o),
          escapement: i32(o + 8),
          weight: i32(o + 16),
          italic: !!bytes[o + 20],
          underline: !!bytes[o + 21],
          strike: !!bytes[o + 22],
          face: readFace(bytes, o + 28, true),
        });
        break;
      }
      case 42: {
        // ELLIPSE
        p.draw(p.ellipseSegs(i32(r + 8), i32(r + 12), i32(r + 16), i32(r + 20)), true, true);
        break;
      }
      case 43: // RECTANGLE
        p.rect(i32(r + 8), i32(r + 12), i32(r + 16), i32(r + 20));
        break;
      case 44: // ROUNDRECT
        p.draw(p.roundRectSegs(i32(r + 8), i32(r + 12), i32(r + 16), i32(r + 20), i32(r + 24), i32(r + 28)), true, true);
        break;
      case 45: // ARC
      case 46: // CHORD
      case 47: // PIE
      case 55: {
        // ARCTO
        const kind = type === 46 ? 'chord' : type === 47 ? 'pie' : 'arc';
        const segs = p.arcSegs(i32(r + 8), i32(r + 12), i32(r + 16), i32(r + 20), i32(r + 24), i32(r + 28), i32(r + 32), i32(r + 36), kind, type !== 55);
        if (type === 55) {
          const start = p.pt(...p.cur);
          if (p.path) {
            if (!p.path.length) p.path.push({ op: 'M', pts: start });
            p.path.push(...segs);
          } else p.draw([{ op: 'M', pts: start }, ...segs], false, true);
        } else p.draw(segs, kind !== 'arc', true);
        break;
      }
      case 59: // BEGINPATH
        p.path = [];
        break;
      case 60: // ENDPATH
        p.pendingPath = p.path;
        p.path = null;
        break;
      case 61: // CLOSEFIGURE
        p.path?.push({ op: 'Z', pts: [] });
        break;
      case 62: // FILLPATH
      case 63: // STROKEANDFILLPATH
      case 64: {
        // STROKEPATH
        const segs = p.pendingPath ?? p.path ?? [];
        p.pendingPath = null;
        p.path = null;
        p.draw(segs, type !== 64, type !== 62);
        break;
      }
      case 67: {
        // SELECTCLIPPATH
        const segs = p.pendingPath ?? [];
        p.pendingPath = null;
        const mode = u32(r + 8);
        if (segs.length && (mode === 1 || mode === 5)) p.clipTo(segs, mode === 5);
        break;
      }
      case 68: // ABORTPATH
        p.path = null;
        p.pendingPath = null;
        break;
      case 76: // BITBLT
      case 77: {
        // STRETCHBLT
        const [x, y, w, h] = [i32(r + 24), i32(r + 28), i32(r + 32), i32(r + 36)];
        const rop = u32(r + 40);
        const cbBmi = u32(r + 88);
        if (!cbBmi) {
          // A pattern fill of the destination with the brush, or plain black or white.
          const brush = p.dc.brush;
          const color = rop === 0x42 ? '000000' : rop === 0xff0062 ? 'FFFFFF' : rop === 0xf00021 && brush.style !== 1 ? brush.color : null;
          if (color) {
            const saved = p.dc.brush;
            p.dc.brush = { kind: 'brush', style: 0, color };
            const pen = p.dc.pen;
            p.dc.pen = NULL_PEN;
            p.rect(x, y, x + w, y + h);
            p.dc.pen = pen;
            p.dc.brush = saved;
          }
          break;
        }
        if (!copiesSource(rop)) break;
        const bmp = dibParts(bytes, off, r + u32(r + 84), cbBmi, r + u32(r + 92), u32(r + 96));
        const sx = i32(r + 44);
        const sy = i32(r + 48);
        const sw = type === 77 ? i32(r + 100) : w;
        const sh = type === 77 ? i32(r + 104) : h;
        p.image(bmp, x, y, w, h, sourceCrop(d, r + u32(r + 84), sx, sy, sw, sh));
        break;
      }
      case 80: {
        // SETDIBITSTODEVICE
        const bmp = dibParts(bytes, off, r + u32(r + 48), u32(r + 52), r + u32(r + 56), u32(r + 60));
        p.image(bmp, i32(r + 24), i32(r + 28), i32(r + 40), i32(r + 44), sourceCrop(d, r + u32(r + 48), i32(r + 32), i32(r + 36), i32(r + 40), i32(r + 44)));
        break;
      }
      case 81: {
        // STRETCHDIBITS
        if (!copiesSource(u32(r + 68))) break;
        const bmp = dibParts(bytes, off, r + u32(r + 48), u32(r + 52), r + u32(r + 56), u32(r + 60));
        p.image(bmp, i32(r + 24), i32(r + 28), i32(r + 72), i32(r + 76), sourceCrop(d, r + u32(r + 48), i32(r + 32), i32(r + 36), i32(r + 40), i32(r + 44)));
        break;
      }
      case 83: // EXTTEXTOUTA
      case 84: {
        // EXTTEXTOUTW
        const e = r + 36;
        const n = u32(e + 8);
        const so = r + u32(e + 12);
        const options = u32(e + 16);
        const dxo = u32(e + 36);
        if (n > 100000 || so + n * (type === 84 ? 2 : 1) > r + size) break;
        let s = '';
        for (let i = 0; i < n; i++) s += String.fromCharCode(type === 84 ? d.getUint16(so + i * 2, true) : bytes[so + i]);
        let dx: number[] | undefined;
        if (dxo && r + dxo + n * 4 <= r + size) {
          dx = [];
          const step = options & 0x2000 ? 8 : 4;
          for (let i = 0; i < n; i++) dx.push(i32(r + dxo + i * step));
        }
        p.text(i32(e), i32(e + 4), s, dx);
        break;
      }
      default:
        break;
    }
  }
  return p.ops;
}

/** Raster operations that put the source bitmap's own pixels on the page. */
function copiesSource(rop: number): boolean {
  return rop === 0xcc0020 || rop === 0xc000ca || rop === 0xee0086 || rop === 0;
}

/** Fractions cut from each side of a DIB to show only its source rectangle. */
function sourceCrop(d: DataView, bmi: number, sx: number, sy: number, sw: number, sh: number): { l: number; t: number; r: number; b: number } | undefined {
  if (bmi + 12 > d.byteLength) return undefined;
  const core = d.getUint32(bmi, true) === 12;
  const W = core ? d.getUint16(bmi + 4, true) : d.getInt32(bmi + 4, true);
  const rawH = core ? d.getInt16(bmi + 6, true) : d.getInt32(bmi + 8, true);
  const H = Math.abs(rawH);
  if (W <= 0 || H <= 0 || sw <= 0 || sh <= 0) return undefined;
  if (sx <= 0 && sy <= 0 && sw >= W && sh >= H) return undefined;
  const l = Math.max(0, sx / W);
  const r = Math.max(0, 1 - (sx + sw) / W);
  // Source rows count from the bottom of a bottom-up bitmap.
  const top = rawH > 0 ? H - (sy + sh) : sy;
  const t = Math.max(0, top / H);
  const b = Math.max(0, 1 - (top + sh) / H);
  return l + r < 1 && t + b < 1 ? { l, t, r, b } : undefined;
}

// ---------------------------------------------------------------------------------------------
// WMF

function playWmf(bytes: Uint8Array, target: MetaTarget): MetaOp[] | null {
  const d = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const p = new Player(target);
  const s16 = (o: number) => d.getInt16(o, true);
  const u16 = (o: number) => d.getUint16(o, true);
  const u32 = (o: number) => d.getUint32(o, true);
  let off = 0;
  let bbox: [number, number, number, number] | null = null;
  if (u32(0) === 0x9ac6cdd7) {
    bbox = [s16(6), s16(8), s16(10), s16(12)];
    off = 22;
  }
  const headerWords = u16(off + 2);
  const numObjects = u16(off + 10);
  off += headerWords * 2;
  const objects: (GdiObject | undefined)[] = new Array(Math.max(numObjects, 1)).fill(undefined);
  const free = () => {
    const i = objects.findIndex((o) => o === undefined);
    return i < 0 ? objects.length : i;
  };
  // The window (from the placeable bounds, then the window records) maps onto the target; WMF has
  // no world transform or reference device.
  p.wmf = true;
  if (bbox) {
    p.dc.wOrg = [bbox[0], bbox[1]];
    p.dc.wExt = [bbox[2] - bbox[0] || 1, bbox[3] - bbox[1] || 1];
  }
  let windowKnown = !!bbox;
  const points = (o: number, n: number): [number, number][] => {
    const out: [number, number][] = [];
    for (let i = 0; i < n; i++) out.push([s16(o + i * 4), s16(o + i * 4 + 2)]);
    return out;
  };

  let records = 0;
  while (off + 6 <= bytes.length && records++ < 2_000_000) {
    const size = u32(off) * 2;
    const fn = u16(off + 4);
    if (size < 6 || off + size > bytes.length) break;
    const r = off + 6;
    off += size;
    switch (fn) {
      case 0x0000: // EOF
        return windowKnown ? p.ops : null;
      case 0x020b: // SETWINDOWORG
        p.dc.wOrg = [s16(r + 2), s16(r)];
        break;
      case 0x020c: // SETWINDOWEXT
        p.dc.wExt = [s16(r + 2), s16(r)];
        windowKnown = true;
        break;
      case 0x020f: // OFFSETWINDOWORG
        p.dc.wOrg = [p.dc.wOrg[0] + s16(r + 2), p.dc.wOrg[1] + s16(r)];
        break;
      case 0x0103: // SETMAPMODE
        break;
      case 0x0102: // SETBKMODE
        p.dc.bkMode = u16(r);
        break;
      case 0x0201: // SETBKCOLOR
        p.dc.bkColor = colorRef(d, r);
        break;
      case 0x0209: // SETTEXTCOLOR
        p.dc.textColor = colorRef(d, r);
        break;
      case 0x012e: // SETTEXTALIGN
        p.dc.textAlign = u16(r);
        break;
      case 0x0106: // SETPOLYFILLMODE
        p.dc.fillMode = u16(r);
        break;
      case 0x001e: // SAVEDC
        p.save();
        break;
      case 0x0127: // RESTOREDC
        p.restore(s16(r));
        break;
      case 0x02fa: // CREATEPENINDIRECT
        objects[free()] = { kind: 'pen', style: u16(r), width: s16(r + 2), color: colorRef(d, r + 6) };
        break;
      case 0x02fc: // CREATEBRUSHINDIRECT
        objects[free()] = { kind: 'brush', style: u16(r), color: colorRef(d, r + 2) };
        break;
      case 0x02fb: // CREATEFONTINDIRECT
        objects[free()] = {
          kind: 'font',
          height: s16(r),
          escapement: s16(r + 4),
          weight: s16(r + 8),
          italic: !!bytes[r + 10],
          underline: !!bytes[r + 11],
          strike: !!bytes[r + 12],
          face: readFace(bytes, r + 18, false),
        };
        break;
      case 0x00f7: // CREATEPALETTE
      case 0x01f9: // CREATEPATTERNBRUSH
      case 0x06ff: // CREATEREGION
        objects[free()] = { kind: 'other' };
        break;
      case 0x0142: // DIBCREATEPATTERNBRUSH
        objects[free()] = { kind: 'brush', style: 0, color: '808080' };
        break;
      case 0x012d: {
        // SELECTOBJECT
        p.select(objects[u16(r)]);
        break;
      }
      case 0x01f0: // DELETEOBJECT
        objects[u16(r)] = undefined;
        break;
      case 0x0214: // MOVETO
        p.cur = [s16(r + 2), s16(r)];
        break;
      case 0x0213: {
        // LINETO
        const to: [number, number] = [s16(r + 2), s16(r)];
        p.line(p.pt(...p.cur), p.pt(...to));
        p.cur = to;
        break;
      }
      case 0x0325: // POLYLINE
        p.poly(points(r + 2, u16(r)), false);
        break;
      case 0x0324: // POLYGON
        p.poly(points(r + 2, u16(r)), true);
        break;
      case 0x0538: {
        // POLYPOLYGON
        const n = u16(r);
        let at = r + 2 + n * 2;
        const segs: MetaSeg[] = [];
        for (let i = 0; i < n; i++) {
          const count = u16(r + 2 + i * 2);
          const pts = points(at, count);
          at += count * 4;
          if (!pts.length) continue;
          segs.push({ op: 'M', pts: p.pt(...pts[0]) });
          for (const q of pts.slice(1)) segs.push({ op: 'L', pts: p.pt(...q) });
          segs.push({ op: 'Z', pts: [] });
        }
        p.draw(segs, true, true);
        break;
      }
      case 0x041b: // RECTANGLE
        p.rect(s16(r + 6), s16(r + 4), s16(r + 2), s16(r));
        break;
      case 0x061c: // ROUNDRECT
        p.draw(p.roundRectSegs(s16(r + 10), s16(r + 8), s16(r + 6), s16(r + 4), s16(r + 2), s16(r)), true, true);
        break;
      case 0x0418: // ELLIPSE
        p.draw(p.ellipseSegs(s16(r + 6), s16(r + 4), s16(r + 2), s16(r)), true, true);
        break;
      case 0x0817: // ARC
      case 0x081a: // PIE
      case 0x0830: {
        // CHORD
        const kind = fn === 0x081a ? 'pie' : fn === 0x0830 ? 'chord' : 'arc';
        p.draw(p.arcSegs(s16(r + 14), s16(r + 12), s16(r + 10), s16(r + 8), s16(r + 6), s16(r + 4), s16(r + 2), s16(r), kind, true), kind !== 'arc', true);
        break;
      }
      case 0x061d: {
        // PATBLT
        const rop = u32(r);
        const [h, w, y, x] = [s16(r + 4), s16(r + 6), s16(r + 8), s16(r + 10)];
        const color = rop === 0x42 ? '000000' : rop === 0xff0062 ? 'FFFFFF' : rop === 0xf00021 && p.dc.brush.style !== 1 ? p.dc.brush.color : null;
        if (!color) break;
        const [brush, pen] = [p.dc.brush, p.dc.pen];
        p.dc.brush = { kind: 'brush', style: 0, color };
        p.dc.pen = NULL_PEN;
        p.rect(x, y, x + w, y + h);
        p.dc.brush = brush;
        p.dc.pen = pen;
        break;
      }
      case 0x0416: {
        // INTERSECTCLIPRECT
        const [b, rr, t, l] = [s16(r), s16(r + 2), s16(r + 4), s16(r + 6)];
        p.clipTo([{ op: 'M', pts: p.pt(l, t) }, { op: 'L', pts: p.pt(rr, t) }, { op: 'L', pts: p.pt(rr, b) }, { op: 'L', pts: p.pt(l, b) }, { op: 'Z', pts: [] }], false);
        break;
      }
      case 0x0521: {
        // TEXTOUT
        const n = u16(r);
        let s = '';
        for (let i = 0; i < n; i++) s += String.fromCharCode(bytes[r + 2 + i]);
        const at = r + 2 + n + (n & 1);
        p.text(s16(at + 2), s16(at), s, undefined);
        break;
      }
      case 0x0a32: {
        // EXTTEXTOUT
        const y = s16(r);
        const x = s16(r + 2);
        const n = u16(r + 4);
        const options = u16(r + 6);
        let at = r + 8 + (options & 6 ? 8 : 0);
        let s = '';
        for (let i = 0; i < n && at + i < r + size - 6; i++) s += String.fromCharCode(bytes[at + i]);
        at += n + (n & 1);
        let dx: number[] | undefined;
        if (at + n * 2 <= r + size - 6) {
          dx = [];
          for (let i = 0; i < n; i++) dx.push(s16(at + i * 2));
        }
        p.text(x, y, s, dx);
        break;
      }
      case 0x0f43: {
        // STRETCHDIB
        const rop = u32(r);
        if (!copiesSource(rop)) break;
        const [sh, sw, sy, sx, dh, dw, dy, dx] = [s16(r + 6), s16(r + 8), s16(r + 10), s16(r + 12), s16(r + 14), s16(r + 16), s16(r + 18), s16(r + 20)];
        const dib = r + 22;
        const head = dibHeaderLength(d, dib);
        p.image(dibParts(bytes, off, dib, head, dib + head, r + size - 6 - dib - head), dx, dy, dw, dh, sourceCrop(d, dib, sx, sy, sw, sh));
        break;
      }
      case 0x0b41: // DIBSTRETCHBLT
      case 0x0940: {
        // DIBBITBLT
        const rop = u32(r);
        const stretch = fn === 0x0b41;
        const hasBitmap = size / 2 !== (stretch ? 3 + 2 + 8 + 1 : 3 + 2 + 6 + 1);
        const base = r + 4;
        const vals: number[] = [];
        for (let i = 0; i < (stretch ? 8 : 6) + (hasBitmap ? 0 : 1); i++) vals.push(s16(base + i * 2));
        if (!hasBitmap) vals.splice(stretch ? 4 : 2, 1);
        const [sh, sw, sy, sx, dh, dw, dy, dx] = stretch ? vals : [vals[2], vals[3], vals[0], vals[1], vals[2], vals[3], vals[4], vals[5]];
        if (!hasBitmap) {
          const color = rop === 0x42 ? '000000' : rop === 0xff0062 ? 'FFFFFF' : rop === 0xf00021 && p.dc.brush.style !== 1 ? p.dc.brush.color : null;
          if (color) {
            const [brush, pen] = [p.dc.brush, p.dc.pen];
            p.dc.brush = { kind: 'brush', style: 0, color };
            p.dc.pen = NULL_PEN;
            p.rect(dx, dy, dx + dw, dy + dh);
            p.dc.brush = brush;
            p.dc.pen = pen;
          }
          break;
        }
        if (!copiesSource(rop)) break;
        const dib = base + (stretch ? 16 : 12);
        const head = dibHeaderLength(d, dib);
        p.image(dibParts(bytes, off, dib, head, dib + head, r + size - 6 - dib - head), dx, dy, dw, dh, sourceCrop(d, dib, sx, sy, sw, sh));
        break;
      }
      case 0x0d33: {
        // SETDIBTODEV
        const [sy, sx, h, w, dy, dx] = [s16(r + 6), s16(r + 8), s16(r + 10), s16(r + 12), s16(r + 14), s16(r + 16)];
        const dib = r + 18;
        const head = dibHeaderLength(d, dib);
        p.image(dibParts(bytes, off, dib, head, dib + head, r + size - 6 - dib - head), dx, dy, w, h, sourceCrop(d, dib, sx, sy, w, h));
        break;
      }
      default:
        break;
    }
  }
  return windowKnown ? p.ops : null;
}
