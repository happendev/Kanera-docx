import type { FontFace } from '../fonts/registry.js';
import type { Color, Drawing, FieldInline, ImageRef, Link } from '../ooxml/model.js';

/** A run of glyphs drawn with one font at one size. Coordinates: points, origin top-left, y down. */
export interface GlyphOp {
  kind: 'glyphs';
  /** Field beginning at this glyph run, for pagination-dependent values. */
  field?: FieldInline;
  x: number;
  /** Baseline y. */
  y: number;
  face: FontFace;
  size: number;
  gids: number[];
  /** Absolute x offset of each glyph relative to `x`. */
  xs: number[];
  /** Text each glyph represents (for copy/paste and search). */
  texts: string[];
  color: string;
  syntheticBold: boolean;
  syntheticItalic: boolean;
  /** Horizontal glyph scale (w:w), 1 = 100%. */
  scaleX: number;
  alpha?: number;
  outline?: boolean;
  /** Where the line's text starts, on lines set from their left edge: a renderer that snaps lines to a
   * device grid moves the whole line by one amount, so its spacing is kept. */
  lineX?: number;
}

export interface RectOp {
  kind: 'rect';
  x: number;
  y: number;
  w: number;
  h: number;
  fill: string;
  alpha?: number;
}

export interface LineOp {
  kind: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  width: number;
  color: string;
  dash?: number[];
  cap?: 'butt' | 'round' | 'square';
}

export interface ImageOp {
  kind: 'image';
  x: number;
  y: number;
  w: number;
  h: number;
  image: ImageRef;
  crop?: { l: number; t: number; r: number; b: number };
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  alpha?: number;
}

export interface PathSeg {
  op: 'M' | 'L' | 'C' | 'Z';
  pts: number[];
}

export interface PathOp {
  kind: 'path';
  segs: PathSeg[];
  fill?: string | null;
  fillAlpha?: number;
  stroke?: string | null;
  strokeWidth?: number;
  strokeAlpha?: number;
  dash?: number[];
  /** Fill by the nonzero winding rule rather than even-odd. */
  nonzero?: boolean;
  cap?: 'butt' | 'round' | 'square';
  join?: 'miter' | 'round' | 'bevel';
}

/** Where a shading's colour runs from position 0 to 1. */
export type ShadeGeometry =
  /** Along the axis from (x0, y0) to (x1, y1), and on beyond both ends. */
  | { type: 'axial'; x0: number; y0: number; x1: number; y1: number }
  /** Out from the centre to the ellipse with radii (rx, ry), and on beyond it. */
  | { type: 'radial'; cx: number; cy: number; rx: number; ry: number }
  /** Out from the inner rectangle to the outer one, in rectangular rings ([left, top, right, bottom]). */
  | { type: 'rect'; inner: [number, number, number, number]; outer: [number, number, number, number] };

/** A smooth gradient painted over the current clip, within `bbox` ([left, top, right, bottom]). */
export interface ShadeOp {
  kind: 'shade';
  geometry: ShadeGeometry;
  bbox: [number, number, number, number];
  /** Colours (RRGGBB) at evenly spaced positions from 0 to 1. */
  colors: string[];
  /** Opacity at the same positions, where it varies. */
  alphas?: number[];
  /** Opacity throughout, where it doesn't. */
  alpha?: number;
}

/** Save state and apply a transform (rotation/flip about a centre); paired with 'restore'. */
export interface TransformOp {
  kind: 'transform';
  /** PDF-style matrix [a b c d e f] in top-left page coordinates. */
  matrix: [number, number, number, number, number, number];
}
/** Save state and clip to a rectangle (or to `segs` when given); paired with 'restore'. */
export interface ClipOp {
  kind: 'clip';
  x: number;
  y: number;
  w: number;
  h: number;
  segs?: PathSeg[];
  /** Clip by the nonzero winding rule rather than even-odd. */
  nonzero?: boolean;
}
export interface RestoreOp {
  kind: 'restore';
}

export interface LinkOp {
  kind: 'link';
  x: number;
  y: number;
  w: number;
  h: number;
  link: Link;
}

/** Destination for internal links (bookmarks). */
export interface AnchorOp {
  kind: 'anchor';
  name: string;
  x: number;
  y: number;
}

/** Floating object anchored inside a box (resolved to page coordinates once the box is placed). */
export interface FloatOp {
  kind: 'float';
  drawing: Drawing;
  /** Position of the anchoring paragraph's top-left inside the box. */
  x: number;
  y: number;
  /** Width of the anchoring column/box. */
  areaWidth: number;
  /** The table cell holding the anchor: its text area's left and width, and its top and height. */
  cell?: { x: number; y: number; w: number; h: number };
}

export type DrawOp = GlyphOp | RectOp | LineOp | ImageOp | PathOp | ShadeOp | TransformOp | ClipOp | RestoreOp | LinkOp | AnchorOp | FloatOp;

export function appendOps(target: DrawOp[], ops: readonly DrawOp[]): void {
  for (const op of ops) target.push(op);
}

export interface Page {
  width: number;
  height: number;
  /** Paint layers, back to front. */
  behind: DrawOp[];
  headerFooter: DrawOp[];
  body: DrawOp[];
  front: DrawOp[];
  /** Section index the page belongs to. */
  section: number;
  /** Displayed page number (after pgNumType restarts). */
  number: number;
  numberFormat: string;
}

export function translateOps(ops: DrawOp[], dx: number, dy: number): DrawOp[] {
  if (dx === 0 && dy === 0) return ops;
  return ops.map((op) => translateOp(op, dx, dy));
}

export function translateOp(op: DrawOp, dx: number, dy: number): DrawOp {
  switch (op.kind) {
    case 'clip':
      return { ...op, x: op.x + dx, y: op.y + dy, segs: op.segs?.map((s) => ({ op: s.op, pts: s.pts.map((v, i) => (i % 2 === 0 ? v + dx : v + dy)) })) };
    case 'glyphs':
    case 'rect':
    case 'image':
    case 'link':
    case 'anchor':
      return { ...op, x: op.x + dx, y: op.y + dy };
    case 'float':
      return { ...op, x: op.x + dx, y: op.y + dy, cell: op.cell && { ...op.cell, x: op.cell.x + dx, y: op.cell.y + dy } };
    case 'line':
      return { ...op, x1: op.x1 + dx, y1: op.y1 + dy, x2: op.x2 + dx, y2: op.y2 + dy };
    case 'shade': {
      const g = op.geometry;
      const rect = ([l, t, r, b]: [number, number, number, number]): [number, number, number, number] => [l + dx, t + dy, r + dx, b + dy];
      const geometry: ShadeGeometry =
        g.type === 'axial' ? { ...g, x0: g.x0 + dx, y0: g.y0 + dy, x1: g.x1 + dx, y1: g.y1 + dy }
        : g.type === 'radial' ? { ...g, cx: g.cx + dx, cy: g.cy + dy }
        : { type: 'rect', inner: rect(g.inner), outer: rect(g.outer) };
      return { ...op, geometry, bbox: rect(op.bbox) };
    }
    case 'path':
      return {
        ...op,
        segs: op.segs.map((s) => ({ op: s.op, pts: s.pts.map((v, i) => (i % 2 === 0 ? v + dx : v + dy)) })),
      };
    case 'transform': {
      // Conjugate with the translation: T * M * T^-1
      const [a, b, c, d, e, f] = op.matrix;
      return { kind: 'transform', matrix: [a, b, c, d, e + dx - a * dx - c * dy, f + dy - b * dx - d * dy] };
    }
    case 'restore':
      return op;
  }
}

export function colorOrBlack(c: Color | undefined): string {
  return c ?? '000000';
}
