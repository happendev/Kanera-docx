import type { FontFace } from '../fonts/registry.js';
import type { Color, Drawing, ImageRef, Link } from '../docx/model.js';

/** A run of glyphs drawn with one font at one size. Coordinates: points, origin top-left, y down. */
export interface GlyphOp {
  kind: 'glyphs';
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
}

export type DrawOp = GlyphOp | RectOp | LineOp | ImageOp | PathOp | TransformOp | ClipOp | RestoreOp | LinkOp | AnchorOp | FloatOp;

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
    case 'float':
      return { ...op, x: op.x + dx, y: op.y + dy };
    case 'line':
      return { ...op, x1: op.x1 + dx, y1: op.y1 + dy, x2: op.x2 + dx, y2: op.y2 + dy };
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
