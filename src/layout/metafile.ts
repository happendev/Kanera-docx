import type { ImageRef } from '../ooxml/model.js';
import { DEFAULT_RUN } from '../ooxml/props.js';
import type { FontRegistry } from '../fonts/registry.js';
import { playMetafile, type MetaOp, type MetaSeg } from '../image/metafile.js';
import { sniffImage } from '../image/image.js';
import { shapeText } from './shaper.js';
import type { DrawOp, GlyphOp } from './types.js';

type Crop = { l: number; t: number; r: number; b: number } | undefined;

/**
 * An EMF or WMF picture as vector drawing commands filling the image box (x, y, w, h); null when
 * the data is not a metafile or cannot be played. Cropped pictures are clipped to the box.
 */
export function metafileOps(image: ImageRef, x: number, y: number, w: number, h: number, crop: Crop, fonts: FontRegistry): DrawOp[] | null {
  const fmt = sniffImage(image.data);
  if (fmt !== 'emf' && fmt !== 'wmf') return null;
  const c = crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const fw = 1 - c.l - c.r;
  const fh = 1 - c.t - c.b;
  if (fw <= 0 || fh <= 0) return [];
  const W = w / fw;
  const H = h / fh;
  const meta = playMetafile(image.data, { x: x - c.l * W, y: y - c.t * H, w: W, h: H });
  if (!meta) return null;
  const out: DrawOp[] = [];
  const cropped = !!(c.l || c.t || c.r || c.b);
  if (cropped) out.push({ kind: 'clip', x, y, w, h });
  let clip: MetaSeg[][] = [];
  let images = 0;
  for (const op of meta) {
    if (op.clip !== clip && !sameClip(op.clip, clip)) {
      for (let i = 0; i < clip.length; i++) out.push({ kind: 'restore' });
      for (const segs of op.clip) out.push({ kind: 'clip', x: 0, y: 0, w: 0, h: 0, segs });
    }
    clip = op.clip;
    if (op.kind === 'path') out.push({ kind: 'path', segs: op.segs, fill: op.fill ?? null, stroke: op.stroke ?? null, strokeWidth: op.width, nonzero: !op.evenOdd });
    else if (op.kind === 'image')
      out.push({ kind: 'image', x: op.x, y: op.y, w: op.w, h: op.h, image: { id: `${image.id}#${++images}`, data: op.dib.header, dibBits: op.dib.bits, contentType: 'image/bmp' }, crop: op.crop, rotation: 0, flipH: false, flipV: false });
    else out.push(...textOps(op, fonts));
  }
  for (let i = 0; i < clip.length; i++) out.push({ kind: 'restore' });
  if (cropped) out.push({ kind: 'restore' });
  return out;
}

function sameClip(a: MetaSeg[][], b: MetaSeg[][]): boolean {
  return a.length === b.length && a.every((s, i) => s === b[i]);
}

function textOps(op: Extract<MetaOp, { kind: 'text' }>, fonts: FontRegistry): DrawOp[] {
  const f = op.font;
  const props = { ...DEFAULT_RUN, fontAscii: f.family, fontHAnsi: f.family, fontEastAsia: f.family, fontCs: f.family, bold: f.bold, italic: f.italic, boldCs: f.bold, italicCs: f.italic, size: 100, sizeCs: 100 };
  const probe = shapeText(op.text, props, fonts);
  if (!probe.length) return [];
  // A font's cell height is its ascent plus descent; the em is what is left of that.
  const cell = Math.max(...probe.map((r) => r.ascent + r.descent));
  const size = f.cell && cell > 0 ? (f.size * 100) / cell : f.size;
  const scale = size / 100;
  const ascent = Math.max(...probe.map((r) => r.ascent)) * scale;
  const descent = Math.max(...probe.map((r) => r.descent)) * scale;
  const glyphs = probe.flatMap((r) => r.glyphs.map((g) => ({ run: r, g })));
  // Character advances from the record, where they line up one to one with the glyphs.
  const dx = op.dx && op.dx.length === glyphs.length ? op.dx : undefined;
  const advances = glyphs.map(({ g }, i) => (dx ? dx[i] : g.advance * scale));
  const width = advances.reduce((a, b) => a + b, 0);
  const x0 = op.align === 1 ? -width / 2 : op.align === 2 ? -width : 0;
  const base = op.valign === 'baseline' ? 0 : op.valign === 'top' ? ascent : -descent;
  const ops: DrawOp[] = [];
  const rotate = Math.abs(op.angle) > 0.01;
  if (rotate) {
    const t = (-op.angle * Math.PI) / 180;
    const [cos, sin] = [Math.cos(t), Math.sin(t)];
    ops.push({ kind: 'transform', matrix: [cos, sin, -sin, cos, op.x - (cos * op.x - sin * op.y), op.y - (sin * op.x + cos * op.y)] });
  }
  let px = 0;
  let current: GlyphOp | null = null;
  glyphs.forEach(({ run, g }, i) => {
    if (!current || current.face !== run.face.face) {
      current = {
        kind: 'glyphs',
        x: op.x + x0 + px,
        y: op.y + base,
        face: run.face.face,
        size: run.size * scale,
        gids: [],
        xs: [],
        texts: [],
        color: op.color,
        syntheticBold: run.face.syntheticBold,
        syntheticItalic: run.face.syntheticItalic,
        scaleX: 1,
      };
      ops.push(current);
    }
    if (g.gid >= 0) {
      current.gids.push(g.gid);
      current.xs.push(op.x + x0 + px - current.x);
      current.texts.push(g.text);
    }
    px += advances[i];
  });
  const line = (dy: number) =>
    ops.push({ kind: 'line', x1: op.x + x0, y1: op.y + base + dy, x2: op.x + x0 + width, y2: op.y + base + dy, width: Math.max(0.5, size * 0.05), color: op.color });
  if (f.underline) line(size * 0.12);
  if (f.strike) line(-size * 0.27);
  if (rotate) ops.push({ kind: 'restore' });
  return ops;
}
