import { singleSubstitutions } from '../fonts/gsub.js';
import type { FontFace } from '../fonts/registry.js';
import type { DrawOp, GlyphOp } from './types.js';

/**
 * Vertical text (a section with w:textDirection tbRl): its lines are laid out across a page turned
 * a quarter, then turned back so they run down the page and follow one another leftwards. East
 * Asian characters stand upright in their line, in the font's vertical forms; other text lies on
 * its side.
 */

/** Characters set upright in vertical text: East Asian scripts, symbols and full-width forms. */
export function uprightInVertical(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0xa4cf) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xac00 && cp <= 0xd7ff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe10 && cp <= 0xfe1f) ||
    (cp >= 0xfe30 && cp <= 0xfe4f) ||
    (cp >= 0xff01 && cp <= 0xff60) ||
    (cp >= 0xffe0 && cp <= 0xffe6) ||
    (cp >= 0x1f200 && cp <= 0x1f2ff) ||
    (cp >= 0x20000 && cp <= 0x3fffd)
  );
}

const verticalForms = new WeakMap<FontFace, Map<number, number>>();

/** A face's vertical alternates (OpenType 'vert'): brackets turned, commas and stops moved to the corner. */
function verticalFormsOf(face: FontFace): Map<number, number> {
  let out = verticalForms.get(face);
  if (!out) {
    const gsub = face.font.tableBytes('GSUB');
    out = gsub ? singleSubstitutions(gsub, 'vert') : new Map<number, number>();
    if (gsub && !out.size) out = singleSubstitutions(gsub, 'vrt2');
    verticalForms.set(face, out);
  }
  return out;
}

/**
 * A page's ops laid out on the turned page (`pageWidth` is the real page's width), turned to run
 * down the real page, with East Asian glyphs set upright about the centre of their em box.
 */
export function turnVertical(ops: DrawOp[], pageWidth: number): DrawOp[] {
  return [{ kind: 'transform', matrix: [0, 1, -1, 0, pageWidth, 0] }, ...uprightEastAsian(ops), { kind: 'restore' }];
}

/** Ops of text that will be turned a quarter clockwise, its East Asian glyphs set to stand upright. */
export function uprightEastAsian(ops: DrawOp[]): DrawOp[] {
  return ops.flatMap((op) => (op.kind === 'glyphs' ? uprightGlyphs(op) : [op]));
}

/** A glyph run split so its East Asian glyphs stand upright and the rest lie along the line. */
function uprightGlyphs(op: GlyphOp): DrawOp[] {
  const upright = op.texts.map((t) => uprightInVertical(t.codePointAt(0) ?? 0));
  if (!upright.some(Boolean)) return [op];
  const out: DrawOp[] = [];
  const forms = verticalFormsOf(op.face);
  const scale = (op.size * op.scaleX) / op.face.font.unitsPerEm;
  // The em box an ideograph fills in a horizontal line: 0.88em above the baseline, 0.12em below.
  const cy = op.y - 0.38 * op.size;
  let field = op.field;
  const piece = (from: number, to: number): GlyphOp => {
    const p: GlyphOp = { ...op, gids: op.gids.slice(from, to), xs: op.xs.slice(from, to), texts: op.texts.slice(from, to), field };
    if (!field) delete p.field;
    field = undefined;
    return p;
  };
  for (let i = 0; i < op.gids.length; ) {
    let j = i + 1;
    while (j < op.gids.length && upright[j] === upright[i] && !upright[i]) j++;
    if (!upright[i]) {
      out.push(piece(i, j));
      i = j;
      continue;
    }
    const g = piece(i, i + 1);
    g.gids = [forms.get(op.gids[i]) ?? op.gids[i]];
    const advance = i + 1 < op.xs.length ? op.xs[i + 1] - op.xs[i] : op.face.font.advance(op.gids[i]) * scale;
    const cx = op.x + op.xs[i] + advance / 2;
    // Turned back a quarter about the em box's centre, the glyph stands upright on the page.
    out.push({ kind: 'transform', matrix: [0, -1, 1, 0, cx - cy, cx + cy] }, g, { kind: 'restore' });
    i = j;
  }
  return out;
}
