import type { FontFace, FontRegistry, ResolvedFace } from '../fonts/registry.js';
import { mapSymbolChar } from '../fonts/symbol-map.js';
import type { Link, RunProps } from '../docx/model.js';

export interface ShapedGlyph {
  /** -1 for zero-width characters that draw nothing. */
  gid: number;
  text: string;
  cp: number;
  advance: number;
}

export interface ShapedRun {
  face: ResolvedFace;
  size: number;
  props: RunProps;
  glyphs: ShapedGlyph[];
  /** Raise (positive) or lower (negative) relative to the line's baseline. */
  baselineShift: number;
  ascent: number;
  descent: number;
  extLeading: number;
  scaleX: number;
  link?: Link;
}

export function isEastAsian(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0x2fdf) ||
    (cp >= 0x2ff0 && cp <= 0x303f) ||
    (cp >= 0x3040 && cp <= 0x9fff) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xac00 && cp <= 0xd7ff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe30 && cp <= 0xfe4f) ||
    (cp >= 0xff00 && cp <= 0xffef) ||
    (cp >= 0x20000 && cp <= 0x2fa1f)
  );
}

export function isComplexScript(cp: number): boolean {
  return (cp >= 0x0590 && cp <= 0x08ff) || (cp >= 0x0900 && cp <= 0x0dff) || (cp >= 0x0e00 && cp <= 0x0eff) || (cp >= 0xfb1d && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff);
}

/** Characters that take the East Asian font when the run hints eastAsia (quotes, some symbols). */
function eastAsianAmbiguous(cp: number): boolean {
  return (cp >= 0x2018 && cp <= 0x201f) || cp === 0x2026 || cp === 0x2025 || (cp >= 0x2010 && cp <= 0x2015) || cp === 0xb7 || (cp >= 0x2460 && cp <= 0x24ff) || (cp >= 0x25a0 && cp <= 0x26ff);
}

const ZERO_WIDTH = new Set([0x200b, 0x200c, 0x200d, 0x200e, 0x200f, 0xfeff, 0x2060, 0x202a, 0x202b, 0x202c, 0x202d, 0x202e]);

interface FontChoice {
  family: string;
  bold: boolean;
  italic: boolean;
  size: number;
}

function fontFor(cp: number, p: RunProps): FontChoice {
  if (p.cs || p.rtl || isComplexScript(cp)) return { family: p.fontCs, bold: p.boldCs, italic: p.italicCs, size: p.sizeCs };
  if (isEastAsian(cp) || (p.hint === 'eastAsia' && eastAsianAmbiguous(cp))) return { family: p.fontEastAsia, bold: p.bold, italic: p.italic, size: p.size };
  if (cp < 0x80) return { family: p.fontAscii, bold: p.bold, italic: p.italic, size: p.size };
  return { family: p.fontHAnsi, bold: p.bold, italic: p.italic, size: p.size };
}

/**
 * Effective glyph size and baseline shift for superscript/subscript. Word works in half-points:
 * the script size is two thirds of the run's half-points rounded down, and a superscript rises by
 * the rest (11pt: 7pt raised 4pt; 10pt: 6.5pt raised 3.5pt, as in Word's PDFs).
 */
export function scriptMetrics(size: number, vertAlign: RunProps['vertAlign']): { size: number; shift: number } {
  if (vertAlign === 'baseline') return { size, shift: 0 };
  const half = Math.round(size * 2);
  const small = Math.max(1, Math.floor((half * 2) / 3 + 1e-9)) / 2;
  if (vertAlign === 'superscript') return { size: small, shift: size - small };
  return { size: small, shift: -size * 0.14 };
}

export interface ShapeOptions {
  /** Ignore caps transformations (used for numbering labels in some cases). */
  raw?: boolean;
}

/**
 * Convert text in one run into glyph runs, splitting where the font changes (script slots,
 * fallback fonts, small caps).
 */
export function shapeText(text: string, props: RunProps, fonts: FontRegistry, link?: Link, opts: ShapeOptions = {}): ShapedRun[] {
  const out: ShapedRun[] = [];
  let cur: ShapedRun | null = null;
  const scale = props.scale / 100;
  for (const ch of text) {
    let cp = ch.codePointAt(0)!;
    let display = ch;
    let small = false;
    if (!opts.raw && (props.caps || props.smallCaps)) {
      const up = ch.toUpperCase();
      if (up !== ch && [...up].length === 1) {
        if (props.smallCaps && !props.caps) small = true;
        display = up;
        cp = up.codePointAt(0)!;
      }
    }
    const choice = fontFor(cp, props);
    let size = choice.size;
    if (small) size *= 0.8;
    const sm = scriptMetrics(size, props.vertAlign);
    const resolved: ResolvedFace = fonts.faceForChar(choice.family, choice.bold, choice.italic, cp);
    const glyphSize = sm.size;
    if (!cur || cur.face !== resolved || cur.size !== glyphSize) {
      const face: FontFace = resolved.face;
      cur = {
        face: resolved,
        size: glyphSize,
        props,
        glyphs: [],
        baselineShift: sm.shift + props.position,
        // Superscript and subscript text sizes the line as if it were neither smaller nor shifted
        // (Word's rows of x¹ marks line up with plain cells); a w:position raise still counts.
        ascent: resolved.ascent * size - sm.shift,
        descent: resolved.descent * size + sm.shift,
        extLeading: resolved.externalLeading * size,
        scaleX: scale,
        link,
      };
      out.push(cur);
    }
    const font = resolved.face.font;
    const mapped = resolved.symbolMap ? mapSymbolChar(resolved.symbolMap, cp) : cp;
    let gid = font.glyphId(mapped);
    let advanceCp = mapped;
    if (ZERO_WIDTH.has(cp)) {
      cur.glyphs.push({ gid: -1, text: '', cp, advance: 0 });
      continue;
    }
    if (cp === 0xad) {
      // Soft hyphen: invisible unless the line breaks here (handled by the line breaker).
      cur.glyphs.push({ gid: font.glyphId(0x2d), text: '-', cp, advance: 0 });
      continue;
    }
    if (gid === 0) {
      if (cp === 0xa0 || cp === 0x202f || cp === 0x2007) advanceCp = 0x20;
      else if (cp === 0x2011 || cp === 0x2010) advanceCp = 0x2d;
      else if (cp === 0x2002 || cp === 0x2003 || cp === 0x2009 || cp === 0x2005 || cp === 0x2006) advanceCp = 0x20;
      if (advanceCp !== mapped) gid = font.glyphId(advanceCp);
    }
    let advance = (font.advance(gid) / font.unitsPerEm) * glyphSize * scale;
    if (cp === 0x2003 && gid === font.glyphId(0x20)) advance = glyphSize * scale;
    if (cp === 0x2002 && gid === font.glyphId(0x20)) advance = (glyphSize * scale) / 2;
    advance += props.spacing;
    // Pair kerning applies when w:kern is set and the font is at least that size.
    if (props.kern > 0 && size >= props.kern && cur.glyphs.length > 0) {
      const prev = cur.glyphs[cur.glyphs.length - 1];
      if (prev.gid > 0 && gid > 0) prev.advance += (font.kerning(prev.gid, gid) / font.unitsPerEm) * glyphSize * scale;
    }
    // Symbol-font characters (Wingdings, Symbol PUA codes) copy as the Unicode they were mapped to.
    const text = mapped !== cp && (mapped < 0xe000 || mapped > 0xf8ff) ? String.fromCodePoint(mapped) : display;
    cur.glyphs.push({ gid, text, cp, advance });
  }
  return out;
}

/** Shape a string with the run's primary face only (tab leaders, field placeholders). */
export function shapeSimple(text: string, props: RunProps, fonts: FontRegistry): ShapedRun[] {
  return shapeText(text, { ...props, caps: false, smallCaps: false }, fonts);
}

export function runWidth(runs: ShapedRun[]): number {
  let w = 0;
  for (const r of runs) for (const g of r.glyphs) w += g.advance;
  return w;
}
