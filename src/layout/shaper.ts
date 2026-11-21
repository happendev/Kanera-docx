import { asksForEmoji, emojiSequence, isTag, isVariationSelector } from '../fonts/emoji.js';
import { fontFor } from '../fonts/script.js';
import type { FontRegistry, ResolvedFace } from '../fonts/registry.js';
import type { Link, RunProps } from '../docx/model.js';

export interface ShapedGlyph {
  /** -1 for zero-width characters that draw nothing. */
  gid: number;
  text: string;
  cp: number;
  advance: number;
  /** How far the glyph is drawn from its position (a proportional form's shift). */
  dx?: number;
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

const ZERO_WIDTH = new Set([0x200b, 0x200c, 0x200d, 0x200e, 0x200f, 0xfeff, 0x2060, 0x202a, 0x202b, 0x202c, 0x202d, 0x202e]);
/** Punctuation a fixed-pitch East Asian font (SimSun, MS Mincho) draws an em wide, like its ideographs. */
const FULL_WIDTH_PUNCT = new Set([0xb7, 0x2014, 0x2015, 0x2018, 0x2019, 0x201c, 0x201d, 0x2026]);

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
  let previousChoice: ReturnType<typeof fontFor> | undefined;
  let primary: ResolvedFace | undefined;
  const chars = Array.from(text);
  const cps = chars.map((c) => c.codePointAt(0)!);
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    let cp = cps[i];
    if (cur && (ZERO_WIDTH.has(cp) || isVariationSelector(cp) || isTag(cp))) {
      cur.glyphs.push({ gid: -1, text: '', cp, advance: 0 });
      continue;
    }
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
    if (!previousChoice || choice.family !== previousChoice.family || choice.bold !== previousChoice.bold || choice.italic !== previousChoice.italic) {
      primary = fonts.resolve(choice.family, choice.bold, choice.italic, props.eastAsiaLang);
      previousChoice = choice;
    }
    // An emoji sequence the font lacks prints from the emoji font, as one glyph where it has one.
    const emoji = !small && (cp < 0x80 ? cps[i + 1] === 0xfe0f || cps[i + 1] === 0x20e3 : asksForEmoji(cps[i + 1]) || !primary!.face.font.hasGlyph(cp)) ? emojiSequence(cps, i) : 0;
    const emojiFace = emoji ? fonts.emojiFace(choice.family, choice.bold, choice.italic, primary!) : undefined;
    if (emojiFace && cps.slice(i, i + emoji).every((c) => isVariationSelector(c) || isTag(c) || c === 0x200d || emojiFace.face.font.hasGlyph(c))) {
      const look = emojiFace.standIn ?? { scale: 1, drop: 0, advance: 1 };
      if (!cur || cur.face !== emojiFace || cur.size !== sm.size * look.scale || cur.scaleX !== scale) {
        cur = newRun(emojiFace, sm.size * look.scale, size, sm.shift, scale, props, link);
        // The drop moves the glyphs only, not the line's top or bottom.
        cur.baselineShift -= look.drop * sm.size;
        cur.ascent += look.drop * sm.size;
        cur.descent -= look.drop * sm.size;
        out.push(cur);
      }
      shapeEmoji(cps.slice(i, i + emoji), cur, props.spacing);
      i += emoji - 1;
      continue;
    }
    const resolved: ResolvedFace = fonts.faceForChar(choice.family, choice.bold, choice.italic, cp, props.eastAsiaLang, props.drawn, primary);
    const glyphSize = sm.size;
    const sx = scale * (resolved.glyphScale ?? 1);
    if (!cur || cur.face !== resolved || cur.size !== glyphSize || cur.scaleX !== sx) {
      cur = newRun(resolved, glyphSize, size, sm.shift, sx, props, link);
      out.push(cur);
    }
    const font = resolved.face.font;
    const mapped = resolved.char ?? (resolved.symbolMap ? fonts.symbolChar(resolved.symbolMap, cp) : cp);
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
    let advance = (font.advance(gid) / font.unitsPerEm) * glyphSize * sx;
    const half = resolved.halfWidth && cp >= 0x20 && cp <= 0x7e;
    const full = !half && FULL_WIDTH_PUNCT.has(cp) ? resolved.fullWidth?.get(gid) : undefined;
    if (half) {
      gid = resolved.halfWidth!.get(gid) ?? gid;
      advance = (glyphSize * sx) / 2;
    } else if (full !== undefined) {
      gid = full;
      advance = glyphSize * sx;
    }
    const prop = half || full !== undefined ? undefined : resolved.proportional?.get(cp);
    if (prop) advance += prop.advance * glyphSize * sx;
    if (cp === 0x2003 && gid === font.glyphId(0x20)) advance = glyphSize * sx;
    if (cp === 0x2002 && gid === font.glyphId(0x20)) advance = (glyphSize * sx) / 2;
    advance += props.spacing;
    // Pair kerning applies when w:kern is set and the font is at least that size.
    if (props.kern > 0 && size >= props.kern && cur.glyphs.length > 0 && !half && full === undefined) {
      const prev = cur.glyphs[cur.glyphs.length - 1];
      if (prev.gid > 0 && gid > 0) prev.advance += (resolved.face.kerning(prev.gid, gid) / font.unitsPerEm) * glyphSize * sx;
    }
    // Symbol-font characters (Wingdings, Symbol PUA codes) copy as the Unicode they were mapped to.
    const text = mapped !== cp && !resolved.char && (mapped < 0xe000 || mapped > 0xf8ff) ? String.fromCodePoint(mapped) : display;
    cur.glyphs.push(prop?.x ? { gid, text, cp, advance, dx: prop.x * glyphSize * sx } : { gid, text, cp, advance });
  }
  return out;
}

function newRun(face: ResolvedFace, glyphSize: number, size: number, shift: number, scaleX: number, props: RunProps, link?: Link): ShapedRun {
  return {
    face,
    size: glyphSize,
    props,
    glyphs: [],
    baselineShift: shift + props.position,
    // Superscript and subscript text sizes the line as if it were neither smaller nor shifted
    // (Word's rows of x¹ marks line up with plain cells); a w:position raise still counts.
    ascent: face.ascent * size - shift,
    descent: face.descent * size + shift,
    extLeading: face.externalLeading * size,
    scaleX,
    link,
  };
}

/**
 * Add an emoji sequence to `run`, whose face is an emoji font: the font's ligatures (OpenType ccmp)
 * join a person and skin tone, or the people of a family, into one glyph. Variation selectors are
 * passed over in matching, as HarfBuzz and DirectWrite pass over default-ignorable characters, and
 * a joiner or tag left over draws nothing.
 */
function shapeEmoji(seq: number[], run: ShapedRun, spacing: number): void {
  const font = run.face.face.font;
  const ligatures = run.face.face.ligatures('ccmp');
  const look = run.face.standIn;
  const widen = look ? look.advance / look.scale : 1;
  const items = seq.flatMap((cp, k) => (isVariationSelector(cp) ? [] : [{ gid: font.glyphId(cp), k }]));
  for (let p = 0; p < items.length; ) {
    let glyph = items[p].gid;
    let next = p + 1;
    for (const lig of ligatures.get(glyph) ?? []) {
      if (lig.rest.every((g, r) => items[p + 1 + r]?.gid === g)) {
        glyph = lig.glyph;
        next = p + 1 + lig.rest.length;
        break;
      }
    }
    const cp = seq[items[p].k];
    const text = String.fromCodePoint(...seq.slice(items[p].k, next < items.length ? items[next].k : seq.length));
    if (next === p + 1 && (cp === 0x200d || isTag(cp))) run.glyphs.push({ gid: -1, text: '', cp, advance: 0 });
    else {
      // A stand-in's glyph is centred in Segoe UI Emoji's advance.
      const own = (font.advance(glyph) / font.unitsPerEm) * run.size * run.scaleX;
      const advance = own * widen;
      run.glyphs.push(widen === 1 ? { gid: glyph, text, cp, advance: advance + spacing } : { gid: glyph, text, cp, advance: advance + spacing, dx: (advance - own) / 2 });
    }
    p = next;
  }
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
