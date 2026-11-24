import { ACCENT_FORMS, INTEGRALS, isMathItalic, mathChar } from '../ooxml/math.js';
import type { MathNode, RunProps } from '../ooxml/model.js';
import type { FontRegistry } from '../fonts/registry.js';
import { normalizeFamily } from '../fonts/substitutes.js';
import type { SfntFont } from '../fonts/sfnt.js';
import { shapeText } from './shaper.js';
import { translateOps, type DrawOp, type GlyphOp, type PathSeg } from './types.js';

/**
 * An equation laid out as Word's equation editor sets it, in Cambria Math's proportions (measured
 * against Word's PDFs). Coordinates are relative to the box's left edge on its baseline, y down.
 */
export interface MathBox {
  w: number;
  /** Height and depth that structures place their parts by: Cambria Math's line for any text. */
  asc: number;
  desc: number;
  /** How far the ink reaches above and below the baseline, which sizes the equation's line. */
  top: number;
  bottom: number;
  ops: DrawOp[];
  /** Size of the box's text, which its spacing and scripts scale with. */
  em: number;
  /** TeX's atom class, for the space Word puts between neighbours. */
  cls: Cls;
  /** Where an equation array row's alignment mark (&) falls. */
  alignX?: number;
}

type Cls = 'ord' | 'op' | 'bin' | 'rel' | 'open' | 'close' | 'punct' | 'space' | 'align';

interface Style {
  /** Display style (an equation on its own line) sets limits and operators larger. */
  display: boolean;
  /** 0 text, 1 script, 2 script-script. */
  level: number;
  /** Denominators, subscripts and radicands keep superscripts lower. */
  cramped: boolean;
  /** Text size of the equation at this level, for structures with no text of their own. */
  size: number;
}

/** A box placed with its baseline origin at (x, y). */
interface Part {
  b: MathBox;
  x: number;
  y: number;
}

const SCALE = [1, 0.7, 0.55];
// In ems of the text size: the axis fraction bars and operators centre on, rule thickness, and the
// height and depth Cambria Math gives a line of text whatever its letters.
const AXIS = 0.29;
const RULE = 0.06;
const ASC = 0.7;
const DESC = 0.22;
const EQ_ROW_GAP = 0.33;

const BIN = new Set([...'+−±∓×÷·∙∗∘⊕⊖⊗⊘⊙∩∪∧∨⊎⊓⊔⋅⋆∖⊲⊳⊴⊵']);
const REL = new Set([...'=<>≤≥≦≧≠≈≡≢∼≃≅≆≇≉≐≔≕∝∈∉∋∌⊂⊃⊄⊅⊆⊇⊊⊋→←↔⇒⇐⇔↦⟶⟵⟹⟸⟺↑↓⇑⇓≪≫≺≻≼≽⊥∥∦∣∤⊢⊣⊨≲≳∶']);
const OPEN = new Set([...'([{⟨⌈⌊']);
const CLOSE = new Set([...')]}⟩⌉⌋!']);
const PUNCT = new Set([...',;']);
const SUMS = new Set([...'∑∏∐⋃⋂⋁⋀⨁⨂⨀⨄⨆']);
const X_HEIGHT = /^[acegmnopqrsuvwxyzıȷα-ωϵϑϕϖ\s]*$/;

/**
 * The formatting that sizes an equation's line like text: Word gives Cambria Math text Cambria's
 * ascent and descent (its stand-in's own reach further), and any extra ink grows the line apart.
 */
export function mathLineProps(props: RunProps): RunProps {
  return normalizeFamily(props.fontAscii ?? '') === 'cambria math' ? { ...props, fontAscii: 'Cambria', fontHAnsi: 'Cambria' } : props;
}

/** Lay out an equation; a displayed one in display style. */
export function layoutMath(node: MathNode, props: RunProps, display: boolean, fonts: FontRegistry): MathBox {
  return layout(node, { display, level: 0, cramped: false, size: props.size }, fonts);
}

function sized(st: Style, props: RunProps): number {
  return props.size * SCALE[st.level];
}

function smaller(st: Style, cramped = st.cramped, levels = 1): Style {
  const level = Math.min(2, st.level + levels);
  return { display: false, level, cramped, size: (st.size / SCALE[st.level]) * SCALE[level] };
}

/** Boxes placed together; the result's width is `w` when given, else where the last part ends. */
function assemble(parts: Part[], em: number, cls: Cls, w?: number): MathBox {
  const out: MathBox = { w: w ?? 0, asc: 0, desc: 0, top: 0, bottom: 0, ops: [], em, cls };
  for (const { b, x, y } of parts) {
    out.ops.push(...translateOps(b.ops, x, y));
    if (w === undefined) out.w = Math.max(out.w, x + b.w);
    out.asc = Math.max(out.asc, b.asc - y);
    out.desc = Math.max(out.desc, b.desc + y);
    out.top = Math.max(out.top, b.top - y);
    out.bottom = Math.max(out.bottom, b.bottom + y);
  }
  return out;
}

/** Drawing that spans y0..y1 (y down) and takes no width of its own. */
function ink(ops: DrawOp[], y0: number, y1: number, em: number): MathBox {
  return { w: 0, asc: -y0, desc: y1, top: -y0, bottom: y1, ops, em, cls: 'ord' };
}

function rule(x: number, y: number, w: number, h: number, color: string, em: number): MathBox {
  return ink([{ kind: 'rect', x, y, w, h, fill: color }], y, y + h, em);
}

function layout(n: MathNode, st: Style, fonts: FontRegistry): MathBox {
  switch (n.k) {
    case 'text':
    case 'row':
      return row(n.k === 'row' ? n.items : [n], st, fonts);
    case 'frac':
      return fraction(n, st, fonts);
    case 'script':
      return scripts(n, st, fonts);
    case 'rad':
      return radical(n, st, fonts);
    case 'nary':
      return nary(n, st, fonts);
    case 'delim':
      return delimited(n, st, fonts);
    case 'func': {
      const name = layout(n.name, st, fonts);
      const e = layout(n.e, st, fonts);
      // A thin space parts a function's name from an argument that has no brackets of its own.
      const gap = firstIsDelim(n.e) ? 0 : (name.em * 3) / 18;
      return assemble([{ b: name, x: 0, y: 0 }, { b: e, x: name.w + gap, y: 0 }], name.em, 'ord');
    }
    case 'limit':
      return limit(n, st, fonts);
    case 'accent':
      return accent(n, st, fonts);
    case 'bar': {
      const e = layout(n.e, st, fonts);
      const em = sized(st, n.props);
      const t = Math.max(0.4, RULE * em);
      const y = n.over ? -(Math.max(e.asc, ASC * em) - 0.06 * em) - t : e.desc + 0.02 * em;
      return assemble([{ b: e, x: 0, y: 0 }, { b: rule(0, y, e.w, t, n.props.color ?? '000000', em), x: 0, y: 0 }], e.em, e.cls, e.w);
    }
    case 'group':
      return group(n, st, fonts);
    case 'matrix':
      return matrix(n, st, fonts);
    case 'box': {
      const e = layout(n.e, st, fonts);
      if (!n.border) return e;
      const em = sized(st, n.props);
      const pad = 0.1 * em;
      const t = Math.max(0.4, RULE * em);
      const x1 = e.w + 2 * pad, y0 = -e.asc - pad, y1 = e.desc + pad;
      const frame: DrawOp = { kind: 'path', segs: [{ op: 'M', pts: [0, y0] }, { op: 'L', pts: [x1, y0] }, { op: 'L', pts: [x1, y1] }, { op: 'L', pts: [0, y1] }, { op: 'Z', pts: [] }], fill: null, stroke: n.props.color ?? '000000', strokeWidth: t };
      return assemble([{ b: e, x: pad, y: 0 }, { b: ink([frame], y0 - t / 2, y1 + t / 2, em), x: 0, y: 0 }], e.em, 'ord', x1);
    }
    case 'phantom': {
      const e = layout(n.e, st, fonts);
      return { ...e, w: n.zeroW ? 0 : e.w, asc: n.zeroAsc ? 0 : e.asc, desc: n.zeroDesc ? 0 : e.desc, top: n.zeroAsc || !n.show ? 0 : e.top, bottom: n.zeroDesc || !n.show ? 0 : e.bottom, ops: n.show ? e.ops : [] };
    }
  }
}

function firstIsDelim(n: MathNode): boolean {
  if (n.k === 'delim') return true;
  if (n.k === 'row') return n.items.length > 0 && firstIsDelim(n.items[0]);
  if (n.k === 'text') return OPEN.has([...n.text][0] ?? '');
  return false;
}

// Cambria Math's advances in ems, read from Word's PDFs, for the characters where its stand-in
// (Caladea, metric-compatible with Cambria but not with Cambria Math) is narrower: operators and
// math-italic letters. Unlisted italic letters widen by the listed letters' typical ratio.
const CAMBRIA_MATH_UPRIGHT: Record<string, number> = {
  '=': 0.747, '+': 0.747, '−': 0.747, '±': 0.747, '∓': 0.747, '÷': 0.747, '≠': 0.747,
  '<': 0.749, '>': 0.749, '≤': 0.749, '≥': 0.749, '≈': 0.736, '×': 0.714,
  '→': 0.838, '←': 0.838, '(': 0.415, ')': 0.415, '/': 0.609,
};
const CAMBRIA_MATH_ITALIC: Record<string, number> = {
  a: 0.557, b: 0.539, c: 0.46, d: 0.58, e: 0.496, f: 0.55, h: 0.555, i: 0.317, j: 0.359, k: 0.552, l: 0.316, m: 0.838,
  n: 0.574, o: 0.533, p: 0.558, q: 0.539, r: 0.476, s: 0.462, t: 0.395, u: 0.574, v: 0.544, x: 0.532, y: 0.56, z: 0.483,
  A: 0.633, B: 0.656, D: 0.692, E: 0.625, F: 0.611, H: 0.728, I: 0.379, K: 0.698, L: 0.532, P: 0.62, R: 0.637,
  T: 0.595, U: 0.689, V: 0.625, X: 0.645,
};
const CAMBRIA_MATH_ITALIC_RATIO = 1.1;

/** Cambria Math's advance for a character drawn from a stand-in, in ems (undefined: keep the stand-in's). */
function cambriaMathAdvance(ch: string, italic: boolean): number | undefined {
  return italic ? CAMBRIA_MATH_ITALIC[ch] : CAMBRIA_MATH_UPRIGHT[ch];
}

/** A glyph's outline bounds in font units (TrueType outlines only). */
function glyphBounds(font: SfntFont, gid: number): [number, number, number, number] | null {
  try {
    const range = font.glyphRange(gid);
    if (!range || range[1] - range[0] < 10) return null;
    const o = range[0];
    return [font.r.i16(o + 2), font.r.i16(o + 4), font.r.i16(o + 6), font.r.i16(o + 8)];
  } catch {
    return null;
  }
}

/** Text set at `size` from x = 0 on the baseline, with how far its ink reaches up and down. */
function glyphs(text: string, props: RunProps, size: number, fonts: FontRegistry, italic: boolean, plain = false): { ops: GlyphOp[]; w: number; top: number; bottom: number } {
  const runs = shapeText(text, { ...props, size, italic, vertAlign: 'baseline', position: 0, underline: null, strike: false, dstrike: false }, fonts);
  const ops: GlyphOp[] = [];
  let x = 0;
  let top = 0;
  let bottom = 0;
  for (const r of runs) {
    const font = r.face.face.font;
    const op: GlyphOp = { kind: 'glyphs', x, y: 0, face: r.face.face, size: r.size, gids: [], xs: [], texts: [], color: props.color ?? '000000', syntheticBold: r.face.syntheticBold, syntheticItalic: r.face.syntheticItalic, scaleX: r.scaleX };
    // Word sets an equation in Cambria Math; drawn from a stand-in, each character keeps Cambria
    // Math's advance with its glyph centred in it, so the equation spans what Word's does.
    const standIn = !plain && normalizeFamily(props.fontAscii ?? '') === 'cambria math' && normalizeFamily(r.face.face.familyName) !== 'cambria math';
    let gx = 0;
    for (const g of r.glyphs) {
      let advance = g.advance;
      if (standIn && [...g.text].length === 1) {
        const em = cambriaMathAdvance(g.text, italic);
        if (em !== undefined) advance = em * r.size * r.scaleX;
        else if (italic && isMathItalic(g.text) && /[A-Za-z]/.test(g.text)) advance = g.advance * CAMBRIA_MATH_ITALIC_RATIO;
      }
      if (g.gid >= 0) {
        op.gids.push(g.gid);
        op.xs.push(gx + (advance - g.advance) / 2);
        op.texts.push(g.text);
        const b = glyphBounds(font, g.gid);
        if (b && b[3] > b[1]) {
          top = Math.max(top, (b[3] / font.unitsPerEm) * r.size);
          bottom = Math.max(bottom, (-b[1] / font.unitsPerEm) * r.size);
        } else if (!b && g.text.trim()) {
          top = Math.max(top, ASC * r.size);
          bottom = Math.max(bottom, DESC * r.size);
        }
      }
      gx += advance;
    }
    x += gx;
    if (op.gids.length) ops.push(op);
  }
  return { ops, w: x, top, bottom };
}

/** Text as a box with Cambria Math's line height. */
function textBox(text: string, props: RunProps, em: number, fonts: FontRegistry, italic: boolean, cls: Cls, plain = false): MathBox {
  const g = glyphs(text, props, em, fonts, italic, plain);
  return { w: g.w, asc: ASC * em, desc: DESC * em, top: g.top, bottom: g.bottom, ops: g.ops, em, cls };
}

function classOf(ch: string): Cls {
  if (/^\s$/.test(ch)) return 'space';
  if (BIN.has(ch)) return 'bin';
  if (REL.has(ch)) return 'rel';
  if (OPEN.has(ch)) return 'open';
  if (CLOSE.has(ch)) return 'close';
  if (PUNCT.has(ch)) return 'punct';
  return 'ord';
}

/** A run of math text split into atoms: operators alone, letters and digits together. */
function textAtoms(n: Extract<MathNode, { k: 'text' }>, st: Style, fonts: FontRegistry): MathBox[] {
  const em = sized(st, n.props);
  if (n.plain) return [textBox(n.text, n.props, em, fonts, n.props.italic, 'ord', true)];
  const out: MathBox[] = [];
  let buf = '';
  let bufItalic = false;
  const flush = () => {
    if (buf) out.push(textBox(buf, n.props, em, fonts, bufItalic, 'ord'));
    buf = '';
  };
  for (const raw of n.text) {
    const ch = mathChar(raw);
    if (ch === '&') {
      // An alignment mark: it draws nothing and leaves its neighbours spaced as if it were not there.
      flush();
      out.push({ w: 0, asc: 0, desc: 0, top: 0, bottom: 0, ops: [], em, cls: 'align' });
      continue;
    }
    const cls = classOf(ch);
    if (cls !== 'ord') {
      flush();
      out.push(textBox(ch, n.props, em, fonts, false, cls));
      continue;
    }
    const italic = !n.upright && isMathItalic(ch);
    if (buf && italic !== bufItalic) flush();
    buf += ch;
    bufItalic = italic;
  }
  flush();
  return out;
}

/** Space between neighbouring atoms, in eighteenths of an em (TeX's table, as Word spaces them). */
function spacing(a: Cls, b: Cls, script: boolean): number {
  // A space typed into the equation stands in for the space Word would add.
  if (a === 'space' || b === 'space') return 0;
  if (a === 'op' && (b === 'ord' || b === 'op')) return 3;
  if (b === 'op' && (a === 'ord' || a === 'close')) return 3;
  if (script) return 0;
  if (a === 'bin' || b === 'bin') return 4;
  if (a === 'rel' || b === 'rel') return a === b ? 0 : 5;
  if (a === 'punct') return 3;
  return 0;
}

function row(items: MathNode[], st: Style, fonts: FontRegistry): MathBox {
  const boxes: MathBox[] = [];
  for (const it of items) {
    if (it.k === 'text') boxes.push(...textAtoms(it, st, fonts));
    else if (it.k === 'row') {
      if (it.items.length) boxes.push(row(it.items, st, fonts));
    } else boxes.push(layout(it, st, fonts));
  }
  // An array row aligns at its first mark; the others draw nothing.
  let alignAt = boxes.findIndex((b) => b.cls === 'align');
  const marked = alignAt >= 0;
  const nonMarks = boxes.filter((b) => b.cls !== 'align');
  if (marked) alignAt = boxes.slice(0, alignAt).filter((b) => b.cls !== 'align').length;
  boxes.splice(0, boxes.length, ...nonMarks);
  if (!boxes.length) return { w: 0, asc: 0, desc: 0, top: 0, bottom: 0, ops: [], em: st.size, cls: 'ord', alignX: marked ? 0 : undefined };
  if (boxes.length === 1 && !marked) return boxes[0];
  // A binary operator with nothing to combine on one side is a sign, spaced like a letter.
  for (let i = 0; i < boxes.length; i++) {
    if (boxes[i].cls !== 'bin') continue;
    const prev = boxes[i - 1]?.cls;
    const next = boxes[i + 1]?.cls;
    if (!prev || prev === 'bin' || prev === 'rel' || prev === 'open' || prev === 'punct' || prev === 'op' || !next || next === 'rel' || next === 'close' || next === 'punct') boxes[i] = { ...boxes[i], cls: 'ord' };
  }
  const parts: Part[] = [];
  let x = 0;
  let alignX: number | undefined;
  boxes.forEach((b, i) => {
    if (i > 0) x += (spacing(boxes[i - 1].cls, b.cls, st.level > 0) * Math.max(b.em, boxes[i - 1].em)) / 18;
    if (i === alignAt) alignX = x;
    parts.push({ b, x, y: 0 });
    x += b.w;
  });
  const out = assemble(parts, Math.max(...boxes.map((b) => b.em)), 'ord', x);
  if (marked) out.alignX = alignX ?? x;
  return out;
}

function fraction(n: Extract<MathNode, { k: 'frac' }>, st: Style, fonts: FontRegistry): MathBox {
  const em = sized(st, n.props);
  const color = n.props.color ?? '000000';
  if (n.type === 'lin') {
    const num = layout(n.num, st, fonts);
    const den = layout(n.den, { ...st, cramped: true }, fonts);
    const slash = textBox('/', n.props, em, fonts, false, 'ord');
    return assemble([{ b: num, x: 0, y: 0 }, { b: slash, x: num.w, y: 0 }, { b: den, x: num.w + slash.w, y: 0 }], em, 'ord');
  }
  if (n.type === 'skw') {
    // Sized like a stacked fraction's parts. Measured in display style: the numerator's baseline
    // 0.70em up and the denominator's 0.48em down, joined by a slash 0.36em wide from the
    // denominator's foot to the numerator's top.
    const inner: Style = st.display ? { ...st, display: false } : smaller(st);
    const num = layout(n.num, inner, fonts);
    const den = layout(n.den, { ...inner, cramped: true }, fonts);
    const k = sized(inner, n.props);
    const up = 0.7 * k;
    const down = 0.48 * k;
    const y0 = -up - Math.max(num.top, 0.45 * k);
    const y1 = down + Math.max(den.bottom, 0.1 * k);
    const slash: DrawOp = { kind: 'line', x1: num.w, y1, x2: num.w + 0.36 * k, y2: y0, width: Math.max(0.5, 0.07 * k), color: n.props.color ?? '000000' };
    return assemble([{ b: num, x: 0, y: -up }, { b: ink([slash], y0, y1, em), x: 0, y: 0 }, { b: den, x: num.w + 0.46 * k, y: down }], em, 'ord');
  }
  // Display style keeps a fraction's parts at full size; in running text they shrink to script size.
  const inner: Style = st.display ? { ...st, display: false } : smaller(st);
  const num = layout(n.num, inner, fonts);
  const den = layout(n.den, { ...inner, cramped: true }, fonts);
  const t = n.type === 'bar' ? Math.max(0.4, RULE * em) : 0;
  const gap = (st.display ? 0.22 : 0.1) * em;
  const axis = AXIS * em;
  // Measured at 12pt in display style: numerator baseline 0.76em up, denominator 0.66em down.
  const up = Math.max(st.display ? 0.76 * em : 0, axis + t / 2 + gap + num.desc);
  const down = Math.max(st.display ? 0.66 * em : 0, den.asc + gap + t / 2 - axis);
  const w = Math.max(num.w, den.w);
  const parts: Part[] = [{ b: num, x: (w - num.w) / 2, y: -up }, { b: den, x: (w - den.w) / 2, y: down }];
  if (t) parts.push({ b: rule(0, -axis - t / 2, w, t, color, em), x: 0, y: 0 });
  return assemble(parts, em, 'ord', w);
}

function scripts(n: Extract<MathNode, { k: 'script' }>, st: Style, fonts: FontRegistry): MathBox {
  const base = layout(n.base, st, fonts);
  const em = base.em || st.size;
  const sup = n.sup && layout(n.sup, smaller(st), fonts);
  const sb = n.sub && layout(n.sub, smaller(st, true), fonts);
  // Measured at 12pt: a superscript rises 0.36em (0.30em cramped), a subscript drops 0.20em, and
  // each moves a further 0.04em when both are present. Tall bases push them apart.
  const both = sup && sb ? 0.04 * em : 0;
  const up = Math.max((st.cramped ? 0.3 : 0.36) * em, base.asc - 0.34 * em) + both;
  const down = Math.max(0.2 * em, base.desc - 0.15 * em) + both;
  const space = 0.04 * em;
  const sw = Math.max(sup?.w ?? 0, sb?.w ?? 0) + space;
  const parts: Part[] = [{ b: base, x: n.pre ? sw : 0, y: 0 }];
  if (sup) parts.push({ b: sup, x: n.pre ? sw - space - sup.w : base.w, y: -up });
  if (sb) parts.push({ b: sb, x: n.pre ? sw - space - sb.w : base.w, y: down });
  return assemble(parts, em, base.cls === 'op' || base.cls === 'close' ? base.cls : 'ord', base.w + sw);
}

function radical(n: Extract<MathNode, { k: 'rad' }>, st: Style, fonts: FontRegistry): MathBox {
  const e = layout(n.e, { ...st, cramped: true }, fonts);
  const em = sized(st, n.props);
  const color = n.props.color ?? '000000';
  const t = Math.max(0.4, RULE * em);
  // The bar clears the radicand by 0.08em; the sign dips 0.12em below the baseline (Word's √x).
  const top = Math.max(e.asc, ASC * em) + 0.08 * em + t;
  const bottom = Math.max(0.12 * em, e.desc - 0.1 * em);
  const h = top + bottom;
  const sw = 0.66 * em + Math.max(0, h - 1.5 * em) * 0.15;
  const parts: Part[] = [];
  let kern = 0;
  if (n.deg) {
    const d = layout(n.deg, smaller(st, false, 2), fonts);
    kern = Math.max(0, d.w - 0.35 * sw);
    parts.push({ b: d, x: kern + 0.32 * sw - d.w, y: bottom - 0.6 * h });
  }
  const yT = -top + t / 2;
  const hook = bottom - Math.min(0.45 * h, 0.42 * em);
  const ax = kern + 0.04 * em, bx = kern + 0.17 * em, cx = kern + 0.33 * em, dx = kern + sw;
  const line = (x1: number, y1: number, x2: number, y2: number, width: number): DrawOp => ({ kind: 'line', x1, y1, x2, y2, width, color, cap: 'round' });
  const sign = [line(ax, hook, bx, hook - 0.05 * em, t * 0.9), line(bx, hook - 0.05 * em, cx, bottom, t * 1.7), line(cx, bottom, dx, yT, t * 0.9)];
  parts.push({ b: ink(sign, -top, bottom, em), x: 0, y: 0 }, { b: rule(dx, -top, e.w + 0.06 * em, t, color, em), x: 0, y: 0 }, { b: e, x: dx, y: 0 });
  const out = assemble(parts, em, 'ord', dx + e.w + 0.06 * em);
  return { ...out, asc: Math.max(out.asc, top), desc: Math.max(out.desc, bottom) };
}

/**
 * A character drawn `height` tall (its ink) and centred at `centre` (y down), its width scaled to
 * `aspect` times its height when given.
 */
function bigGlyph(ch: string, props: RunProps, em: number, fonts: FontRegistry, height: number, centre: number, aspect?: number): MathBox {
  const g = glyphs(ch, props, em, fonts, false);
  const op = g.ops[0];
  const font = op?.face.font;
  const b = font && op.gids.length === 1 ? glyphBounds(font, op.gids[0]) : null;
  const upm = font?.unitsPerEm ?? 1000;
  const known = b && b[3] > b[1];
  const inkH = known ? ((b[3] - b[1]) / upm) * em : 0.75 * em;
  const inkMid = known ? ((b[3] + b[1]) / 2 / upm) * em : 0.35 * em;
  const k = height / inkH;
  const sx = aspect && known && b[2] > b[0] ? Math.min(1.4, Math.max(0.6, (aspect * inkH) / (((b[2] - b[0]) / upm) * em))) : 1;
  const y = centre + inkMid * k;
  const ops: DrawOp[] = g.ops.map((o) => ({ ...o, y, size: o.size * k, scaleX: o.scaleX * sx, xs: o.xs.map((v) => v * k * sx) }));
  const top = -(centre - height / 2);
  const bottom = centre + height / 2;
  return { w: g.w * k * sx, asc: top, desc: bottom, top, bottom, ops, em, cls: 'op' };
}

function nary(n: Extract<MathNode, { k: 'nary' }>, st: Style, fonts: FontRegistry): MathBox {
  const em = sized(st, n.props);
  const integral = INTEGRALS.has(n.chr);
  const big = st.display && st.level === 0;
  // Ink heights measured from Word's ∑ and ∫ at 12pt (display 1.9em and 2.3em).
  // Without a character, the limits sit just above and below the baseline (measured).
  const height = n.chr ? (integral ? (big ? 2.3 : 1.5) : big ? 1.9 : 1.3) * em : 0;
  const centre = n.chr ? -(AXIS - (integral ? 0.08 : 0)) * em : 0;
  const op = n.chr ? bigGlyph(n.chr, n.props, em, fonts, height, centre, integral ? 0.32 : SUMS.has(n.chr) ? 0.645 : undefined) : ink([], 0, 0, em);
  const opTop = centre - height / 2;
  const opBottom = centre + height / 2;
  const sup = n.sup && layout(n.sup, smaller(st), fonts);
  const sb = n.sub && layout(n.sub, smaller(st, true), fonts);
  const parts: Part[] = [];
  let w: number;
  if (n.under) {
    w = Math.max(op.w, sup?.w ?? 0, sb?.w ?? 0);
    parts.push({ b: op, x: (w - op.w) / 2, y: 0 });
    if (sup) parts.push({ b: sup, x: (w - sup.w) / 2, y: opTop - 0.12 * em - sup.desc });
    if (sb) parts.push({ b: sb, x: (w - sb.w) / 2, y: opBottom + 0.12 * em + sb.asc });
    w += (3 * em) / 18;
  } else {
    // Limits beside an integral tuck under its slant.
    const tuck = integral ? (0.37 * height) / 2.3 : 0;
    parts.push({ b: op, x: 0, y: 0 });
    if (sup) parts.push({ b: sup, x: op.w, y: opTop + 0.09 * height });
    if (sb) parts.push({ b: sb, x: op.w - tuck, y: opBottom });
    w = Math.max(op.w, op.w + (sup?.w ?? 0), op.w - tuck + (sb?.w ?? 0)) + 0.04 * em;
  }
  parts.push({ b: layout(n.e, st, fonts), x: w, y: 0 });
  return assemble(parts, em, 'op');
}

function delimiter(ch: string, props: RunProps, em: number, fonts: FontRegistry, half: number): MathBox {
  const color = props.color ?? '000000';
  const height = 2 * half;
  // Text-height content keeps the font's own bracket.
  if (height <= 1.05 * em) return textBox(ch, props, em, fonts, false, 'ord');
  const top = -AXIS * em - half;
  const bottom = -AXIS * em + half;
  const t = Math.max(0.5, 0.055 * em);
  const line = (x1: number, y1: number, x2: number, y2: number): DrawOp => ({ kind: 'line', x1, y1, x2, y2, width: t, color });
  const g = glyphs(ch, props, em, fonts, false);
  // Straight brackets and bars are drawn to any height; curved ones stretch the font's glyph.
  switch (ch) {
    case '[':
    case ']':
    case '⌈':
    case '⌉':
    case '⌊':
    case '⌋': {
      const w = Math.max(g.w, 0.35 * em);
      const left = ch === '[' || ch === '⌈' || ch === '⌊';
      const x = left ? 0.12 * em : w - 0.12 * em;
      const end = left ? x + 0.2 * em : x - 0.2 * em;
      const ops = [line(x, top, x, bottom)];
      if (ch !== '⌊' && ch !== '⌋') ops.push(line(x, top + t / 2, end, top + t / 2));
      if (ch !== '⌈' && ch !== '⌉') ops.push(line(x, bottom - t / 2, end, bottom - t / 2));
      return { ...ink(ops, top, bottom, em), w };
    }
    case '|':
    case '│':
    case '∣':
      return { ...ink([line(g.w / 2, top, g.w / 2, bottom)], top, bottom, em), w: g.w };
    case '‖':
    case '∥':
      return { ...ink([line(g.w / 2 - 0.08 * em, top, g.w / 2 - 0.08 * em, bottom), line(g.w / 2 + 0.08 * em, top, g.w / 2 + 0.08 * em, bottom)], top, bottom, em), w: g.w };
  }
  // Curved brackets stretch vertically about the axis from their text-height form.
  const natural = bigGlyph(ch, props, em, fonts, 0.95 * em, -AXIS * em);
  const sy = height / (0.95 * em);
  const cy = -AXIS * em;
  return { ...ink([{ kind: 'transform', matrix: [1, 0, 0, sy, 0, cy - sy * cy] }, ...natural.ops, { kind: 'restore' }], top, bottom, em), w: natural.w };
}

function delimited(n: Extract<MathNode, { k: 'delim' }>, st: Style, fonts: FontRegistry): MathBox {
  const em = sized(st, n.props);
  const items = n.items.map((it) => layout(it, st, fonts));
  let top = 0, bottom = 0;
  for (const it of items) {
    top = Math.max(top, it.top);
    bottom = Math.max(bottom, it.bottom);
  }
  // Brackets cover 90% of the ink they enclose about the axis (TeX's delimiter factor; a 2x2
  // matrix's brackets in Word's PDFs).
  const half = 0.9 * Math.max(top - AXIS * em, bottom + AXIS * em);
  const boxes: MathBox[] = [];
  if (n.beg) boxes.push(delimiter(n.beg, n.props, em, fonts, half));
  items.forEach((it, i) => {
    if (i > 0 && n.sep) boxes.push(delimiter(n.sep, n.props, em, fonts, half));
    boxes.push(it);
  });
  if (n.end) boxes.push(delimiter(n.end, n.props, em, fonts, half));
  const parts: Part[] = [];
  let x = 0;
  for (const b of boxes) {
    parts.push({ b, x, y: 0 });
    x += b.w;
  }
  return assemble(parts, em, 'ord', x);
}

function limit(n: Extract<MathNode, { k: 'limit' }>, st: Style, fonts: FontRegistry): MathBox {
  const base = layout(n.e, st, fonts);
  const em = base.em || st.size;
  const lim = layout(n.lim, smaller(st, !n.over), fonts);
  const w = Math.max(base.w, lim.w);
  // Word tucks a lower limit up into the base's depth (lim's limit sits 0.6em below at 12pt).
  const y = n.over ? -(base.asc + 0.05 * em + lim.desc) : Math.max(0, base.desc - 0.11 * em) + lim.asc;
  return assemble([{ b: base, x: (w - base.w) / 2, y: 0 }, { b: lim, x: (w - lim.w) / 2, y }], em, 'op', w);
}

function accent(n: Extract<MathNode, { k: 'accent' }>, st: Style, fonts: FontRegistry): MathBox {
  const base = layout(n.e, { ...st, cramped: true }, fonts);
  const em = sized(st, n.props);
  const mark = textBox(ACCENT_FORMS[n.chr] ?? n.chr, n.props, em, fonts, false, 'ord');
  // Spacing accents sit at the height of lower-case letters; anything taller lifts them.
  const short = n.e.k === 'text' && X_HEIGHT.test(n.e.text);
  const lift = short ? 0 : Math.max(0.23 * em, base.asc - ASC * em + 0.23 * em);
  const italic = n.e.k === 'text' && !n.e.upright && [...n.e.text].some(isMathItalic);
  const x = (base.w - mark.w) / 2 + (italic ? 0.08 * em : 0);
  const out = assemble([{ b: base, x: 0, y: 0 }, { b: { ...mark, asc: 0, desc: 0 }, x, y: -lift }], base.em, base.cls, Math.max(base.w, mark.w));
  return { ...out, asc: Math.max(base.asc, 0.78 * em + lift) };
}

function group(n: Extract<MathNode, { k: 'group' }>, st: Style, fonts: FontRegistry): MathBox {
  const base = layout(n.e, st, fonts);
  const em = sized(st, n.props);
  const w = base.w;
  const brace = n.chr === '⏟' || n.chr === '⏞' || n.chr === '︸' || n.chr === '︷';
  if (!brace) {
    // Arrows and other characters stretch across the base.
    const g = textBox(n.chr, n.props, em, fonts, false, 'ord');
    const sx = g.w > 0 ? Math.max(1, w / g.w) : 1;
    const ops = g.ops.map((o) => (o.kind === 'glyphs' ? { ...o, x: o.x * sx, scaleX: o.scaleX * sx, xs: o.xs.map((v) => v * sx) } : o));
    const y = n.over ? -(base.top + 0.05 * em + g.bottom) : base.bottom + 0.05 * em + g.top;
    return assemble([{ b: base, x: 0, y: 0 }, { b: { ...g, ops, w: g.w * sx }, x: (w - g.w * sx) / 2, y }], base.em, 'ord', w);
  }
  const h = 0.3 * em;
  const t = Math.max(0.5, RULE * em);
  const s = n.over ? -1 : 1;
  const y0 = n.over ? -(base.top + 0.06 * em) : base.bottom + 0.06 * em;
  const mid = y0 + (s * h) / 2;
  const tip = y0 + s * h;
  const r = Math.min(0.2 * em, w / 4);
  const segs: PathSeg[] = [
    { op: 'M', pts: [0, y0] },
    { op: 'C', pts: [0, mid, 0, mid, r, mid] },
    { op: 'L', pts: [w / 2 - r, mid] },
    { op: 'C', pts: [w / 2, mid, w / 2, mid, w / 2, tip] },
    { op: 'C', pts: [w / 2, mid, w / 2, mid, w / 2 + r, mid] },
    { op: 'L', pts: [w - r, mid] },
    { op: 'C', pts: [w, mid, w, mid, w, y0] },
  ];
  const path: DrawOp = { kind: 'path', segs, fill: null, stroke: n.props.color ?? '000000', strokeWidth: t };
  const mark = ink([path], Math.min(y0, tip) - t / 2, Math.max(y0, tip) + t / 2, em);
  return assemble([{ b: base, x: 0, y: 0 }, { b: mark, x: 0, y: 0 }], base.em, 'ord', w);
}

function matrix(n: Extract<MathNode, { k: 'matrix' }>, st: Style, fonts: FontRegistry): MathBox {
  const em = sized(st, n.props);
  // An equation array's rows keep the equation's style; a matrix's cells are set as running text.
  const cells = n.rows.map((r) => r.map((c) => layout(c, n.eqArr ? st : { ...st, display: false }, fonts)));
  const cols = Math.max(0, ...cells.map((r) => r.length));
  const colW = Array.from({ length: cols }, (_, j) => Math.max(0, ...cells.map((r) => r[j]?.w ?? 0)));
  // Rows of an equation array line up at their alignment marks; rows without one centre.
  const marked = cells.map((r) => r[0]).filter((c) => c?.alignX !== undefined);
  const lead = Math.max(0, ...marked.map((c) => c.alignX!));
  if (n.eqArr && marked.length) colW[0] = Math.max(colW[0], lead + Math.max(0, ...marked.map((c) => c.w - c.alignX!)));
  const cellX = (c: MathBox, j: number) => (n.eqArr && c.alignX !== undefined ? lead - c.alignX : (colW[j] - c.w) / 2);
  // Measured on a 2x2 matrix at 12pt: columns 1em apart, rows 0.26em apart.
  const colGap = em;
  const rowGap = 0.26 * em;
  const rows = cells.map((r) => ({ asc: Math.max(ASC * em, ...r.map((c) => c.asc)), desc: Math.max(DESC * em, ...r.map((c) => c.desc)), top: Math.max(0, ...r.map((c) => c.top)), bottom: Math.max(0, ...r.map((c) => c.bottom)) }));
  const baselines: number[] = [];
  let y = 0;
  rows.forEach((r, i) => {
    const prev = rows[i - 1];
    // An equation array spaces its rows by their ink, at least a line apart (measured at 12pt).
    if (prev && n.eqArr) y += Math.max(1.18 * em, prev.bottom + r.top + EQ_ROW_GAP * em);
    else y += (prev ? prev.desc + rowGap : 0) + r.asc;
    baselines.push(y);
  });
  const total = y + (rows.length ? rows[rows.length - 1].desc : 0);
  // The matrix's ink centres on the axis.
  const inkTop = baselines.length ? baselines[0] - Math.max(0, ...cells[0].map((c) => c.top)) : 0;
  const inkBottom = baselines.length ? baselines[baselines.length - 1] + Math.max(0, ...cells[cells.length - 1].map((c) => c.bottom)) : 0;
  const top = -AXIS * em - (inkTop + inkBottom) / 2;
  const parts: Part[] = [];
  cells.forEach((r, i) => {
    let x = 0;
    r.forEach((c, j) => {
      parts.push({ b: c, x: x + cellX(c, j), y: top + baselines[i] });
      x += colW[j] + colGap;
    });
  });
  const w = colW.reduce((a, b) => a + b, 0) + colGap * Math.max(0, cols - 1);
  const out = assemble(parts, em, 'ord', w);
  return { ...out, asc: Math.max(out.asc, -top), desc: Math.max(out.desc, total + top) };
}
