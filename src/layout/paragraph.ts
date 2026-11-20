import type { FontRegistry } from '../fonts/registry.js';
import { arabicForms, bidiClass, bidiLevels, joinsArabic, mirrored, visualOrder } from './bidi.js';
import type { DocSettings, Drawing, FieldInline, Link, Paragraph, RunProps, TabStop } from '../docx/model.js';
import { shapeSimple, shapeText, type ShapedRun } from './shaper.js';
import { translateOps, type DrawOp, type GlyphOp } from './types.js';
import { layoutMath, mathLineProps, type MathBox } from './math.js';

const EPS = 0.01;
/**
 * Full-width punctuation that Word's punctuation compression squeezes, in equal shares, by up to
 * 0.3 em each when a line overflows: opening brackets lose their blank left half, closing ones
 * their right half, and centred commas and stops some of both.
 */
const OPENING = new Set([...'（「『【〈《〔〖〘［｛'].map((c) => c.codePointAt(0)!));
const CLOSING = new Set([...'）」』】〉》〕〗〙］｝'].map((c) => c.codePointAt(0)!));
const COMPRESSIBLE = new Set([...OPENING, ...CLOSING, ...[...'、。，．：；'].map((c) => c.codePointAt(0)!)]);
const PUNCT_SQUEEZE = 0.3;
/** Full-width punctuation that may not start a line and so hangs past the right edge (w:overflowPunct). */
const HANGING = new Set([...CLOSING, ...[...'、。，．'].map((c) => c.codePointAt(0)!)]);
/**
 * How much of their width the spaces of a justified line may give up so that one more word fits.
 * Word 2013+ (compatibility mode 15) does this; documents in older modes only ever stretch spaces.
 * A quarter matched Word's PDFs best (tried 0.2 to 0.4).
 */
const SPACE_SHRINK = 0.25;
/** Word keeps the word on a justified line only when shrinking the spaces takes less than this
 * share of the stretch (per unit of space width) the line would need without the word (fitted to
 * about 3,500 measured Word line breaks). */
const SQUEEZE_VS_STRETCH = 0.5;
/**
 * WordPerfect justification (compat wpJustification) keeps a word on the line when compressing the
 * line by up to this share of its natural width makes room, taking it out of every character's
 * spacing (measured: Word compresses lines by up to 4.26% and wraps lines that would need 4.5%).
 */
const WP_COMPRESS = 0.0435;

export interface ParagraphEnv {
  fonts: FontRegistry;
  settings: DocSettings;
  fieldText(field: FieldInline): string;
  /** Repeated note stories resolve a field independently for each placement. */
  placedField?(field: FieldInline): FieldInline;
  noteRefText(noteType: 'footnote' | 'endnote', id: string): string;
  /** Number shown by w:footnoteRef inside the note currently being laid out. */
  noteNumberText?: string;
  /** Line pitch when the section uses a "lines" document grid. */
  gridPitch?: number;
  /** Extra width per character when the section uses a character grid. */
  charGrid?: number;
  drawInline(drawing: Drawing, x: number, y: number, w: number, h: number): DrawOp[];
}

type Atom =
  | { k: 'g'; run: ShapedRun; gi: number; w: number; cp: number; field?: FieldInline; grid?: number; aspace?: number }
  | { k: 'tab'; run: ShapedRun; label?: boolean }
  | { k: 'ptab'; run: ShapedRun; alignment: 'left' | 'center' | 'right'; relativeTo: 'margin' | 'indent'; leader: TabStop['leader'] }
  | { k: 'br'; type: 'line' | 'page' | 'column'; run: ShapedRun }
  | { k: 'obj'; drawing: Drawing; w: number; h: number; props: RunProps; link?: Link; run: ShapedRun }
  | { k: 'math'; box: MathBox; run: ShapedRun; align?: 'left' | 'center' | 'right' }
  | { k: 'sep'; w: number; run: ShapedRun; continuation: boolean }
  | { k: 'float'; drawing: Drawing }
  | { k: 'bm'; name: string }
  | { k: 'note'; noteType: 'footnote' | 'endnote'; id: string }
  | { k: 'rpb' };

export interface Line {
  start: number;
  end: number;
  /** x of each atom relative to the layout area's left edge. */
  xs: number[];
  ws: number[];
  left: number;
  right: number;
  /** x where the visible content ends (trailing spaces excluded). */
  contentRight: number;
  ascent: number;
  descent: number;
  height: number;
  /** Baseline offset from the line top. */
  baseline: number;
  brk: 'line' | 'page' | 'column' | null;
  last: boolean;
  first: boolean;
  floats: Drawing[];
  notes: Array<{ noteType: 'footnote' | 'endnote'; id: string }>;
  renderedPageBreak: boolean;
  /** Extra width per expandable gap for justification. */
  justify: { from: number; perGap: number; perChar: number };
  /** Horizontal offset applied for alignment. */
  alignShift: number;
  tabLeaders: Array<{ x: number; w: number; leader: TabStop['leader']; run: ShapedRun }>;
}

function isSpaceCp(cp: number): boolean {
  return cp === 0x20 || cp === 0x3000 || (cp >= 0x2000 && cp <= 0x2006) || cp === 0x2008 || cp === 0x2009 || cp === 0x200a || cp === 0x205f;
}

/** Spaces that may hang past the right margin at a line break. Word wraps en, em and the other
 * fixed-width spaces like text instead (the blank of a FORMTEXT field moves to the next line). */
function hangs(cp: number): boolean {
  return cp === 0x20 || cp === 0x3000;
}

const NO_BREAK_BEFORE = new Set(
  [...'、。，．：；？！）」』】〕〉》’”ー々ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ…‥・゛゜ヽヾゝゞ〟｝］｠｡｣､﹐﹑﹒﹔﹕﹖﹗!%),.:;?]}¢°’”‰℃'].map((c) => c.codePointAt(0)!),
);
const NO_BREAK_AFTER = new Set([...'（「『【〔〈《‘“｛［｟｢([{£¥$'].map((c) => c.codePointAt(0)!));

function isIdeographic(cp: number): boolean {
  return (cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xac00 && cp <= 0xd7af) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef) || (cp >= 0x20000 && cp <= 0x2fa1f);
}

/** Ideographs, kana and hangul: the East Asian text Word spaces apart from Latin text (its
 * full-width punctuation and forms take no such space). */
function isEastAsianText(cp: number): boolean {
  return (
    (cp >= 0x2e80 && cp <= 0x2fdf) ||
    (cp >= 0x3040 && cp <= 0x30ff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0x31f0 && cp <= 0x9fff) ||
    (cp >= 0xac00 && cp <= 0xd7af) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0x20000 && cp <= 0x2fa1f)
  );
}

function isLatinLetter(cp: number): boolean {
  return (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || (cp >= 0xc0 && cp <= 0x24f && cp !== 0xd7 && cp !== 0xf7) || (cp >= 0x370 && cp <= 0x4ff);
}

function isAlnum(cp: number): boolean {
  return (cp >= 0x30 && cp <= 0x39) || (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || (cp >= 0xc0 && cp <= 0x24f) || (cp >= 0x370 && cp <= 0x52f);
}

export class ParagraphLayout {
  readonly atoms: Atom[] = [];
  private breakAfter: boolean[] = [];
  /** Bidi embedding level per atom, when the paragraph has right-to-left text. */
  private levels: number[] | null = null;
  pos = 0;
  private firstLine = true;
  readonly markRun: ShapedRun;
  /** Width the numbering label is shifted left by (lvlJc right/center). */
  private labelShift = 0;
  private labelEnd = 0;
  /** Atoms before the label: page or column breaks opening the paragraph. */
  private labelStart = 0;
  private hasLabel = false;
  readonly tabs: TabStop[];

  constructor(
    readonly para: Paragraph,
    readonly env: ParagraphEnv,
  ) {
    const fonts = env.fonts;
    this.markRun = shapeSimple(' ', para.markProps, fonts)[0] ?? emptyRun(para.markProps, fonts);
    this.tabs = para.props.tabs.filter((t) => t.type !== 'bar');
    if (para.numbering) {
      const lab = para.numbering;
      this.hasLabel = true;
      if (lab.picture) {
        const h = lab.props.size * 0.8;
        const run = shapeSimple(' ', lab.props, fonts)[0] ?? this.markRun;
        this.atoms.push({ k: 'obj', drawing: pictureBullet(lab.picture, h), w: h, h, props: lab.props, run });
      } else {
        for (const run of shapeText(lab.text, lab.props, fonts)) for (let gi = 0; gi < run.glyphs.length; gi++) this.pushGlyph(run, gi);
      }
      let labelWidth = 0;
      for (const a of this.atoms) labelWidth += atomWidth(a);
      if (lab.align === 'right') this.labelShift = labelWidth;
      else if (lab.align === 'center') this.labelShift = labelWidth / 2;
      if (lab.suffix === 'tab') this.atoms.push({ k: 'tab', run: shapeSimple(' ', lab.props, fonts)[0] ?? this.markRun, label: true });
      else if (lab.suffix === 'space') for (const run of shapeText(' ', lab.props, fonts)) for (let gi = 0; gi < run.glyphs.length; gi++) this.pushGlyph(run, gi);
      this.labelEnd = this.atoms.length;
    }
    for (const inl of para.content) {
      switch (inl.type) {
        case 'text':
          for (const run of shapeText(inl.text, inl.props, fonts, inl.link)) for (let gi = 0; gi < run.glyphs.length; gi++) this.pushGlyph(run, gi);
          break;
        case 'tab':
          this.atoms.push({ k: 'tab', run: { ...(shapeSimple(' ', inl.props, fonts)[0] ?? this.markRun), link: inl.link } });
          break;
        case 'ptab':
          this.atoms.push({ k: 'ptab', run: shapeSimple(' ', inl.props, fonts)[0] ?? this.markRun, alignment: inl.alignment, relativeTo: inl.relativeTo, leader: inl.leader });
          break;
        case 'break':
          this.atoms.push({ k: 'br', type: inl.breakType, run: shapeSimple(' ', inl.props, fonts)[0] ?? this.markRun });
          break;
        case 'drawing': {
          const d = inl.drawing;
          if (d.inline) {
            const w = d.width + d.effectL + d.effectR;
            const h = d.height + d.effectT + d.effectB;
            this.atoms.push({ k: 'obj', drawing: d, w, h, props: inl.props, link: inl.link, run: shapeSimple(' ', inl.props, fonts)[0] ?? this.markRun });
          } else this.atoms.push({ k: 'float', drawing: d });
          break;
        }
        case 'field': {
          const field = env.placedField?.(inl) ?? inl;
          const text = env.fieldText(field);
          let first = true;
          for (const run of shapeText(text, inl.props, fonts)) {
            for (let gi = 0; gi < run.glyphs.length; gi++) {
              this.pushGlyph(run, gi, first ? field : undefined);
              first = false;
            }
          }
          break;
        }
        case 'noteRef': {
          if (inl.customMark === undefined) {
            const text = env.noteRefText(inl.noteType, inl.id);
            for (const run of shapeText(text, inl.props, fonts)) for (let gi = 0; gi < run.glyphs.length; gi++) this.pushGlyph(run, gi);
          }
          this.atoms.push({ k: 'note', noteType: inl.noteType, id: inl.id });
          break;
        }
        case 'noteNumber': {
          const text = env.noteNumberText ?? '';
          for (const run of shapeText(text, inl.props, fonts)) for (let gi = 0; gi < run.glyphs.length; gi++) this.pushGlyph(run, gi);
          break;
        }
        case 'separator':
          this.atoms.push({ k: 'sep', w: inl.continuation ? -1 : 144, run: shapeSimple(' ', inl.props, fonts)[0] ?? this.markRun, continuation: inl.continuation });
          break;
        case 'bookmark':
          this.atoms.push({ k: 'bm', name: inl.name });
          break;
        case 'renderedPageBreak':
          this.atoms.push({ k: 'rpb' });
          break;
        case 'math': {
          const box = layoutMath(inl.math, inl.props, inl.display, fonts);
          // The equation's own text size (its runs' w:sz), not the paragraph's, sizes its line.
          const run = shapeSimple(' ', { ...mathLineProps(inl.props), size: box.em || inl.props.size }, fonts)[0] ?? this.markRun;
          this.atoms.push({ k: 'math', box, run, align: inl.display ? inl.align : undefined });
          break;
        }
      }
    }
    if (this.hasLabel) this.moveLabelPastBreaks();
    this.shapeArabic();
    this.autoSpace();
    this.computeBreaks();
    const base = para.props.bidi ? 1 : 0;
    const classes = this.atoms.map((a) => (a.k === 'g' ? bidiClass(a.cp) : 'N'));
    if (base || classes.some((c) => c === 'R' || c === 'AL')) {
      this.levels = bidiLevels(classes, this.atoms.map((a) => a.k === 'g' && a.run.props.rtl), base);
      this.mirrorBrackets(this.levels);
    }
  }

  /** Brackets in right-to-left text face the other way (rule L4): ")" ends what reads as "(". */
  private mirrorBrackets(levels: number[]): void {
    this.atoms.forEach((a, i) => {
      if (a.k !== 'g' || levels[i] % 2 === 0) return;
      const m = mirrored(a.cp);
      const g = a.run.glyphs[a.gi];
      const font = a.run.face.face.font;
      const gid = m === undefined ? 0 : font.glyphId(m);
      if (g.gid < 0 || gid <= 0) return;
      g.advance += (font.advance(gid) - font.advance(g.gid)) * ((a.run.size * a.run.scaleX) / font.unitsPerEm);
      g.gid = gid;
      g.text = String.fromCodePoint(m!);
      a.w = g.advance;
    });
  }

  /** Arabic letters take their joining forms from the letters around them. */
  private shapeArabic(): void {
    const atoms = this.atoms;
    const cps = atoms.map((a) => (a.k === 'g' ? a.cp : -1));
    if (!cps.some(joinsArabic)) return;
    const forms = arabicForms(cps);
    for (let i = 0; i < atoms.length; i++) {
      const a = atoms[i];
      if (a.k !== 'g' || forms[i] === cps[i] || forms[i] === 0) continue;
      const g = a.run.glyphs[a.gi];
      const font = a.run.face.face.font;
      const gid = font.glyphId(forms[i]);
      if (g.gid < 0 || gid <= 0) continue;
      const scale = (a.run.size * a.run.scaleX) / font.unitsPerEm;
      g.advance += (font.advance(gid) - font.advance(g.gid)) * scale;
      g.gid = gid;
      a.w = g.advance;
      // A lam-alef ligature draws the alef too.
      if (forms[i] >= 0xfef5 && forms[i] <= 0xfefc) {
        for (let k = i + 1; k < atoms.length; k++) {
          const b = atoms[k];
          if (b.k !== 'g' || forms[k] !== 0) continue;
          const bg = b.run.glyphs[b.gi];
          g.text += bg.text;
          bg.gid = -1;
          bg.advance = 0;
          b.w = 0;
          break;
        }
      }
    }
  }

  /**
   * Word sets East Asian text a quarter em apart from Latin letters (w:autoSpaceDE) and digits
   * (w:autoSpaceDN), both on unless a paragraph turns them off. The space trails the first of the
   * pair, so like a space it can hang past the margin, and justification stretches it first.
   */
  private autoSpace(): void {
    const { autoSpaceDE, autoSpaceDN } = this.para.props;
    if (!autoSpaceDE && !autoSpaceDN) return;
    const atoms = this.atoms;
    const kind = (cp: number): number => (isEastAsianText(cp) ? 1 : cp >= 0x30 && cp <= 0x39 ? (autoSpaceDN ? 2 : 0) : isLatinLetter(cp) ? (autoSpaceDE ? 2 : 0) : 0);
    for (let i = 0; i < atoms.length; i++) {
      const a = atoms[i];
      if (a.k !== 'g' || (i >= this.labelStart && i < this.labelEnd) || a.run.glyphs[a.gi].gid === -1) continue;
      const ka = kind(a.cp);
      if (!ka) continue;
      let j = i + 1;
      while (j < atoms.length && (atoms[j].k === 'bm' || atoms[j].k === 'rpb' || atoms[j].k === 'note')) j++;
      const b = atoms[j];
      if (!b || b.k !== 'g' || (j >= this.labelStart && j < this.labelEnd) || b.run.glyphs[b.gi].gid === -1) continue;
      const kb = kind(b.cp);
      if (!kb || ka === kb || (ka !== 1 && kb !== 1)) continue;
      const gap = (ka === 1 ? a.run.size : b.run.size) / 4;
      a.run.glyphs[a.gi].advance += gap;
      a.w += gap;
      a.aspace = gap;
    }
  }

  /**
   * A numbered paragraph that opens with a page or column break shows its number after the break,
   * with the text on the new page (Word's PDFs), so the label atoms move behind the break.
   */
  private moveLabelPastBreaks(): void {
    const atoms = this.atoms;
    let lead = this.labelEnd;
    for (let k = this.labelEnd; k < atoms.length; k++) {
      const a = atoms[k];
      if (a.k === 'br' && a.type !== 'line') lead = k + 1;
      else if (a.k !== 'bm' && a.k !== 'rpb') break;
    }
    if (lead === this.labelEnd) return;
    // With nothing after the break, Word shows no number on either page.
    if (atoms.slice(lead).every((a) => a.k === 'bm' || a.k === 'rpb')) {
      atoms.splice(0, this.labelEnd);
      this.labelEnd = 0;
      this.labelShift = 0;
      this.hasLabel = false;
      return;
    }
    const label = atoms.slice(0, this.labelEnd);
    const breaks = atoms.slice(this.labelEnd, lead);
    atoms.splice(0, lead, ...breaks, ...label);
    this.labelStart = breaks.length;
    this.labelEnd = breaks.length + label.length;
  }

  private pushGlyph(run: ShapedRun, gi: number, field?: FieldInline): void {
    const g = run.glyphs[gi];
    // A character grid (w:docGrid linesAndChars) widens every character by the grid's spacing.
    const grid = this.env.charGrid && this.para.props.snapToGrid && run.props.snapToGrid !== false && g.advance > 0 ? this.env.charGrid : 0;
    g.advance += grid;
    this.atoms.push({ k: 'g', run, gi, w: g.advance, cp: g.cp, field, grid: grid || undefined });
  }

  private computeBreaks(): void {
    const atoms = this.atoms;
    const n = atoms.length;
    this.breakAfter = new Array(n).fill(false);
    const cpAt = (i: number): number => {
      const a = atoms[i];
      return a && a.k === 'g' ? a.cp : -1;
    };
    for (let i = 0; i < n; i++) {
      const a = atoms[i];
      if (a.k === 'tab' || a.k === 'ptab' || a.k === 'obj' || a.k === 'math' || a.k === 'br') {
        this.breakAfter[i] = true;
        if ((a.k === 'obj' || a.k === 'math') && i > 0) this.breakAfter[i - 1] = true;
        continue;
      }
      if (a.k !== 'g') continue;
      const cp = a.cp;
      // Look past zero-width atoms for the next character.
      let j = i + 1;
      while (j < n && atoms[j].k !== 'g' && atoms[j].k !== 'tab' && atoms[j].k !== 'obj' && atoms[j].k !== 'math' && atoms[j].k !== 'br' && atoms[j].k !== 'ptab') j++;
      const next = cpAt(j);
      if (isSpaceCp(cp)) {
        // Break after the last space of a run of spaces. A run of fixed-width spaces after ordinary
        // ones wraps as a unit, like a word.
        this.breakAfter[i] = !(next >= 0 && isSpaceCp(next) && (!hangs(cp) || hangs(next)));
      } else if (cp === 0xad || cp === 0x200b) {
        this.breakAfter[i] = true;
      } else if (cp === 0x2d || cp === 0x2010 || cp === 0x2013 || cp === 0x2014) {
        const prev = cpAt(i - 1);
        this.breakAfter[i] = prev >= 0 && !isSpaceCp(prev) && next >= 0 && !isSpaceCp(next) && (cp !== 0x2d || isAlnum(next) || next > 0x7f) && !NO_BREAK_BEFORE.has(next);
      } else if (next >= 0 && (isIdeographic(cp) || isIdeographic(next))) {
        this.breakAfter[i] = !NO_BREAK_AFTER.has(cp) && !NO_BREAK_BEFORE.has(next) && !isSpaceCp(next);
      }
    }
  }

  get done(): boolean {
    return this.pos >= this.atoms.length && !this.firstLine && !this.pendingMarkLine;
  }

  /** Set when the paragraph ended with a line break: the paragraph mark sits on its own line. */
  private pendingMarkLine = false;

  /** Restart line breaking from the beginning (used when the flow backtracks). */
  reset(): void {
    this.pos = 0;
    this.firstLine = true;
    this.pendingMarkLine = false;
  }

  get isFirstLine(): boolean {
    return this.firstLine;
  }

  /** Line-breaking state, so the flow can retry a line with different bounds. */
  /** True when the line holds nothing but a break (and bookmarks). */
  onlyBreak(line: Line): boolean {
    for (let k = line.start; k < line.end; k++) {
      const a = this.atoms[k];
      if (a.k !== 'br' && a.k !== 'bm' && a.k !== 'rpb') return false;
    }
    return true;
  }

  snapshot(): { pos: number; firstLine: boolean; pendingMarkLine: boolean } {
    return { pos: this.pos, firstLine: this.firstLine, pendingMarkLine: this.pendingMarkLine };
  }

  restore(s: { pos: number; firstLine: boolean; pendingMarkLine: boolean }): void {
    this.pos = s.pos;
    this.firstLine = s.firstLine;
    this.pendingMarkLine = s.pendingMarkLine;
  }

  /** Floating objects anchored in this paragraph (positioned before its lines are laid out). */
  get floatDrawings(): Drawing[] {
    const out: Drawing[] = [];
    for (const a of this.atoms) if (a.k === 'float') out.push(a.drawing);
    return out;
  }

  /** Left/right edges of the next line inside an area of `areaWidth`, from the paragraph indents. */
  lineBounds(areaWidth: number): { left: number; right: number } {
    const p = this.para.props;
    const first = this.firstLine ? p.indFirstLine : 0;
    // A right-to-left paragraph's lines start on the right.
    if (p.bidi) return { left: p.indLeft, right: areaWidth - p.indRight - first };
    return { left: p.indLeft + first, right: areaWidth - p.indRight };
  }

  /**
   * Lay out the next line between `left` and `right` (area coordinates). Returns null when the
   * paragraph is complete.
   */
  nextLine(left: number, right: number, areaWidth: number): Line | null {
    const atoms = this.atoms;
    // A horizontal rule given as a share of the column takes it from the text width.
    for (let k = this.pos; k < atoms.length; k++) {
      const a = atoms[k];
      if (a.k === 'obj' && a.drawing.hr && a.drawing.hr.pct > 0) a.w = a.drawing.width = a.drawing.hr.pct * Math.max(0, right - left);
    }
    if (this.pos >= atoms.length && !this.firstLine && !this.pendingMarkLine) return null;
    const start = this.pos;
    const first = this.firstLine;
    this.firstLine = false;
    const xs: number[] = [];
    const ws: number[] = [];
    const leaders: Line['tabLeaders'] = [];
    let x = left - (first && this.hasLabel ? this.labelShift : 0);
    let candidate = -1; // line end (exclusive) at the last break opportunity
    let candidateX = 0;
    let i = start;
    let brk: Line['brk'] = null;
    let hasContent = false;
    let lastTab = -1;
    // Where text must wrap; a right tab past the right indent moves it out to the tab.
    let limit = right;

    const finish = (end: number): number => end;
    let end = -1;
    // Word keeps a word on a justified line when shrinking the line's spaces makes room for it
    // (Word 2013+). WordPerfect justification instead compresses every character a little.
    const wp = this.para.props.align === 'both' && this.env.settings.wpJustification;
    const shrink = this.para.props.align === 'both' && !wp && this.env.settings.compatibilityMode >= 15 ? SPACE_SHRINK : 0;
    const compress = wp ? WP_COMPRESS : 0;
    const lineStart = x;
    let spaceSum = 0;
    // End of the last visible atom and the space width before it, now and at the break candidate.
    let inkX = x;
    let inkSpaces = 0;
    let candInkX = x;
    let candInkSpaces = 0;
    // Room East Asian punctuation can give up when the line would otherwise overflow.
    const punctOn = !!this.env.settings.compressPunctuation;
    let punct = 0;
    let hungAt = -1;
    for (; i < atoms.length; i++) {
      const a = atoms[i];
      if (a.k === 'br') {
        xs.push(x);
        ws.push(0);
        brk = a.type;
        i++;
        end = i;
        break;
      }
      if (a.k === 'tab' || a.k === 'ptab') {
        const t: { width: number; leader: TabStop['leader']; overflow: boolean; reach?: number } = a.k === 'tab' ? this.resolveTab(i, x, first, areaWidth, a.label === true) : this.resolvePTab(i, x, a, areaWidth);
        // Before Word 2013 a line of nothing but tabs keeps them all, past the margin; from Word
        // 2013 an overflowing tab wraps there too, as it does after text.
        const keepTabs = this.env.settings.compatibilityMode < 15 && this.blank(start, i);
        if (t.overflow && hasContent && right - left > 1 && !keepTabs) {
          end = candidate > start ? candidate : i;
          break;
        }
        limit = Math.max(right, t.reach ?? right);
        xs.push(x);
        ws.push(t.width);
        if (t.leader !== 'none' && t.width > 0) leaders.push({ x, w: t.width, leader: t.leader, run: a.run });
        x += t.width;
        lastTab = i;
        spaceSum = 0;
        punct = 0;
        hasContent = true;
        candidate = i + 1;
        candidateX = x;
        inkX = candInkX = x;
        inkSpaces = candInkSpaces = 0;
        continue;
      }
      const w = atomWidth(a);
      if (a.k === 'g' && hangs(a.cp)) {
        xs.push(x);
        ws.push(w);
        x += w;
        if (a.cp === 0x20) spaceSum += w;
        if (this.breakAfter[i]) {
          candidate = i + 1;
          candidateX = x;
          candInkX = inkX;
          candInkSpaces = inkSpaces;
        }
        continue;
      }
      // Shrinking the spaces wins only when it changes them less than half as much as stretching
      // them would if the word wrapped instead.
      let squeeze = shrink * spaceSum;
      if (squeeze > 0 && candidate > start && candInkSpaces > 0) squeeze = Math.min(squeeze, SQUEEZE_VS_STRETCH * ((right - candInkX) / candInkSpaces) * spaceSum);
      // A character's grid spacing and the space after it before Latin text may hang past the margin.
      const tail = a.k === 'g' ? (a.grid ?? 0) + (a.aspace ?? 0) : 0;
      const room = limit + squeeze + compress * (x + w - lineStart) + punct + EPS;
      if (w > 0 && x + w - tail > room && hasContent) {
        // Word lets one closing punctuation mark that starts inside the line hang past its edge
        // rather than wrap it with the character before it.
        if (hungAt < 0 && this.para.props.overflowPunct && a.k === 'g' && HANGING.has(a.cp) && x <= room) hungAt = i;
        else {
          if (candidate > start) end = candidate;
          else end = i; // emergency break inside a word
          break;
        }
      }
      xs.push(x);
      ws.push(w);
      x += w;
      if (punctOn && a.k === 'g' && COMPRESSIBLE.has(a.cp) && (i < this.labelStart || i >= this.labelEnd)) punct += (w - (a.grid ?? 0)) * PUNCT_SQUEEZE;
      if (w > 0 || a.k === 'obj' || a.k === 'math') hasContent = true;
      if (w > 0) {
        inkX = x;
        inkSpaces = spaceSum;
      }
      if (this.breakAfter[i]) {
        candidate = i + 1;
        candidateX = x;
        candInkX = inkX;
        candInkSpaces = inkSpaces;
      }
    }
    if (end < 0) end = finish(i);
    if (end <= start && start < atoms.length) end = start + 1; // always make progress
    xs.length = end - start;
    ws.length = end - start;
    void candidateX;
    void lastTab;

    // A soft hyphen at the break point becomes visible.
    const lastAtom = atoms[end - 1];
    if (lastAtom && lastAtom.k === 'g' && lastAtom.cp === 0xad && end < atoms.length) {
      const run = lastAtom.run;
      const font = run.face.face.font;
      ws[end - 1 - start] = (font.advance(font.glyphId(0x2d)) / font.unitsPerEm) * run.size * run.scaleX;
    }

    this.pos = end;
    // Trailing paragraph-mark line after a manual line break at the very end.
    let isLast = this.pos >= atoms.length;
    if (isLast && brk === 'line') {
      this.pendingMarkLine = true;
      isLast = false;
    } else if (this.pendingMarkLine && start >= atoms.length) {
      this.pendingMarkLine = false;
      isLast = true;
    }
    // Word 2013+: a page/column break that ends the paragraph keeps the mark on the same line.
    if (brk === 'page' || brk === 'column') {
      let rest = end;
      while (rest < atoms.length && (atoms[rest].k === 'bm' || atoms[rest].k === 'rpb')) rest++;
      if (rest >= atoms.length) {
        isLast = true;
        this.pos = atoms.length;
      }
    }

    // Visible extent (exclude trailing spaces). Only ordinary spaces trail: Word centres or
    // right-aligns a line with its trailing en or em spaces.
    let contentRight = left;
    for (let k = end - 1; k >= start; k--) {
      const a = atoms[k];
      if (a.k === 'g' && (a.cp === 0x20 || a.cp === 0x3000)) continue;
      if (a.k === 'br') continue;
      if ((a.k === 'g' || a.k === 'obj' || a.k === 'math' || a.k === 'tab' || a.k === 'ptab' || a.k === 'sep') && ws[k - start] >= 0) {
        contentRight = xs[k - start] + ws[k - start] - (a.k === 'g' ? (a.aspace ?? 0) : 0);
        break;
      }
    }

    // A line whose last mark hangs keeps its natural spacing.
    const hung = hungAt >= start && hungAt < end;
    // An overfull line takes the overflow out of its East Asian punctuation.
    if (punctOn && !hung && contentRight > right + EPS) {
      const amounts: number[] = [];
      let total = 0;
      for (let k = start; k < end; k++) {
        const a = atoms[k];
        const amt = a.k === 'g' && COMPRESSIBLE.has(a.cp) && (k < this.labelStart || k >= this.labelEnd) && xs[k - start] + ws[k - start] <= contentRight ? (ws[k - start] - (a.grid ?? 0)) * PUNCT_SQUEEZE : 0;
        amounts.push(amt);
        total += amt;
      }
      const over = Math.min(contentRight - right, total);
      let shift = 0;
      for (let k = start; total > 0 && k < end; k++) {
        const c = (amounts[k - start] / total) * over;
        const cp = (atoms[k] as { cp: number }).cp;
        // The part of the squeeze taken before the glyph moves it left; the rest narrows it.
        const before = c > 0 ? (OPENING.has(cp) ? c : CLOSING.has(cp) ? 0 : c / 2) : 0;
        shift += before;
        xs[k - start] -= shift;
        ws[k - start] -= c - before;
        shift += c - before;
      }
      contentRight -= over;
    }

    // Metrics
    let ascent = 0;
    let descent = 0;
    let ext = 0;
    let textAscent = 0;
    let textDescent = 0;
    let textTop = 0;
    let measured = false;
    const add = (asc: number, desc: number, lead: number, text = true) => {
      ascent = Math.max(ascent, asc);
      descent = Math.max(descent, desc);
      ext = Math.max(ext, lead);
      if (text) {
        textAscent = Math.max(textAscent, asc);
        textDescent = Math.max(textDescent, desc);
        textTop = Math.max(textTop, asc + lead);
      }
      measured = true;
    };
    const floats: Drawing[] = [];
    const notes: Line['notes'] = [];
    let rpb = false;
    // A numbering label lends the line its ascent but never its descent: Courier New "o" bullets
    // leave Calibri lines exactly as tall as plain ones, while Symbol's taller "•" pushes the
    // baseline down (measured against Word's PDFs).
    let labelAscent = 0;
    let pictures = false;
    for (let k = start; k < end; k++) {
      const a = atoms[k];
      if (k >= this.labelStart && k < this.labelEnd && (a.k === 'g' || a.k === 'tab')) {
        if (a.k === 'tab' || a.run.glyphs[a.gi].gid !== -1) labelAscent = Math.max(labelAscent, a.run.ascent + (a.k === 'g' ? a.run.baselineShift : 0));
        continue;
      }
      switch (a.k) {
        case 'g':
          if (a.run.glyphs[a.gi].gid === -1) break;
          // Spaces and tabs never size a line: Word sizes it by its text, and a line of nothing
          // but spaces and tabs by the paragraph mark.
          if (isSpaceCp(a.cp)) break;
          add(a.run.ascent + a.run.baselineShift, a.run.descent - a.run.baselineShift, a.run.extLeading);
          break;
        case 'tab':
        case 'ptab':
          break;
        case 'br':
        case 'sep':
          // A line break after nothing but pictures leaves the line sized as a line of pictures.
          if (a.k === 'br' && pictures && textAscent + textDescent === 0) break;
          add(a.run.ascent, a.run.descent, a.run.extLeading);
          break;
        case 'obj':
          pictures = true;
          add(a.h + a.props.position, Math.max(0, -a.props.position), 0, false);
          // A horizontal rule's line is as tall as its run's text.
          if (a.drawing.hr) add(a.run.ascent, a.run.descent, a.run.extLeading);
          break;
        case 'math':
          // An equation's font sizes its line like text. Where its ink reaches further, the line
          // grows like a picture, beyond what "multiple" spacing scales: to the ink below the
          // baseline and 0.15em above it (measured against Word's PDFs).
          add(a.run.ascent, a.run.descent, a.run.extLeading);
          add(a.box.top + 0.15 * a.box.em, a.box.bottom, 0, false);
          break;
        case 'float':
          floats.push(a.drawing);
          break;
        case 'note':
          notes.push({ noteType: a.noteType, id: a.id });
          break;
        case 'rpb':
          rpb = true;
          break;
      }
    }
    // The paragraph mark only sizes a line that has no text of its own (an empty paragraph, or a
    // line holding just an inline object). Word ignores a larger mark on a non-empty last line;
    // counting it made 8pt clauses with 16pt heading marks run two pages long.
    const mark = this.markRun;
    if (!measured) add(mark.ascent, mark.descent, mark.extLeading);
    else if (textAscent + textDescent === 0) {
      // A line holding only inline pictures is as tall as the taller of the pictures and the mark's
      // text height: a picture taller than the text leaves no descent below it (measured against
      // Word's PDFs). "Multiple" spacing still adds the mark's extra.
      const top = Math.max(ascent, mark.ascent);
      descent = Math.max(descent, mark.ascent + mark.descent - top);
      ascent = top;
      ext = Math.max(ext, mark.extLeading);
      textAscent = mark.ascent;
      textDescent = mark.descent;
    }
    // A font's line gap sits above its own ascent, so a taller font takes in a shorter one's: Times
    // New Roman beside SimSun or Cambria, or Arial beside 18pt Courier New, leaves the line as tall
    // as the taller font makes it (measured against Word's PDFs).
    if (textTop > 0) ext = Math.min(ext, textTop - textAscent);
    // The label's ascent competes with the text's ascent plus its leading: Symbol's "•" in a Times
    // line adds only what its ascent exceeds Times' ascent and line gap. "Multiple" spacing scales
    // the text's height only, not what a larger label adds.
    ascent = Math.max(ascent, labelAscent - ext);
    const { height, baseline } = lineHeight(ascent, descent, ext, this.para.props, this.env.gridPitch, textAscent + textDescent + ext);

    // Alignment and justification.
    const p = this.para.props;
    const avail = right - left;
    const used = contentRight - left;
    let alignShift = 0;
    const justify = { from: start, perGap: 0, perChar: 0 };
    // A displayed equation keeps its own alignment, whatever the paragraph's.
    let display: 'left' | 'center' | 'right' | undefined;
    for (let k = start; k < end; k++) {
      const a = atoms[k];
      if (a.k === 'math' && a.align) display = a.align;
    }
    const endsWithLineBreak = brk === 'line';
    const justifyThis = (p.align === 'both' && !isLast && !(endsWithLineBreak && this.env.settings.doNotExpandShiftReturn)) || p.align === 'distribute';
    // A justified line that kept its last word by shrinking its spaces stays shrunk, even when it
    // is the paragraph's last line (which otherwise isn't justified).
    if (display) alignShift = display === 'center' ? (avail - used) / 2 : display === 'right' ? avail - used : 0;
    else if ((justifyThis && used < avail) || (p.align === 'both' && !hung && used > avail + EPS)) {
      const extra = avail - used;
      // Only spaces after the last tab stretch.
      let from = start;
      for (let k = start; k < end; k++) if (atoms[k].k === 'tab' || atoms[k].k === 'ptab') from = k + 1;
      if (first && this.hasLabel) from = Math.max(from, this.labelEnd);
      let lastVisible = end - 1;
      while (lastVisible >= from && atoms[lastVisible].k === 'g' && isSpaceCp((atoms[lastVisible] as { cp: number }).cp)) lastVisible--;
      let gaps = 0;
      let chars = 0;
      for (let k = from; k <= lastVisible; k++) {
        const a = atoms[k];
        if (stretches(a, k < lastVisible)) gaps++;
        if (a.k === 'g' || a.k === 'obj') chars++;
      }
      justify.from = from;
      if (extra < 0 && this.env.settings.wpJustification) {
        // WordPerfect justification takes the overflow out of every character's spacing.
        if (chars > 0) justify.perChar = extra / chars;
      } else if (p.align === 'distribute') {
        if (chars > 1) justify.perChar = extra / (chars - 1);
      } else if (gaps > 0) justify.perGap = extra / gaps;
      else {
        // East Asian text justifies between characters.
        let cjk = false;
        for (let k = from; k <= lastVisible; k++) {
          const a = atoms[k];
          if (a.k === 'g' && isIdeographic(a.cp)) cjk = true;
        }
        if (cjk && chars > 1) justify.perChar = extra / (chars - 1);
      }
      (justify as { last?: number }).last = lastVisible;
    } else if (p.align === 'center') alignShift = (avail - used) / 2;
    else if (p.align === 'right') alignShift = avail - used;
    // A horizontal rule sits by its own alignment.
    let hr: Drawing['hr'];
    for (let k = start; k < end; k++) if (atoms[k].k === 'obj') hr = (atoms[k] as { drawing: Drawing }).drawing.hr ?? hr;
    if (hr) {
      justify.perGap = justify.perChar = 0;
      alignShift = hr.align === 'center' ? (avail - used) / 2 : hr.align === 'right' ? avail - used : 0;
    }

    // The breaks before a moved label leave the next line the paragraph's first, with the label.
    const breakLine = first && this.labelStart > 0 && end <= this.labelStart;
    if (breakLine) this.firstLine = true;
    return {
      start,
      end,
      xs,
      ws,
      left,
      right,
      contentRight,
      ascent,
      descent,
      height,
      baseline,
      brk,
      last: isLast,
      first: first && (breakLine || this.labelStart === 0 || start < this.labelStart),
      floats,
      notes,
      renderedPageBreak: rpb,
      justify,
      alignShift,
      tabLeaders: leaders,
    };
  }

  /**
   * Widest unbreakable stretch of the paragraph (its spaces hang, so they don't count): the
   * narrowest an autofit table column can make it without breaking a word.
   */
  minContentWidth(): number {
    let best = 0;
    let seg = 0;
    for (let i = 0; i < this.atoms.length; i++) {
      const a = this.atoms[i];
      if (a.k === 'br' || a.k === 'tab' || a.k === 'ptab') {
        best = Math.max(best, seg);
        seg = 0;
        continue;
      }
      if (!(a.k === 'g' && hangs(a.cp))) seg += atomWidth(a);
      if (this.breakAfter[i]) {
        best = Math.max(best, seg);
        seg = 0;
      }
    }
    const p = this.para.props;
    return Math.max(best, seg) + Math.max(0, p.indLeft) + Math.max(0, p.indRight);
  }

  // -------------------------------------------------------------------------------------------
  // Tabs

  /** Width of the text following atom i up to the next tab/break (and to a decimal point if asked). */
  private segmentWidth(i: number, decimal = false): number {
    let w = 0;
    for (let k = i + 1; k < this.atoms.length; k++) {
      const a = this.atoms[k];
      if (a.k === 'tab' || a.k === 'ptab' || a.k === 'br') break;
      if (decimal && a.k === 'g' && (a.cp === 0x2e || a.cp === 0x2c)) break;
      w += atomWidth(a);
    }
    // Trailing spaces do not count towards alignment.
    for (let k = i + 1; k < this.atoms.length; k++) {
      const a = this.atoms[k];
      if (a.k === 'tab' || a.k === 'ptab' || a.k === 'br') break;
    }
    return w;
  }

  private resolveTab(i: number, x: number, first: boolean, areaWidth: number, isLabelTab: boolean): { width: number; leader: TabStop['leader']; overflow: boolean; reach?: number } {
    const p = this.para.props;
    let stop: TabStop | null = null;
    for (const t of this.tabs) {
      if (t.pos > x + EPS) {
        stop = t;
        break;
      }
    }
    // The hanging indent acts as a tab stop on the first line (list numbers and plain text).
    if (first && p.indFirstLine < 0 && p.indLeft > x + EPS && (!stop || stop.pos > p.indLeft)) {
      stop = { pos: p.indLeft, type: 'left', leader: 'none' };
    }
    let isDefault = false;
    if (!stop) {
      const dts = this.env.settings.defaultTabStop > 0 ? this.env.settings.defaultTabStop : 36;
      const pos = (Math.floor((x + EPS) / dts) + 1) * dts;
      stop = { pos, type: 'left', leader: 'none' };
      isDefault = true;
    }
    void isLabelTab;
    const right = areaWidth - p.indRight;
    let target = stop.pos;
    if (stop.type === 'right' || stop.type === 'center' || stop.type === 'decimal') {
      const seg = this.segmentWidth(i, stop.type === 'decimal');
      // Since Word 2013, a right tab past the margin aligns its text at the margin instead.
      const pos = stop.type === 'right' && this.env.settings.compatibilityMode >= 15 ? Math.min(stop.pos, areaWidth) : stop.pos;
      if (stop.type === 'right' || stop.type === 'decimal') target = pos - seg;
      else target = pos - seg / 2;
      // Text aligned on a right tab may run past the right indent up to the tab (contents page
      // numbers hang into the indent that wraps their titles).
      if (stop.type === 'right') return { width: Math.max(0, target - x), leader: stop.leader, overflow: false, reach: pos };
    }
    const width = Math.max(0, target - x);
    const overflow = isDefault ? stop.pos > right + EPS : stop.type === 'left' && stop.pos > Math.max(right, areaWidth) + EPS;
    return { width, leader: stop.leader, overflow };
  }

  /** Are the atoms from `from` up to `to` only tabs, spaces and bookmarks? */
  private blank(from: number, to: number): boolean {
    for (let k = from; k < to; k++) {
      const a = this.atoms[k];
      if (a.k === 'tab' || a.k === 'bm' || a.k === 'rpb' || (a.k === 'g' && isSpaceCp(a.cp))) continue;
      return false;
    }
    return true;
  }

  private resolvePTab(i: number, x: number, a: Extract<Atom, { k: 'ptab' }>, areaWidth: number): { width: number; leader: TabStop['leader']; overflow: boolean } {
    const p = this.para.props;
    const left = a.relativeTo === 'indent' ? p.indLeft : 0;
    const right = a.relativeTo === 'indent' ? areaWidth - p.indRight : areaWidth;
    const seg = this.segmentWidth(i);
    let target = left;
    if (a.alignment === 'center') target = (left + right) / 2 - seg / 2;
    else if (a.alignment === 'right') target = right - seg;
    return { width: Math.max(0, target - x), leader: a.leader, overflow: false };
  }

  // -------------------------------------------------------------------------------------------
  // Rendering

  /**
   * Move a line's characters into visual order: right-to-left stretches run from right to left.
   * Trailing spaces stay at the line's end.
   */
  private reorder(line: Line, positions: number[], widthOf: (k: number) => number): void {
    const atoms = this.atoms;
    let end = line.end;
    while (end > line.start) {
      const a = atoms[end - 1];
      if (a.k === 'g' && isSpaceCp(a.cp)) end--;
      else if (a.k === 'br' || a.k === 'bm' || a.k === 'rpb') end--;
      else break;
    }
    if (end - line.start < 2) return;
    const levels = this.levels!.slice(line.start, end);
    if (!levels.some((l) => l > 0)) return;
    const order = visualOrder(levels);
    let x = positions[0];
    for (const i of order) {
      positions[i] = x;
      x += widthOf(line.start + i);
    }
  }

  /** Draw ops for a line whose top-left area origin is (ox, top). */
  renderLine(line: Line, ox: number, top: number): DrawOp[] {
    const ops: DrawOp[] = [];
    const deco: DrawOp[] = [];
    const back: DrawOp[] = [];
    const atoms = this.atoms;
    const baseline = top + line.baseline;
    const positions: number[] = [];
    let shift = line.alignShift;
    const jlast = (line.justify as { last?: number }).last ?? line.end - 1;
    let charIndex = 0;
    for (let k = line.start; k < line.end; k++) {
      const a = atoms[k];
      let x = ox + line.xs[k - line.start] + shift;
      if (k >= line.justify.from && k <= jlast) {
        if (line.justify.perChar && (a.k === 'g' || a.k === 'obj')) {
          x += charIndex * line.justify.perChar;
          charIndex++;
        }
      }
      positions.push(x);
      if (k >= line.justify.from && k <= jlast && line.justify.perGap && stretches(a, k < jlast)) shift += line.justify.perGap;
    }
    const widthOf = (k: number): number => {
      const a = atoms[k];
      let w = line.ws[k - line.start];
      if (k >= line.justify.from && k <= jlast && line.justify.perGap && stretches(a, k < jlast)) w += line.justify.perGap;
      return w;
    };
    if (this.levels) this.reorder(line, positions, widthOf);

    const p = this.para.props;
    const lineX = !this.levels && !line.alignShift && (p.align === 'left' || p.align === 'both') ? positions[0] : undefined;
    // Glyph runs
    let cur: GlyphOp | null = null;
    let curRun: ShapedRun | null = null;
    let segStart = -1;
    const flushDeco = (endK: number) => {
      if (!curRun || segStart < 0) return;
      this.decorate(curRun, segStart, endK, positions, widthOf, line, baseline, deco, back, ops);
    };
    for (let k = line.start; k < line.end; k++) {
      const a = atoms[k];
      const x = positions[k - line.start];
      if (a.k === 'g') {
        const g = a.run.glyphs[a.gi];
        if (curRun !== a.run) {
          flushDeco(k - 1);
          curRun = a.run;
          segStart = k;
          cur = null;
        }
        // A soft hyphen shows only where a line breaks at it, not at the paragraph's end.
        if (g.gid === -1 || (a.cp === 0xad && (k !== line.end - 1 || line.end >= atoms.length))) continue;
        if (!cur) {
          const run = a.run;
          cur = {
            kind: 'glyphs',
            field: a.field,
            x,
            y: baseline - run.baselineShift,
            face: run.face.face,
            size: run.size,
            gids: [],
            xs: [],
            texts: [],
            color: run.props.color ?? '000000',
            syntheticBold: run.face.syntheticBold,
            syntheticItalic: run.face.syntheticItalic,
            scaleX: run.scaleX,
            outline: run.props.outline || undefined,
          };
          if (lineX !== undefined) cur.lineX = lineX;
          ops.push(cur);
        }
        cur.gids.push(a.cp === 0xad ? run_hyphen(a.run) : g.gid);
        cur.xs.push(x - cur.x);
        cur.texts.push(a.cp === 0xad ? '-' : g.text);
      } else {
        flushDeco(k - 1);
        curRun = null;
        segStart = -1;
        cur = null;
        if (a.k === 'math') {
          ops.push(...translateOps(a.box.ops, x, baseline));
        } else if (a.k === 'obj') {
          const w = a.w;
          const h = a.h;
          const y = baseline - h - a.props.position;
          ops.push(...this.env.drawInline(a.drawing, x + a.drawing.effectL, y + a.drawing.effectT, a.drawing.width, a.drawing.height));
          if (a.link) ops.push({ kind: 'link', x, y, w, h, link: a.link });
        } else if (a.k === 'sep') {
          const w = a.w < 0 ? line.right - line.left : a.w;
          const y = top + line.height / 2;
          ops.push({ kind: 'line', x1: x, y1: y, x2: x + w, y2: y, width: 0.5, color: '000000' });
        } else if (a.k === 'bm') {
          ops.push({ kind: 'anchor', name: a.name, x, y: top });
        } else if ((a.k === 'tab' || a.k === 'ptab') && (a.run.props.underline || a.run.props.shading || a.run.props.highlight)) {
          this.decorate(a.run, k, k, positions, widthOf, line, baseline, deco, back, ops, true);
        }
        if ((a.k === 'tab' || a.k === 'ptab') && a.run.link) {
          ops.push({ kind: 'link', x, y: baseline - a.run.ascent, w: widthOf(k), h: a.run.ascent + a.run.descent, link: a.run.link });
        }
      }
    }
    flushDeco(line.end - 1);

    // Tab leaders
    for (const L of line.tabLeaders) {
      const ch = L.leader === 'dot' ? '.' : L.leader === 'hyphen' ? '-' : L.leader === 'underscore' || L.leader === 'heavy' ? '_' : L.leader === 'middleDot' ? '·' : '';
      if (!ch) continue;
      const runs = shapeSimple(ch, L.run.props, this.env.fonts);
      const r = runs[0];
      if (!r || !r.glyphs[0]) continue;
      const gw = r.glyphs[0].advance;
      if (gw <= 0) continue;
      const x0 = ox + L.x + line.alignShift;
      // Word aligns leader characters to a fixed grid so dots line up across lines.
      const startX = Math.ceil((x0 + gw * 0.25) / gw) * gw;
      const endX = x0 + L.w - gw * 0.25;
      const op: GlyphOp = {
        kind: 'glyphs',
        x: startX,
        y: baseline,
        face: r.face.face,
        size: r.size,
        gids: [],
        xs: [],
        texts: [],
        color: r.props.color ?? '000000',
        syntheticBold: r.face.syntheticBold || L.leader === 'heavy',
        syntheticItalic: false,
        scaleX: r.scaleX,
      };
      for (let px = startX; px + gw <= endX + EPS; px += gw) {
        op.gids.push(r.glyphs[0].gid);
        op.xs.push(px - startX);
        op.texts.push(ch);
      }
      if (op.gids.length) ops.push(op);
    }
    return [...back, ...ops, ...deco];
  }

  private decorate(
    run: ShapedRun,
    from: number,
    to: number,
    positions: number[],
    widthOf: (k: number) => number,
    line: Line,
    baseline: number,
    deco: DrawOp[],
    back: DrawOp[],
    ops: DrawOp[],
    isTab = false,
  ): void {
    const props = run.props;
    if (to < from) return;
    const lineStart = line.start;
    // Trailing spaces at the end of a line are not underlined or highlighted.
    let last = to;
    if (!isTab) {
      while (last >= from) {
        const a = this.atoms[last];
        if (a.k === 'g' && isSpaceCp(a.cp) && last >= this.lastVisibleIndex(line)) last--;
        else break;
      }
    }
    if (last < from) return;
    const x1 = positions[from - lineStart];
    const x2 = positions[last - lineStart] + widthOf(last);
    if (x2 <= x1) return;
    const font = run.face.face.font;
    const em = run.size / font.unitsPerEm;
    const yBase = baseline - run.baselineShift;
    const asc = run.face.ascent * run.size;
    const desc = run.face.descent * run.size;
    if (props.highlight) back.push({ kind: 'rect', x: x1, y: yBase - asc, w: x2 - x1, h: asc + desc, fill: props.highlight });
    else if (props.shading) back.push({ kind: 'rect', x: x1, y: yBase - asc, w: x2 - x1, h: asc + desc, fill: props.shading });
    const color = props.color ?? '000000';
    if (props.underline) {
      const uColor = props.underlineColor ?? color;
      // Word positions the underline from the base font size even for super/subscript.
      const thick = Math.max(0.5, (font.underlineThickness || 50) * em);
      const uy = yBase - (font.underlinePosition || -100) * em + thick / 2;
      const style = props.underline;
      const w = style === 'thick' || style.endsWith('Heavy') ? thick * 2 : thick;
      if (style === 'words') {
        // Underline each word, not the spaces between them.
        let segX: number | null = null;
        for (let k = from; k <= last; k++) {
          const a = this.atoms[k];
          const isSp = a.k === 'g' && isSpaceCp(a.cp);
          const px = positions[k - lineStart];
          if (!isSp && segX === null) segX = px;
          if (isSp && segX !== null) {
            deco.push({ kind: 'line', x1: segX, y1: uy, x2: px, y2: uy, width: thick, color: uColor });
            segX = null;
          }
        }
        if (segX !== null) deco.push({ kind: 'line', x1: segX, y1: uy, x2: x2, y2: uy, width: thick, color: uColor });
      } else if (style === 'double') {
        deco.push({ kind: 'line', x1, y1: uy - thick, x2, y2: uy - thick, width: thick, color: uColor });
        deco.push({ kind: 'line', x1, y1: uy + thick, x2, y2: uy + thick, width: thick, color: uColor });
      } else if (style.startsWith('wave')) {
        const amp = Math.max(0.6, thick);
        const step = amp * 3;
        const segs: Array<{ op: 'M' | 'L'; pts: number[] }> = [{ op: 'M', pts: [x1, uy] }];
        let up = true;
        for (let px = x1 + step / 2; px < x2; px += step / 2) {
          segs.push({ op: 'L', pts: [Math.min(px, x2), uy + (up ? -amp : amp)] });
          up = !up;
        }
        deco.push({ kind: 'path', segs, stroke: uColor, strokeWidth: thick * 0.8, fill: null });
      } else {
        const dash = style.startsWith('dotted') || style === 'dottedHeavy' ? [w, w * 1.5] : style.startsWith('dash') ? [w * 4, w * 2] : style.startsWith('dotDash') ? [w * 4, w * 2, w, w * 2] : undefined;
        deco.push({ kind: 'line', x1, y1: uy, x2, y2: uy, width: w, color: uColor, dash });
      }
    }
    if (props.strike || props.dstrike) {
      const thick = Math.max(0.5, (font.strikeoutSize || 50) * em);
      const sy = yBase - (font.strikeoutPosition || font.unitsPerEm * 0.25) * em + thick / 2;
      if (props.dstrike) {
        deco.push({ kind: 'line', x1, y1: sy - thick, x2, y2: sy - thick, width: thick, color });
        deco.push({ kind: 'line', x1, y1: sy + thick, x2, y2: sy + thick, width: thick, color });
      } else deco.push({ kind: 'line', x1, y1: sy, x2, y2: sy, width: thick, color });
    }
    if (props.border) {
      const b = props.border;
      const bx1 = x1 - b.space;
      const bx2 = x2 + b.space;
      const by1 = yBase - asc;
      const by2 = yBase + desc;
      const c = b.color ?? '000000';
      deco.push({ kind: 'line', x1: bx1, y1: by1, x2: bx2, y2: by1, width: b.width, color: c });
      deco.push({ kind: 'line', x1: bx1, y1: by2, x2: bx2, y2: by2, width: b.width, color: c });
      deco.push({ kind: 'line', x1: bx1, y1: by1, x2: bx1, y2: by2, width: b.width, color: c });
      deco.push({ kind: 'line', x1: bx2, y1: by1, x2: bx2, y2: by2, width: b.width, color: c });
    }
    if (run.link && !isTab) ops.push({ kind: 'link', x: x1, y: yBase - asc, w: x2 - x1, h: asc + desc, link: run.link });
  }

  private lastVisibleIndex(line: Line): number {
    for (let k = line.end - 1; k >= line.start; k--) {
      const a = this.atoms[k];
      if (!(a.k === 'g' && isSpaceCp(a.cp)) && a.k !== 'br') return k + 1;
    }
    return line.start;
  }
}

function run_hyphen(run: ShapedRun): number {
  return run.face.face.font.glyphId(0x2d);
}

/** Whether justification widens the space this atom carries: a space, or the space before Latin
 * text inside an East Asian line (`inside`: not the line's last visible atom). */
function stretches(a: Atom, inside: boolean): boolean {
  return a.k === 'g' && (a.cp === 0x20 || a.cp === 0x3000 || (inside && !!a.aspace));
}

function atomWidth(a: Atom): number {
  switch (a.k) {
    case 'g':
      return a.w;
    case 'obj':
      return a.w;
    case 'math':
      return a.box.w;
    case 'sep':
      return a.w < 0 ? 0 : a.w;
    default:
      return 0;
  }
}

function emptyRun(props: RunProps, fonts: FontRegistry): ShapedRun {
  const face = fonts.resolve(props.fontAscii, props.bold, props.italic);
  return {
    face,
    size: props.size,
    props,
    glyphs: [],
    baselineShift: 0,
    ascent: face.ascent * props.size,
    descent: face.descent * props.size,
    extLeading: face.externalLeading * props.size,
    scaleX: 1,
  };
}

function pictureBullet(image: Drawing['image'], size: number): Drawing {
  return { kind: 'image', inline: true, width: size, height: size, effectL: 0, effectT: 0, effectR: 0, effectB: 0, rotation: 0, flipH: false, flipV: false, geometry: 'rect', hidden: false, image };
}

/**
 * Word's line height: font height (Windows ascent + descent, plus GDI external leading) scaled
 * by the line spacing rule. Text sits at the bottom of the line box, so extra spacing goes above.
 */
export function lineHeight(
  ascent: number,
  descent: number,
  extLeading: number,
  p: Paragraph['props'],
  gridPitch?: number,
  textHeight = ascent + descent + extLeading,
): { height: number; baseline: number } {
  const natural = ascent + descent + extLeading;
  let height: number;
  if (p.lineRule === 'exact') height = p.line;
  else if (p.lineRule === 'atLeast') height = Math.max(natural, p.line);
  // "Multiple" spacing scales the text height only: inline pictures are not stretched.
  else height = natural + textHeight * (p.line - 1);
  if (gridPitch && p.snapToGrid && p.lineRule === 'auto' && gridPitch > 0) {
    // The grid lines a line takes count the font height without its external leading.
    const base = Math.max(ascent + descent, gridPitch);
    // A line 0.9% taller than a grid line takes two (14pt SimHei, 18.16pt, on an 18pt grid). The
    // 0.5% allowance covers fallback fonts' approximate heights: KaiTi_GB2312 at Noto Sans CJK's
    // line comes out 0.3% over a 15.6pt grid line that Word's own KaiTi fits.
    const snapped = Math.ceil(base / gridPitch - 0.005) * gridPitch;
    // "Multiple" spacing scales one grid line: text that takes more grid lines than that keeps
    // just the lines it takes.
    height = p.line > 1 ? Math.max(snapped, gridPitch * p.line) : snapped * p.line;
    // The document grid centres the text in the grid lines it takes.
    return { height, baseline: (snapped - natural) / 2 + natural - descent };
  }
  // Extra "multiple" spacing goes below the text; exact and at-least heights keep the text at the
  // bottom of the line box.
  if (p.lineRule !== 'exact' && p.lineRule !== 'atLeast') return { height, baseline: natural - descent };
  return { height, baseline: height - descent };
}
