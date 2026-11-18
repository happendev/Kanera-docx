import type { FontRegistry } from '../fonts/registry.js';
import type { DocSettings, Drawing, FieldInline, Link, Paragraph, RunProps, TabStop } from '../docx/model.js';
import { shapeSimple, shapeText, type ShapedRun } from './shaper.js';
import type { DrawOp, GlyphOp } from './types.js';

const EPS = 0.01;

export interface ParagraphEnv {
  fonts: FontRegistry;
  settings: DocSettings;
  fieldText(field: FieldInline): string;
  noteRefText(noteType: 'footnote' | 'endnote', id: string): string;
  /** Number shown by w:footnoteRef inside the note currently being laid out. */
  noteNumberText?: string;
  /** Line pitch when the section uses a "lines" document grid. */
  gridPitch?: number;
  drawInline(drawing: Drawing, x: number, y: number, w: number, h: number): DrawOp[];
}

type Atom =
  | { k: 'g'; run: ShapedRun; gi: number; w: number; cp: number }
  | { k: 'tab'; run: ShapedRun; label?: boolean }
  | { k: 'ptab'; run: ShapedRun; alignment: 'left' | 'center' | 'right'; relativeTo: 'margin' | 'indent'; leader: TabStop['leader'] }
  | { k: 'br'; type: 'line' | 'page' | 'column'; run: ShapedRun }
  | { k: 'obj'; drawing: Drawing; w: number; h: number; props: RunProps; link?: Link; run: ShapedRun }
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

const NO_BREAK_BEFORE = new Set(
  [...'、。，．：；？！）」』】〕〉》’”ー々ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ…‥・゛゜ヽヾゝゞ〟｝］｠｡｣､﹐﹑﹒﹔﹕﹖﹗!%),.:;?]}¢°’”‰℃'].map((c) => c.codePointAt(0)!),
);
const NO_BREAK_AFTER = new Set([...'（「『【〔〈《‘“｛［｟｢([{£¥$'].map((c) => c.codePointAt(0)!));

function isIdeographic(cp: number): boolean {
  return (cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xac00 && cp <= 0xd7af) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef) || (cp >= 0x20000 && cp <= 0x2fa1f);
}

function isAlnum(cp: number): boolean {
  return (cp >= 0x30 && cp <= 0x39) || (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || (cp >= 0xc0 && cp <= 0x24f) || (cp >= 0x370 && cp <= 0x52f);
}

export class ParagraphLayout {
  readonly atoms: Atom[] = [];
  private breakAfter: boolean[] = [];
  pos = 0;
  private firstLine = true;
  readonly markRun: ShapedRun;
  /** Width the numbering label is shifted left by (lvlJc right/center). */
  private labelShift = 0;
  private labelEnd = 0;
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
          const text = env.fieldText(inl);
          for (const run of shapeText(text, inl.props, fonts)) for (let gi = 0; gi < run.glyphs.length; gi++) this.pushGlyph(run, gi);
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
      }
    }
    this.computeBreaks();
  }

  private pushGlyph(run: ShapedRun, gi: number): void {
    const g = run.glyphs[gi];
    this.atoms.push({ k: 'g', run, gi, w: g.advance, cp: g.cp });
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
      if (a.k === 'tab' || a.k === 'ptab' || a.k === 'obj' || a.k === 'br') {
        this.breakAfter[i] = true;
        if (a.k === 'obj' && i > 0) this.breakAfter[i - 1] = true;
        continue;
      }
      if (a.k !== 'g') continue;
      const cp = a.cp;
      // Look past zero-width atoms for the next character.
      let j = i + 1;
      while (j < n && atoms[j].k !== 'g' && atoms[j].k !== 'tab' && atoms[j].k !== 'obj' && atoms[j].k !== 'br' && atoms[j].k !== 'ptab') j++;
      const next = cpAt(j);
      if (isSpaceCp(cp)) {
        // Break after the last space of a run of spaces.
        this.breakAfter[i] = !(next >= 0 && isSpaceCp(next));
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
    const left = p.indLeft + (this.firstLine ? p.indFirstLine : 0);
    return { left, right: areaWidth - p.indRight };
  }

  /**
   * Lay out the next line between `left` and `right` (area coordinates). Returns null when the
   * paragraph is complete.
   */
  nextLine(left: number, right: number, areaWidth: number): Line | null {
    const atoms = this.atoms;
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

    const finish = (end: number): number => end;
    let end = -1;
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
        const t = a.k === 'tab' ? this.resolveTab(i, x, first, areaWidth, a.label === true) : this.resolvePTab(i, x, a, areaWidth);
        if (t.overflow && hasContent && right - left > 1) {
          end = candidate > start ? candidate : i;
          break;
        }
        xs.push(x);
        ws.push(t.width);
        if (t.leader !== 'none' && t.width > 0) leaders.push({ x, w: t.width, leader: t.leader, run: a.run });
        x += t.width;
        lastTab = i;
        hasContent = true;
        candidate = i + 1;
        candidateX = x;
        continue;
      }
      const w = atomWidth(a);
      if (a.k === 'g' && isSpaceCp(a.cp)) {
        xs.push(x);
        ws.push(w);
        x += w;
        if (this.breakAfter[i]) {
          candidate = i + 1;
          candidateX = x;
        }
        continue;
      }
      if (w > 0 && x + w > right + EPS && hasContent) {
        if (candidate > start) end = candidate;
        else end = i; // emergency break inside a word
        break;
      }
      xs.push(x);
      ws.push(w);
      x += w;
      if (w > 0 || a.k === 'obj') hasContent = true;
      if (this.breakAfter[i]) {
        candidate = i + 1;
        candidateX = x;
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

    // Visible extent (exclude trailing spaces).
    let contentRight = left;
    for (let k = end - 1; k >= start; k--) {
      const a = atoms[k];
      if (a.k === 'g' && isSpaceCp(a.cp)) continue;
      if (a.k === 'br') continue;
      if ((a.k === 'g' || a.k === 'obj' || a.k === 'tab' || a.k === 'ptab' || a.k === 'sep') && ws[k - start] >= 0) {
        contentRight = xs[k - start] + ws[k - start];
        break;
      }
    }

    // Metrics
    let ascent = 0;
    let descent = 0;
    let ext = 0;
    let textAscent = 0;
    let textDescent = 0;
    let measured = false;
    const add = (asc: number, desc: number, lead: number, text = true) => {
      ascent = Math.max(ascent, asc);
      descent = Math.max(descent, desc);
      ext = Math.max(ext, lead);
      if (text) {
        textAscent = Math.max(textAscent, asc);
        textDescent = Math.max(textDescent, desc);
      }
      measured = true;
    };
    const floats: Drawing[] = [];
    const notes: Line['notes'] = [];
    let rpb = false;
    for (let k = start; k < end; k++) {
      const a = atoms[k];
      switch (a.k) {
        case 'g':
          if (a.run.glyphs[a.gi].gid === -1) break;
          add(a.run.ascent + a.run.baselineShift, a.run.descent - a.run.baselineShift, a.run.extLeading);
          break;
        case 'tab':
        case 'ptab':
        case 'br':
        case 'sep':
          add(a.run.ascent, a.run.descent, a.run.extLeading);
          break;
        case 'obj':
          add(a.h + a.props.position, Math.max(0, -a.props.position), 0, false);
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
    if (!measured || textAscent + textDescent === 0) add(this.markRun.ascent, this.markRun.descent, this.markRun.extLeading);
    const { height, baseline } = lineHeight(ascent, descent, ext, this.para.props, this.env.gridPitch, textAscent + textDescent + ext);

    // Alignment and justification.
    const p = this.para.props;
    const avail = right - left;
    const used = contentRight - left;
    let alignShift = 0;
    const justify = { from: start, perGap: 0, perChar: 0 };
    const endsWithLineBreak = brk === 'line';
    const justifyThis = (p.align === 'both' && !isLast && !(endsWithLineBreak && this.env.settings.doNotExpandShiftReturn)) || p.align === 'distribute';
    if (justifyThis && used < avail) {
      const extra = avail - used;
      // Only spaces after the last tab stretch.
      let from = start;
      for (let k = start; k < end; k++) if (atoms[k].k === 'tab' || atoms[k].k === 'ptab') from = k + 1;
      if (first && this.hasLabel) from = Math.max(from, start + this.labelEnd);
      let lastVisible = end - 1;
      while (lastVisible >= from && atoms[lastVisible].k === 'g' && isSpaceCp((atoms[lastVisible] as { cp: number }).cp)) lastVisible--;
      let gaps = 0;
      let chars = 0;
      for (let k = from; k <= lastVisible; k++) {
        const a = atoms[k];
        if (a.k === 'g' && (a.cp === 0x20 || a.cp === 0x3000)) gaps++;
        if (a.k === 'g' || a.k === 'obj') chars++;
      }
      justify.from = from;
      if (p.align === 'distribute') {
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
      first,
      floats,
      notes,
      renderedPageBreak: rpb,
      justify,
      alignShift,
      tabLeaders: leaders,
    };
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

  private resolveTab(i: number, x: number, first: boolean, areaWidth: number, isLabelTab: boolean): { width: number; leader: TabStop['leader']; overflow: boolean } {
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
      if (stop.type === 'right' || stop.type === 'decimal') target = stop.pos - seg;
      else target = stop.pos - seg / 2;
    }
    const width = Math.max(0, target - x);
    const overflow = isDefault ? stop.pos > right + EPS : stop.type === 'left' && stop.pos > Math.max(right, areaWidth) + EPS;
    return { width, leader: stop.leader, overflow };
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
      if (k >= line.justify.from && k <= jlast && line.justify.perGap && a.k === 'g' && (a.cp === 0x20 || a.cp === 0x3000)) shift += line.justify.perGap;
    }
    const widthOf = (k: number): number => {
      const a = atoms[k];
      let w = line.ws[k - line.start];
      if (k >= line.justify.from && k <= jlast && line.justify.perGap && a.k === 'g' && (a.cp === 0x20 || a.cp === 0x3000)) w += line.justify.perGap;
      return w;
    };

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
        if (g.gid === -1 || (a.cp === 0xad && k !== line.end - 1)) continue;
        if (!cur) {
          const run = a.run;
          cur = {
            kind: 'glyphs',
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
        if (a.k === 'obj') {
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
    const asc = run.face.face.ascent * run.size;
    const desc = run.face.face.descent * run.size;
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

function atomWidth(a: Atom): number {
  switch (a.k) {
    case 'g':
      return a.w;
    case 'obj':
      return a.w;
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
    ascent: face.face.ascent * props.size,
    descent: face.face.descent * props.size,
    extLeading: face.face.externalLeading * props.size,
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
    const base = Math.max(natural, gridPitch);
    height = Math.ceil(base / gridPitch - 0.01) * gridPitch * p.line;
  }
  return { height, baseline: height - descent };
}
