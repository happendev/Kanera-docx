/**
 * Plain text and Markdown from the resolved document model. The parser has already applied what
 * decides which text a reader sees: tracked changes (insertions kept, deletions dropped), hidden
 * text, field results, content-control bindings and list numbers. This module puts that text in
 * reading order and, for Markdown, maps Word's structure (headings, lists, tables, emphasis, links,
 * notes, equations) onto Markdown's.
 */
import type { Block, Chart, DocxDocument, Drawing, Link, NumberingLabel, Paragraph, RunProps, Section, Table, TableCell } from '../ooxml/model.js';
import { formatNumber as formatChartNumber } from '../ooxml/number-format.js';
import { formatNumber } from '../docx/numbering.js';
import { fontFor } from '../fonts/script.js';
import { mapSymbolChar, symbolFontKind, type SymbolFontKind } from '../fonts/symbol-map.js';
import { mathToLatex, mathToText } from './math.js';

export type TextFormat = 'text' | 'markdown';

export interface TextOptions {
  /** 'text' (default) for plain text, 'markdown' for GitHub-flavoured Markdown. */
  format?: TextFormat;
  /**
   * Include headers and footers (default false). Each distinct header appears once, before the
   * first section that uses it, and each footer once, after it.
   */
  includeHeadersFooters?: boolean;
  /**
   * Include footnotes and endnotes (default true). References show as [1] in plain text and as
   * Markdown footnotes ([^1]); the notes follow the body.
   */
  includeNotes?: boolean;
  /** Markdown only: show pictures as `![alt text](word/media/image1.png)` (default true). */
  includeImages?: boolean;
}

// Inline formatting flags.
const BOLD = 1;
const ITALIC = 2;
const STRIKE = 4;
const SUP = 8;
const SUB = 16;
const CODE = 32;

/** A piece of paragraph text with uniform formatting. */
interface Seg {
  text: string;
  fmt: number;
  link?: Link;
  /** Markdown emitted as is (note references, pictures, equations, line breaks). */
  raw?: boolean;
  /** A line break. */
  br?: boolean;
  /** Raw Markdown that emphasis opens and closes around like a word: an autolink. */
  auto?: boolean;
}

/** Something a paragraph's drawings hold that is shown after the paragraph. */
type Extra = { blocks: Block[] } | { chart: Chart } | { text: string } | { image: string } | { rule: true };

/** An output block: Markdown blocks are separated by blank lines, list items and quotes by less. */
interface Chunk {
  text: string;
  kind: 'block' | 'list' | 'quote';
}

const PUA_SYMBOL = /[\uf020-\uf0ff]/;

/** The face Word draws a run in, chosen by the script of its first letter (or first character). */
function face(props: RunProps, text: string): ReturnType<typeof fontFor> {
  const m = /\p{L}/u.exec(text) ?? /\S/u.exec(text);
  return fontFor(m ? m[0].codePointAt(0)! : 0x20, props);
}

function isMonospace(family: string): boolean {
  return /mono|courier|consolas|menlo|monaco|lucida console|lucida sans typewriter|inconsolata|source code|fira code|cascadia|andale|letter gothic|prestige elite|fixedsys|\bocr[- ]?[ab]\b|iosevka|\bhack\b/i.test(family);
}

const NOTE_LABEL = /^[A-Za-z0-9_-]+$/;

const PUNCTUATION = /[\p{P}\p{S}]/u;
const FLANKED = [BOLD, ITALIC, STRIKE];

/**
 * CommonMark opens `**`, `*` or `~~` before punctuation only after whitespace or punctuation, and
 * closes it after punctuation only before whitespace or punctuation: in `Page**: Black**` or
 * `**名称:**lxp` the markers would show as text. Punctuation at such an edge of bold, italic or
 * struck text is left outside the markers instead; emphasis of punctuation alone is dropped.
 */
/**
 * Links that show their own address become autolinks set in the text around them, so emphasis
 * over the address and its sentence stays one run: **Visit <https://example.com> today**.
 */
function autolinks(segs: Seg[]): Seg[] {
  const out: Seg[] = [];
  for (let i = 0; i < segs.length; ) {
    const link = segs[i].link;
    let j = i + 1;
    while (j < segs.length && segs[j].link === link) j++;
    const group = segs.slice(i, j);
    i = j;
    const url = link?.url && !link.anchor ? link.url : '';
    const text = group.map((s) => s.text).join('');
    const shown = group.filter((s) => s.text.trim());
    const fmt = shown[0]?.fmt ?? 0;
    if (!url || text.trim() !== url || !/^(https?|mailto|ftp):/i.test(url) || /[<>\s]/.test(url) ||
        group.some((s) => s.raw || s.br) || shown.some((s) => s.fmt !== fmt) || fmt & ~(BOLD | ITALIC | STRIKE)) {
      out.push(...group);
      continue;
    }
    const lead = /^\s*/.exec(text)![0];
    const trail = /\s*$/.exec(text)![0];
    if (lead) out.push({ text: lead, fmt: group[0].fmt });
    out.push({ text: `<${url}>`, fmt, raw: true, auto: true });
    if (trail) out.push({ text: trail, fmt: group[group.length - 1].fmt });
  }
  return out;
}

function flanking(segs: Seg[]): Seg[] {
  if (!segs.some((s) => (!s.raw || s.auto) && !s.br && s.fmt & (BOLD | ITALIC | STRIKE))) return segs;
  // One entry per character; line breaks and raw Markdown separate the text like whitespace.
  const chars: Array<{ ch: string; fmt: number; seg: number }> = [];
  segs.forEach((s, seg) => {
    // An autolink starts and ends with punctuation: < and >.
    if (s.auto) for (const ch of '<a>') chars.push({ ch, fmt: s.fmt, seg });
    else if (s.raw || s.br) chars.push({ ch: ' ', fmt: -1, seg });
    else for (const ch of s.text) chars.push({ ch, fmt: s.fmt, seg });
  });
  const kind = (i: number) => {
    if (i < 0 || i >= chars.length || chars[i].fmt < 0 || /\s/.test(chars[i].ch)) return 'space';
    return PUNCTUATION.test(chars[i].ch) ? 'punct' : 'other';
  };
  let changed = false;
  for (const bit of FLANKED) {
    for (let i = 0; i < chars.length; ) {
      if (chars[i].fmt < 0 || !(chars[i].fmt & bit) || kind(i) === 'space') {
        i++;
        continue;
      }
      // The run with this bit, through whitespace, trimmed to its first and last visible characters.
      let end = i;
      for (let k = i; k < chars.length && chars[k].fmt >= 0 && (chars[k].fmt & bit || kind(k) === 'space'); k++) if (kind(k) !== 'space') end = k;
      const start = i;
      i = end + 1;
      let a = start;
      let b = end;
      if (kind(a) === 'punct' && kind(a - 1) === 'other') while (a <= b && kind(a) !== 'other') a++;
      if (a <= b && kind(b) === 'punct' && kind(b + 1) === 'other') while (b >= a && kind(b) !== 'other') b--;
      if (a === start && b === end) continue;
      changed = true;
      for (let k = start; k <= end; k++) if ((k < a || k > b) && chars[k].fmt >= 0) chars[k].fmt &= ~bit;
    }
  }
  if (!changed) return segs;
  const out: Seg[] = [];
  for (const [k, c] of chars.entries()) {
    const source = segs[c.seg];
    const last = out[out.length - 1];
    if (source.auto) {
      // An autolink keeps only the emphasis all of it has.
      if (chars[k - 1]?.seg !== c.seg) out.push({ ...source, fmt: chars.filter((x) => x.seg === c.seg).reduce((m, x) => m & x.fmt, source.fmt) });
    } else if (c.fmt < 0) out.push(source);
    else if (last && !last.raw && !last.br && last.fmt === c.fmt && last.link === source.link) last.text += c.ch;
    else out.push({ ...source, text: c.ch, fmt: c.fmt });
  }
  return out;
}

interface ListFrame {
  level: number;
  /** Column where the item's text starts, so nested items and continuation lines line up with it. */
  indent: number;
}

class TextExtractor {
  private readonly md: boolean;
  private readonly styleNames: Map<string, string>;
  /** Note numbers in reference order, as Word numbers them. */
  private readonly noteLabels = new Map<string, string>();
  private readonly noteOrder: Array<{ type: 'footnote' | 'endnote'; id: string; label: string; custom: boolean }> = [];
  private readonly usedLabels = new Set<string>();
  private footnoteCounter: number | undefined;
  private endnoteCounter: number;
  private customNotes = 0;
  private section: Section;
  private readonly codeFonts: boolean;
  private readonly escapeDollars: boolean;
  private symbolCache = new WeakMap<RunProps, boolean>();

  constructor(private readonly doc: DocxDocument, private readonly options: TextOptions) {
    this.md = options.format === 'markdown';
    this.styleNames = doc.styleNames ?? new Map();
    this.section = doc.finalSection;
    this.endnoteCounter = doc.settings.endnoteStart ?? 1;
    // Monospaced runs read as code only where the body is set in another face: a contract typed
    // throughout in Courier is not code.
    this.codeFonts = this.md && this.monospaceShare() < 0.5;
    this.escapeDollars = this.md && doc.hasMath;
  }

  run(): string {
    const chunks: Chunk[] = [];
    const sections = this.sections();
    const seen = new Set<Block[]>();
    const hf = (refs: Section['headers'], section: Section) => {
      // Only the headers Word shows: a first-page one when the section has a title page, an
      // even-page one when the document distinguishes odd and even pages.
      const used = [section.titlePage ? refs.first : undefined, refs.default, this.doc.settings.evenAndOddHeaders ? refs.even : undefined];
      for (const blocks of used) {
        if (!blocks || seen.has(blocks)) continue;
        seen.add(blocks);
        this.blocks(blocks, chunks);
      }
    };
    for (const { section, blocks } of sections) {
      this.section = section;
      if (section.footnoteRestart === 'eachSect') this.footnoteCounter = undefined;
      if (this.options.includeHeadersFooters) hf(section.headers, section);
      this.blocks(blocks, chunks);
      if (this.options.includeHeadersFooters) hf(section.footers, section);
    }
    if (this.options.includeNotes !== false) this.notes(chunks);
    return this.join(chunks);
  }

  // -------------------------------------------------------------------------------------------
  // Document structure

  private sections(): Array<{ section: Section; blocks: Block[] }> {
    const out: Array<{ section: Section; blocks: Block[] }> = [];
    let start = 0;
    const body = this.doc.body;
    body.forEach((b, i) => {
      if (b.type === 'paragraph' && b.sectPr) {
        out.push({ section: b.sectPr, blocks: body.slice(start, i + 1) });
        start = i + 1;
      }
    });
    out.push({ section: this.doc.finalSection, blocks: body.slice(start) });
    return out;
  }

  private monospaceShare(): number {
    let mono = 0;
    let total = 0;
    const visit = (blocks: Block[]) => {
      for (const b of blocks) {
        if (total > 20000) return;
        if (b.type === 'table') {
          for (const r of b.rows) for (const c of r.cells) visit(c.content);
          continue;
        }
        for (const c of b.content) {
          if (c.type !== 'text') continue;
          total += c.text.length;
          if (isMonospace(face(c.props, c.text).family)) mono += c.text.length;
        }
      }
    };
    visit(this.doc.body);
    return total ? mono / total : 0;
  }

  private join(chunks: Chunk[]): string {
    let out = '';
    let prev: Chunk | undefined;
    for (const c of chunks) {
      if (prev) {
        if (!this.md) out += '\n';
        else if (prev.kind === 'list' && c.kind === 'list') out += '\n';
        else if (prev.kind === 'quote' && c.kind === 'quote') out += '\n>\n';
        else out += '\n\n';
      }
      out += c.text;
      prev = c;
    }
    if (this.md) out = out.replace(/\n{3,}/g, '\n\n');
    else out = out.replace(/[ \u00a0]+$/gm, '').replace(/\n{3,}/g, '\n\n');
    out = out.replace(/^\n+/, '').replace(/\s+$/, '');
    return out ? out + '\n' : '';
  }

  private styleName(styleId: string | undefined): string {
    return (styleId && this.styleNames.get(styleId)?.toLowerCase()) || '';
  }

  // -------------------------------------------------------------------------------------------
  // Blocks

  private blocks(blocks: Block[], out: Chunk[]): void {
    let list: ListFrame[] = [];
    let code: string[] | undefined;
    let dropCap: Seg[] | undefined;
    const flushCode = () => {
      if (!code) return;
      const body = code.join('\n').replace(/\s+$/, '');
      const fence = '`'.repeat(Math.max(3, longestRun(body, '`') + 1));
      out.push({ text: `${fence}\n${body}\n${fence}`, kind: 'block' });
      code = undefined;
    };
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'table') {
        flushCode();
        list = [];
        if (dropCap) out.push({ text: this.md ? this.renderSegs(dropCap) : this.plain(dropCap), kind: 'block' });
        dropCap = undefined;
        this.table(b, out);
        continue;
      }
      const { segs, extras } = this.inlines(b);
      // A dropped capital is a frame of its own holding the first letter(s) of the next paragraph.
      if (b.props.frame?.dropCap && b.props.frame.dropCap !== 'none' && segs.some((s) => s.text.trim())) {
        dropCap = (dropCap ?? []).concat(segs.filter((s) => !s.br));
        this.extras(extras, out);
        continue;
      }
      if (dropCap) {
        segs.unshift(...dropCap);
        dropCap = undefined;
      }
      if (!this.md) {
        const label = b.numbering ? this.labelText(b.numbering) : '';
        const text = label + this.plain(segs);
        out.push({ text, kind: 'block' });
        this.extras(extras, out);
        continue;
      }
      // Markdown
      const name = this.styleName(b.props.styleId);
      if (this.isCode(b, segs, name)) {
        list = [];
        (code ??= []).push(segs.map((s) => (s.br ? '\n' : s.raw ? '' : s.text)).join('').replace(/\t/g, '    '));
        this.extras(extras, out);
        continue;
      }
      flushCode();
      const displayMath = this.displayMath(b);
      if (displayMath !== undefined) {
        list = [];
        out.push({ text: `$$\n${displayMath}\n$$`, kind: 'block' });
        this.extras(extras, out);
        continue;
      }
      const heading = this.headingLevel(b, name, segs);
      const toc = /^toc ([1-9])$/.exec(name);
      const label = b.numbering && b.numbering.text.trim() ? b.numbering : undefined;
      if (heading) {
        list = [];
        const prefix = label ? escapeMd(this.labelText(label, true).trim(), false, this.escapeDollars) + ' ' : '';
        const text = this.renderSegs(segs.map((s) => (s.br ? { text: ' ', fmt: 0 } : { ...s, fmt: s.fmt & ~(BOLD | ITALIC) })), false);
        if (prefix || text.trim()) out.push({ text: `${'#'.repeat(heading)} ${(prefix + text).trim()}`, kind: 'block' });
      } else if (label || toc) {
        // Table-of-contents entries nest by their TOC level, as Word indents them, even when numbered.
        const level = toc ? Number(toc[1]) - 1 : label!.level;
        while (list.length && list[list.length - 1].level >= level) list.pop();
        const indent = list.length ? list[list.length - 1].indent : 0;
        const marker = label ? this.listMarker(label) : '- ';
        const body = this.renderSegs(segs, false);
        // An item's content starts after its list marker ("- ", "12. "), not after a label kept as
        // text ("- a) "): nested items four or more columns further in would read as code.
        const content = indent + /^(?:-|\d+[.)]) /.exec(marker)![0].length;
        list.push({ level, indent: content });
        out.push({ text: ' '.repeat(indent) + marker + indentLines(body, content), kind: 'list' });
      } else {
        const body = this.renderSegs(segs, false);
        if (!body.trim()) {
          this.extras(extras, out);
          continue;
        }
        list = [];
        if (/quote/.test(name)) out.push({ text: body.split('\n').map((l) => '> ' + l).join('\n'), kind: 'quote' });
        else out.push({ text: body, kind: 'block' });
      }
      this.extras(extras, out);
    }
    flushCode();
    if (dropCap) out.push({ text: this.md ? this.renderSegs(dropCap) : this.plain(dropCap), kind: 'block' });
  }

  private extras(extras: Extra[], out: Chunk[]): void {
    for (const e of extras) {
      if ('blocks' in e) this.blocks(e.blocks, out);
      else if ('chart' in e) this.chart(e.chart, out);
      else if ('text' in e) out.push({ text: this.md ? escapeMd(e.text, false, this.escapeDollars) : e.text, kind: 'block' });
      else if ('image' in e) out.push({ text: e.image, kind: 'block' });
      else if (this.md) out.push({ text: '---', kind: 'block' });
    }
  }

  private headingLevel(p: Paragraph, name: string, segs: Seg[]): number {
    let level = 0;
    const m = /^heading\s*([1-9])$/.exec(name);
    if (m) level = Number(m[1]);
    else if (name === 'title') level = 1;
    else if (p.props.outlineLevel !== null && p.props.outlineLevel < 9) level = p.props.outlineLevel + 1;
    if (!level) return 0;
    // A long paragraph that carries an outline level is body text listed in the navigation pane.
    const length = segs.reduce((n, s) => n + s.text.length, 0);
    return length > 400 ? 0 : Math.min(level, 6);
  }

  private isCode(p: Paragraph, segs: Seg[], name: string): boolean {
    if (/^(html preformatted|code|source code|code block)$/.test(name)) return true;
    // Word's Plain Text style is Courier, but documents restyle it: it is code only in a monospaced face.
    if (name === 'plain text') return p.content.every((c) => c.type !== 'text' || !c.text.trim() || isMonospace(face(c.props, c.text).family));
    if (!this.codeFonts || p.numbering) return false;
    let any = false;
    for (const s of segs) {
      if (s.br || !s.text.trim()) continue;
      if (!(s.fmt & CODE) || s.link) return false;
      any = true;
    }
    return any;
  }

  /** The LaTeX of a paragraph that holds only displayed equations. */
  private displayMath(p: Paragraph): string | undefined {
    let math: string[] | undefined;
    for (const c of p.content) {
      if (c.type === 'math' && c.display) (math ??= []).push(mathToLatex(c.math));
      else if (c.type === 'text' && !c.text.trim()) continue;
      else if (c.type === 'break' || c.type === 'bookmark' || c.type === 'renderedPageBreak') continue;
      else return undefined;
    }
    return math?.join(' \\\\\n');
  }

  // -------------------------------------------------------------------------------------------
  // Lists

  /** The number or bullet Word shows, with symbol-font bullets as the characters they draw. */
  private labelText(label: NumberingLabel, bare = false): string {
    let text = label.picture && !label.text.trim() ? '•' : this.mapSymbols(label.text, label.props);
    if (bare) return text;
    if (label.suffix === 'tab') text += '\t';
    else if (label.suffix === 'space') text += ' ';
    return text;
  }

  private listMarker(label: NumberingLabel): string {
    const text = this.labelText(label, true).trim();
    if (label.format === 'bullet' || !text) return '- ';
    const m = /^([0-9]{1,9})([.)])$/.exec(text);
    if (m && !/^0\d/.test(m[1])) return `${m[1]}${m[2]} `;
    // Letters, Roman numerals and outline numbers (1.2.3) have no Markdown form: keep them as text.
    return `- ${escapeMd(text, false, this.escapeDollars)} `;
  }

  // -------------------------------------------------------------------------------------------
  // Paragraph content

  private inlines(p: Paragraph): { segs: Seg[]; extras: Extra[] } {
    const segs: Seg[] = [];
    const extras: Extra[] = [];
    const push = (text: string, props: RunProps, link?: Link) => {
      if (!text) return;
      const fmt = this.md ? this.fmt(props, text) : 0;
      const last = segs[segs.length - 1];
      if (last && !last.raw && !last.br && last.fmt === fmt && last.link === link) last.text += text;
      else segs.push({ text, fmt, link });
    };
    // Ruby: Word prints the reading above the base text. Markdown keeps it as HTML ruby, which
    // viewers without ruby show as the base followed by the reading in brackets, as plain text does.
    let ruby: { text: string } | undefined;
    const endRuby = () => {
      if (!ruby) return;
      const reading = ruby.text.trim();
      if (this.md) segs.push({ text: `<rp>（</rp><rt>${escapeHtml(reading)}</rt><rp>）</rp></ruby>`, fmt: 0, raw: true });
      else segs.push({ text: `（${reading}）`, fmt: 0 });
      ruby = undefined;
    };
    for (const c of p.content) {
      if (c.type !== 'text' || c.ruby !== ruby) endRuby();
      switch (c.type) {
        case 'text':
          if (c.ruby && c.ruby !== ruby && c.ruby.text.trim()) {
            ruby = c.ruby;
            if (this.md) segs.push({ text: '<ruby>', fmt: 0, raw: true });
          }
          push(this.cleanText(c.text, c.props), c.props, external(c.link));
          break;
        case 'tab':
        case 'ptab':
          push(this.md ? ' ' : '\t', c.props, c.type === 'tab' ? external(c.link) : undefined);
          break;
        case 'break':
          segs.push({ text: '\n', fmt: 0, br: true });
          break;
        case 'field':
          // Page numbers and references show Word's last result.
          push(c.cached, c.props);
          break;
        case 'noteRef': {
          const label = this.noteLabel(c.noteType, c.id, c.customMark !== undefined);
          if (this.md) segs.push({ text: `[^${label}]`, fmt: 0, raw: true });
          else if (c.customMark === undefined) push(`[${label}]`, c.props);
          break;
        }
        // A note's own number (w:footnoteRef) is left out: the note's label carries it.
        case 'math':
          if (this.md) segs.push({ text: `$${mathToLatex(c.math)}$`, fmt: 0, raw: true });
          else push(mathToText(c.math), c.props);
          break;
        case 'drawing':
          this.drawing(c.drawing, segs, extras, true);
          break;
      }
    }
    endRuby();
    return { segs, extras };
  }

  private fmt(props: RunProps, text: string): number {
    // Arabic, Hebrew and other complex scripts take their face from the run's complex-script
    // font. Word draws them bold or italic from either set of properties: Word's PDFs show Arabic
    // bold with only w:b set.
    const font = face(props, text);
    let f = 0;
    if (props.bold || font.bold) f |= BOLD;
    if (props.italic || font.italic) f |= ITALIC;
    if (props.strike || props.dstrike) f |= STRIKE;
    if (props.vertAlign === 'superscript') f |= SUP;
    else if (props.vertAlign === 'subscript') f |= SUB;
    if (this.codeFonts && isMonospace(font.family)) f |= CODE;
    return f;
  }

  private drawing(d: Drawing, segs: Seg[], extras: Extra[], top: boolean): void {
    if (d.hidden) return;
    if (d.hr) {
      extras.push({ rule: true });
      return;
    }
    if (d.kind === 'image' && d.image && this.md && this.options.includeImages !== false) {
      const target = d.image.id.replace(/#.*$/, '');
      const alt = (d.alt ?? '').replace(/\s+/g, ' ').trim().replace(/[\\[\]]/g, '\\$&');
      const md = `![${alt}](${encodeUrl(target)})`;
      if (top) segs.push({ text: md, fmt: 0, raw: true });
      else extras.push({ image: md });
    }
    if (d.textBox) extras.push({ blocks: d.textBox.blocks });
    if (d.textPath?.text.trim()) extras.push({ text: d.textPath.text });
    if (d.chart) extras.push({ chart: d.chart });
    for (const kid of d.children ?? []) this.drawing(kid, segs, extras, false);
  }

  /** Text as Word shows it: symbol-font characters mapped to Unicode, soft hyphens dropped. */
  private cleanText(text: string, props: RunProps): string {
    text = this.mapSymbols(text, props);
    if (text.includes('\u00ad')) text = text.replace(/\u00ad/g, '');
    if (text.includes('\u2011')) text = text.replace(/\u2011/g, '-');
    // A tab character in the text is a tab, which Markdown shows as a space.
    if (this.md && text.includes('\t')) text = text.replace(/\t/g, ' ');
    return text;
  }

  private symbolRun(props: RunProps): boolean {
    let v = this.symbolCache.get(props);
    if (v === undefined) {
      v = !!(symbolFontKind(props.fontAscii) || symbolFontKind(props.fontHAnsi) || symbolFontKind(props.fontEastAsia) || symbolFontKind(props.fontCs));
      this.symbolCache.set(props, v);
    }
    return v;
  }

  private mapSymbols(text: string, props: RunProps): string {
    const symbolRun = this.symbolRun(props);
    if (!symbolRun && !PUA_SYMBOL.test(text)) return text;
    let out = '';
    for (const ch of text) {
      const cp = ch.codePointAt(0)!;
      const kind: SymbolFontKind | null = symbolRun ? symbolFontKind(fontFor(cp, props).family) : null;
      const pua = cp >= 0xf020 && cp <= 0xf0ff;
      if (!kind && !pua) {
        out += ch;
        continue;
      }
      const mapped = mapSymbolChar(kind ?? 'symbol', cp);
      if (mapped && mapped !== cp) out += String.fromCodePoint(mapped);
      else if (!pua) out += ch;
      // A private-use symbol code with no Unicode equivalent shows nothing readable.
    }
    return out;
  }

  // -------------------------------------------------------------------------------------------
  // Plain text

  private plain(segs: Seg[]): string {
    let s = '';
    for (const seg of segs) s += seg.text;
    return s;
  }

  // -------------------------------------------------------------------------------------------
  // Markdown inline formatting

  /** Segments as Markdown: emphasis markers nest properly and never sit against whitespace. */
  private renderSegs(segs: Seg[], inTable = false): string {
    segs = autolinks(segs);
    let out = '';
    for (let i = 0; i < segs.length; ) {
      const link = segs[i].link;
      let j = i + 1;
      while (j < segs.length && segs[j].link === link) j++;
      const group = segs.slice(i, j);
      // Brackets inside a link's text are escaped so they cannot end it early.
      const inner = this.formatted(group, inTable, !!link?.url);
      const url = link?.url ? link.url + (link.anchor ? '#' + link.anchor : '') : undefined;
      if (url && inner.trim() && /^[a-z][a-z0-9+.-]*:/i.test(url)) {
        const lead = /^\s*/.exec(inner)![0];
        const trail = /\s*$/.exec(inner)![0];
        const label = inner.trim();
        const plain = group.map((s) => s.text).join('').trim();
        if (plain === url && /^(https?|mailto|ftp):/i.test(url) && !/[<>\s]/.test(url)) out += `${lead}<${url}>${trail}`;
        else out += `${lead}[${label}](${encodeUrl(url)})${trail}`;
      } else out += inner;
      i = j;
    }
    // Tidy lines: no leading indentation (it would read as code), no trailing spaces, and
    // characters that would start a block construct escaped.
    // Breaks at the start or end of a paragraph (a page break before its text) show no line.
    const lines = out.split('\n').map((l) => escapeLineStart(l.replace(/^[ \t\u00a0]+/, '').replace(/[ \t]+$/, '').replace(/[ \t]+(\uE000)$/, '$1')));
    const blank = (l: string) => !l.replace(/\uE000$/, '').trim();
    while (lines.length && blank(lines[lines.length - 1])) lines.pop();
    while (lines.length && blank(lines[0])) lines.shift();
    if (lines.length) lines[lines.length - 1] = lines[lines.length - 1].replace(/\uE000$/, '');
    return inTable ? lines.map((l) => l.replace(/\uE000$/, '')).join('<br>') : lines.join('\n').replace(/\uE000/g, '\\');
  }

  private formatted(segs: Seg[], inTable: boolean, inLink = false): string {
    segs = flanking(segs);
    const MARKS: Array<[number, string, string]> = [
      [BOLD, '**', '**'],
      [ITALIC, '*', '*'],
      [STRIKE, '~~', '~~'],
      [SUP, '<sup>', '</sup>'],
      [SUB, '<sub>', '</sub>'],
    ];
    const markOf = (bit: number) => MARKS.find((m) => m[0] === bit)!;
    const stack: number[] = [];
    let out = '';
    const close = (n: number) => {
      // Whitespace before a closing marker moves after it.
      // CommonMark counts every Unicode space (no-break, ideographic...) as whitespace.
      const ws = /(?:\uE000\n|\s)*$/.exec(out)![0];
      out = out.slice(0, out.length - ws.length);
      while (stack.length > n) out += markOf(stack.pop()!)[2];
      out += ws;
    };
    const visible = segs.map((s) => !s.br && !!s.text.trim());
    for (let i = 0; i < segs.length; i++) {
      const seg = segs[i];
      // Whitespace keeps the emphasis that continues past it but opens none.
      let want = seg.fmt & ~CODE;
      if (!visible[i]) {
        let next = 0;
        for (let k = i + 1; k < segs.length; k++) if (visible[k]) { next = segs[k].fmt; break; }
        want = stack.reduce((m, b) => m | b, 0) & next;
      }
      if (seg.raw && !seg.auto) want = stack.reduce((m, b) => m | b, 0) & want;
      const keep = stack.findIndex((b) => !(want & b));
      if (keep >= 0) close(keep);
      // A line break is a backslash at the end of the line, marked until the lines are tidied.
      const text = seg.br ? '\uE000\n' : seg.raw ? seg.text : seg.fmt & CODE && visible[i] ? codeSpan(seg.text, inTable) : escapeMd(seg.text, inTable, this.escapeDollars, inLink);
      const lead = seg.br || seg.raw ? '' : /^\s*/.exec(text)![0];
      out += lead;
      // Open the missing markers, the longest-lasting outermost.
      const open = MARKS.map((m) => m[0]).filter((b) => want & b && !stack.includes(b));
      if (open.length) {
        const span = (b: number) => {
          let k = i;
          while (k < segs.length && (segs[k].fmt & b || !visible[k])) k++;
          return k;
        };
        open.sort((a, b) => span(b) - span(a));
        for (const b of open) {
          stack.push(b);
          out += markOf(b)[1];
        }
      }
      out += text.slice(lead.length);
    }
    close(0);
    return out;
  }

  // -------------------------------------------------------------------------------------------
  // Tables

  private table(t: Table, out: Chunk[]): void {
    const rows = t.rows.filter((r) => r.cells.length);
    if (!rows.length) return;
    if (!this.md) {
      for (const row of rows) {
        const cells: string[] = [];
        for (let k = 0; k < row.gridBefore; k++) cells.push('');
        for (const cell of row.cells) {
          cells.push(cell.vMerge === 'continue' ? '' : this.cellText(cell));
          for (let k = 1; k < cell.gridSpan; k++) cells.push('');
        }
        while (cells.length && cells[cells.length - 1] === '') cells.pop();
        out.push({ text: cells.join('\t'), kind: 'block' });
      }
      return;
    }
    // Rows with nothing in them only space the table out.
    const filled = rows.filter((r) => r.cells.some((c) => c.vMerge !== 'continue' && hasContent(c.content)));
    if (!filled.length) return;
    // A single cell is a box around its content, not tabular data.
    if (filled.length === 1 && filled[0].cells.length === 1) {
      this.blocks(filled[0].cells[0].content, out);
      return;
    }
    const nested = rows.some((r) => r.cells.some((c) => c.content.some((b) => b.type === 'table')));
    out.push({ text: nested ? this.htmlTable(t) : this.gfmTable(filled), kind: 'block' });
  }

  /** A cell's text on one line: its paragraphs separated by spaces. */
  private cellText(cell: TableCell): string {
    const chunks: Chunk[] = [];
    this.blocks(cell.content, chunks);
    return chunks
      .map((c) => c.text.replace(/[\t\n]+/g, ' ').trim())
      .filter(Boolean)
      .join(' ');
  }

  private cellMarkdown(cell: TableCell): string {
    const parts: string[] = [];
    const visit = (blocks: Block[]) => {
      for (const b of blocks) {
        if (b.type === 'table') {
          for (const r of b.rows) for (const c of r.cells) visit(c.content);
          continue;
        }
        const { segs, extras } = this.inlines(b);
        const label = b.numbering && b.numbering.text.trim() ? escapeMd(this.labelText(b.numbering, true).trim(), true, this.escapeDollars) + ' ' : '';
        const math = this.displayMath(b);
        const text = math !== undefined ? `$${math.replace(/\s*\\\\\n/g, ' \\\\ ')}$` : this.renderSegs(segs, true);
        if (text.trim()) parts.push(label + text);
        for (const e of extras) {
          if ('blocks' in e) visit(e.blocks);
          else if ('text' in e) parts.push(escapeMd(e.text, true, this.escapeDollars));
          else if ('image' in e) parts.push(e.image);
          else if ('chart' in e) {
            const title = chartTitle(e.chart);
            if (title) parts.push(escapeMd(title, true, this.escapeDollars));
          }
        }
      }
    };
    visit(cell.content);
    return parts.join('<br>');
  }

  private gfmTable(rows: Table['rows']): string {
    const grid: string[][] = rows.map((row) => {
      const cells: string[] = [];
      for (let k = 0; k < row.gridBefore; k++) cells.push('');
      for (const cell of row.cells) {
        cells.push(cell.vMerge === 'continue' ? '' : this.cellMarkdown(cell));
        for (let k = 1; k < cell.gridSpan; k++) cells.push('');
      }
      for (let k = 0; k < row.gridAfter; k++) cells.push('');
      return cells;
    });
    let width = Math.max(...grid.map((r) => r.length));
    // Trailing columns empty in every row are grid padding.
    while (width > 1 && grid.every((r) => !r[width - 1])) width--;
    const line = (cells: string[]) => '| ' + Array.from({ length: width }, (_, k) => cells[k] ?? '').join(' | ') + ' |';
    const lines = [line(grid[0]), '|' + ' --- |'.repeat(width), ...grid.slice(1).map(line)];
    return lines.map((l) => l.replace(/ {2,}\|/g, ' |')).join('\n');
  }

  private htmlTable(t: Table): string {
    const rows = t.rows.filter((r) => r.cells.length);
    // Each cell's starting grid column, to find the cells a vertical merge continues into.
    const starts = rows.map((r) => {
      let col = r.gridBefore;
      return r.cells.map((c) => {
        const at = col;
        col += c.gridSpan;
        return at;
      });
    });
    let html = '<table>\n';
    rows.forEach((row, r) => {
      html += '<tr>';
      row.cells.forEach((cell, k) => {
        if (cell.vMerge === 'continue') return;
        let rowspan = 1;
        if (cell.vMerge === 'restart') {
          for (let n = r + 1; n < rows.length; n++) {
            const idx = starts[n].indexOf(starts[r][k]);
            if (idx < 0 || rows[n].cells[idx].vMerge !== 'continue') break;
            rowspan++;
          }
        }
        const tag = row.isHeader ? 'th' : 'td';
        const attrs = (cell.gridSpan > 1 ? ` colspan="${cell.gridSpan}"` : '') + (rowspan > 1 ? ` rowspan="${rowspan}"` : '');
        html += `<${tag}${attrs}>${this.cellHtml(cell)}</${tag}>`;
      });
      html += '</tr>\n';
    });
    return html + '</table>';
  }

  private cellHtml(cell: TableCell): string {
    const parts: string[] = [];
    for (const b of cell.content) {
      if (b.type === 'table') {
        parts.push(this.htmlTable(b));
        continue;
      }
      const { segs, extras } = this.inlines(b);
      const label = b.numbering && b.numbering.text.trim() ? escapeHtml(this.labelText(b.numbering, true).trim()) + ' ' : '';
      const text = segsToHtml(segs);
      if (text.trim()) parts.push(label + text);
      for (const e of extras) {
        if ('blocks' in e) parts.push(this.cellHtml({ ...cell, content: e.blocks }));
        else if ('text' in e) parts.push(escapeHtml(e.text));
      }
    }
    return parts.join('<br>');
  }

  // -------------------------------------------------------------------------------------------
  // Charts

  private chart(chart: Chart, out: Chunk[]): void {
    const title = chartTitle(chart);
    if (title) out.push({ text: this.md ? `**${escapeMd(title, false, this.escapeDollars)}**` : title, kind: 'block' });
    const series = chart.plots.flatMap((p) => p.series);
    if (!series.length) return;
    const count = Math.max(...series.map((s) => Math.max(s.values.length, s.categories.length, s.xValues?.length ?? 0)));
    if (!count || count > 1000) return;
    const first = series.find((s) => s.categories.length) ?? series[0];
    const scatter = series.some((s) => s.xValues?.length);
    // Numeric and date categories show in the category axis's format, unless it takes the data's.
    const plot = chart.plots.find((p) => p.series.includes(first));
    const catAxis = chart.axes.find((a) => plot?.axes.includes(a.id) && a.kind !== 'val');
    const catFormat = catAxis && !catAxis.sourceLinked ? catAxis.format : (first.categoryFormat ?? 'General');
    const header = [scatter ? 'X' : '', ...series.map((s, i) => s.name || `Series ${i + 1}`)];
    const rows: string[][] = [];
    for (let i = 0; i < count; i++) {
      const num = scatter ? first.xValues?.[i] : first.categoryNumbers?.[i];
      const cat = num !== undefined && num !== null ? value(num, catFormat) : (first.categories[i] ?? String(i + 1));
      rows.push([cat, ...series.map((s) => value(s.values[i], s.format))]);
    }
    if (this.md) {
      const esc = (s: string) => escapeMd(s, true, this.escapeDollars);
      const lines = ['| ' + header.map(esc).join(' | ') + ' |', '|' + ' --- |'.repeat(header.length), ...rows.map((r) => '| ' + r.map(esc).join(' | ') + ' |')];
      out.push({ text: lines.join('\n').replace(/ {2,}\|/g, ' |'), kind: 'block' });
    } else {
      out.push({ text: header.join('\t'), kind: 'block' });
      for (const r of rows) out.push({ text: r.join('\t'), kind: 'block' });
    }
  }

  // -------------------------------------------------------------------------------------------
  // Notes

  private noteLabel(type: 'footnote' | 'endnote', id: string, custom: boolean): string {
    const key = type + ':' + id;
    let label = this.noteLabels.get(key);
    if (label !== undefined) return label;
    if (custom) label = `note${++this.customNotes}`;
    else if (type === 'endnote') label = formatNumber(this.endnoteCounter++, this.doc.settings.endnoteFormat ?? 'lowerRoman');
    else {
      if (this.footnoteCounter === undefined) this.footnoteCounter = this.section.footnoteStart;
      label = formatNumber(this.footnoteCounter++, this.section.footnoteFormat);
    }
    if (this.md) {
      if (!NOTE_LABEL.test(label) || this.usedLabels.has(label)) label = `${type === 'endnote' ? 'e' : 'n'}${this.noteOrder.length + 1}`;
      this.usedLabels.add(label);
    }
    this.noteLabels.set(key, label);
    this.noteOrder.push({ type, id, label, custom });
    return label;
  }

  private notes(out: Chunk[]): void {
    // Footnotes first, then endnotes, each in reference order. Notes may reference further notes.
    // In plain text a blank line sets them off from the body.
    const done = new Set<string>();
    if (!this.md && this.noteOrder.length) out.push({ text: '', kind: 'block' });
    for (const type of ['footnote', 'endnote'] as const) {
      for (let i = 0; i < this.noteOrder.length; i++) {
        const ref = this.noteOrder[i];
        if (ref.type !== type || done.has(type + ref.id)) continue;
        done.add(type + ref.id);
        const note = (type === 'footnote' ? this.doc.footnotes : this.doc.endnotes).get(ref.id);
        if (!note) continue;
        const chunks: Chunk[] = [];
        this.blocks(note.blocks, chunks);
        const text = chunks.map((c) => c.text).filter((t) => t.trim());
        if (!text.length) text.push('');
        text[0] = text[0].replace(/^\s+/, '');
        const custom = ref.custom;
        if (this.md) {
          const body = text.join('\n\n').split('\n').map((l, k) => (k && l ? '    ' + l : l)).join('\n');
          out.push({ text: `[^${ref.label}]: ${body}`, kind: 'block' });
        } else out.push({ text: (custom ? '' : `[${ref.label}] `) + text.join('\n'), kind: 'block' });
      }
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Helpers

function longestRun(s: string, ch: string): number {
  let best = 0;
  let run = 0;
  for (const c of s) {
    run = c === ch ? run + 1 : 0;
    if (run > best) best = run;
  }
  return best;
}

function codeSpan(text: string, inTable: boolean): string {
  const lead = /^\s*/.exec(text)![0];
  const trail = /\s*$/.exec(text)![0];
  let core = text.slice(lead.length, text.length - trail.length);
  if (inTable) core = core.replace(/\|/g, '\\|');
  const fence = '`'.repeat(longestRun(core, '`') + 1);
  const pad = core.startsWith('`') || core.endsWith('`') ? ' ' : '';
  return `${lead}${fence}${pad}${core}${pad}${fence}${trail}`;
}

/** Escape characters Markdown would read as formatting. */
export function escapeMd(s: string, inTable: boolean, dollars: boolean, brackets = false): string {
  s = s.replace(brackets ? /[\\`*~[\]]/g : /[\\`*~]|\[(?=\^)|\](?=[([:])/g, '\\$&');
  // An underscore between letters or digits cannot start or end emphasis; every other one is escaped.
  if (s.includes('_')) s = s.replace(/(?<![\p{L}\p{N}])_|_(?![\p{L}\p{N}])/gu, '\\_');
  if (s.includes('<')) s = s.replace(/<(?=[A-Za-z/!?])/g, '\\<');
  if (s.includes('&')) s = s.replace(/&(?=#?[A-Za-z0-9]+;)/g, '\\&');
  if (dollars) s = s.replace(/\$/g, '\\$');
  if (inTable) s = s.replace(/\|/g, '\\|');
  return s;
}

/** Escape what would make a line a heading, list item, quote or rule. */
function escapeLineStart(line: string): string {
  const m = /^(#{1,6}(?=\s|$)|>|[-+](?=\s|$)|=+\s*$|-{2,}\s*$|\d{1,9}(?=[.)](?:\s|$)))/.exec(line);
  if (!m) return line;
  if (/^\d/.test(m[0])) return m[0] + '\\' + line.slice(m[0].length);
  return '\\' + line;
}

function indentLines(text: string, indent: number): string {
  if (!text.includes('\n')) return text;
  const pad = ' '.repeat(indent);
  return text.split('\n').map((l, i) => (i && l ? pad + l : l)).join('\n');
}

function encodeUrl(url: string): string {
  return url.replace(/[\s()<>]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'));
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function segsToHtml(segs: Seg[]): string {
  let out = '';
  for (const s of segs) {
    if (s.br) {
      out += '<br>';
      continue;
    }
    let t = s.raw ? s.text : escapeHtml(s.text);
    if (s.fmt & CODE) t = `<code>${t}</code>`;
    if (s.fmt & SUP) t = `<sup>${t}</sup>`;
    if (s.fmt & SUB) t = `<sub>${t}</sub>`;
    if (s.fmt & STRIKE) t = `<del>${t}</del>`;
    if (s.fmt & ITALIC) t = `<em>${t}</em>`;
    if (s.fmt & BOLD) t = `<strong>${t}</strong>`;
    if (s.link?.url && /^[a-z][a-z0-9+.-]*:/i.test(s.link.url)) t = `<a href="${escapeHtml(s.link.url).replace(/"/g, '&quot;')}">${t}</a>`;
    out += t;
  }
  return out.replace(/\s+/g, ' ');
}

/** Whether blocks hold anything to show: text, an equation, a picture or a nested table. */
function hasContent(blocks: Block[]): boolean {
  return blocks.some((b) =>
    b.type === 'table'
      ? b.rows.some((r) => r.cells.some((c) => hasContent(c.content)))
      : !!b.numbering?.text.trim() || b.content.some((c) => (c.type === 'text' && c.text.trim()) || c.type === 'math' || c.type === 'drawing' || (c.type === 'field' && c.cached.trim())),
  );
}

/** A link to somewhere outside the document; links to bookmarks have no Markdown form. */
function external(link: Link | undefined): Link | undefined {
  return link?.url ? link : undefined;
}

function chartTitle(chart: Chart): string {
  return (chart.title?.paragraphs ?? []).map((p) => p.map((r) => r.text).join('')).join(' ').trim();
}

function value(v: number | null | undefined, format: string): string {
  return v === null || v === undefined ? '' : formatChartNumber(v, format);
}

/** Text or Markdown from a parsed document. */
export function documentToText(doc: DocxDocument, options: TextOptions = {}): string {
  return new TextExtractor(doc, options).run();
}
