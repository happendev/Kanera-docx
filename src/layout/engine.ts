import type { Block, DocxDocument, Drawing, FieldInline, Paragraph, Section, Table } from '../docx/model.js';
import { formatNumber } from '../docx/numbering.js';
import type { FontRegistry } from '../fonts/registry.js';
import {
  gapBetween,
  spaceAboveTable,
  hasDecorations,
  layoutBlocks,
  measureTable,
  paraDecorations,
  shadedSpaceAfter,
  sideBorderRunOn,
  findSlot,
  gapBeside,
  textHeight,
  paraFrame,
  renderSegment,
  spaceAfter,
  wrapLine,
  type BoxEnv,
  type BoxItem,
  type BoxResult,
  type BoxWrap,
  type CellLayout,
  type Exclusion,
  type MergeProgress,
  type RowPiece,
  type TableLayout,
} from './box.js';
import { renderDrawing } from './drawing.js';
import { ParagraphLayout, type Line } from './paragraph.js';
import { translateOps, type DrawOp, type FloatOp, type Page } from './types.js';

const EPS = 0.01;

export interface LayoutStats {
  /** Word's own page breaks recorded in the file (w:lastRenderedPageBreak) in body paragraphs. */
  renderedBreaks: number;
  /** How many of those start a page or column in this layout too. */
  matchedBreaks: number;
  /** Where each recorded break landed here (page index from 0, and whether it starts a column). */
  breakLog: Array<{ page: number; top: boolean; y: number; block: number }>;
}

export interface LayoutResult {
  pages: Page[];
  stats: LayoutStats;
  warnings: string[];
}

/** What positions a floating object on a page. */
type PageFrame = Pick<PageState, 'section' | 'width' | 'height' | 'marginLeft' | 'marginRight'>;

interface FloatEntry {
  z: number;
  behind: boolean;
  ops: DrawOp[];
}

interface Column {
  x: number;
  w: number;
}

interface FootnotePiece {
  id: string;
  box: BoxResult;
  /** Coordinates in the original note, retained when a note continues on another page. */
  start: number;
  end: number;
  offset: number;
}

interface PageState {
  section: Section;
  sectionIndex: number;
  firstOfSection: boolean;
  /** Displayed page number. */
  number: number;
  width: number;
  height: number;
  marginLeft: number;
  marginRight: number;
  top: number;
  bottom: number;
  cols: Column[];
  /** y where the current set of columns starts (below an earlier continuous section). */
  colTop: number;
  body: DrawOp[];
  floats: FloatEntry[];
  exclusions: Exclusion[];
  notes: FootnotePiece[];
  overflow: FootnotePiece[];
  continuation: boolean;
  /** An empty page that only keeps odd and even page numbers apart: it prints nothing. */
  blank?: boolean;
  noteHeight: number;
  /** Lowest content y per column of the current section on this page. */
  colBottoms: number[];
  /** Sections that place content on this page with their column layout (for separators). */
  colSets: Array<{ section: Section; cols: Column[]; top: number; bottom: number }>;
}

interface Checkpoint {
  pages: number;
  body: number;
  floats: number;
  exclusions: number;
  notes: number;
  overflow: number;
  noteHeight: number;
  col: number;
  y: number;
  atTop: boolean;
  topKind: 'hard' | 'soft' | 'section';
  prevPara: Paragraph | null;
  colBottoms: number[];
  colTop: number;
  cols: Column[];
  colSets: number;
  renderedBreaks: number;
  matchedBreaks: number;
  // Break records are append-only until the paginator rolls back.
  breakLogLength: number;
}

const FIELD_FORMATS: Record<string, string> = {
  roman: 'lowerRoman',
  ROMAN: 'upperRoman',
  alphabetic: 'lowerLetter',
  ALPHABETIC: 'upperLetter',
  Arabic: 'decimal',
  arabic: 'decimal',
};

/** Float positions that don't depend on where the anchor paragraph falls on its page. */
const PAGE_FIXED_H = new Set(['page', 'margin', 'leftMargin', 'rightMargin', 'insideMargin', 'outsideMargin']);
const PAGE_FIXED_V = new Set(['page', 'margin', 'topMargin', 'bottomMargin', 'insideMargin', 'outsideMargin']);

export function layoutDocument(doc: DocxDocument, fonts: FontRegistry): LayoutResult {
  // Field width can change wrapping, so feed measured page values back into normal layout.
  const values = new Map<FieldInline, string>();
  const early = new Map<Drawing, number>();
  for (let pass = 0; ; pass++) {
    const paginator = new Paginator(doc, fonts, values, early);
    const result = paginator.run();
    if (!paginator.fieldsChanged && !paginator.earlyChanged) return result;
    if (pass === 7) {
      result.warnings.push('Page fields did not stabilize after repeated pagination');
      return result;
    }
  }
}

class Paginator {
  fieldsChanged = false;
  /** A float was found that belongs on its page from the top (see placeFloats). */
  earlyChanged = false;
  private measuredFields = new Map<FieldInline, { value: string; section: Section }>();
  /** The page each bookmark landed on, for page references. */
  private anchorPages = new Map<string, { number: number; format: string }>();
  private pages: PageState[] = [];
  private page!: PageState;
  private col = 0;
  private y = 0;
  /** Nothing has been placed in the current column yet. */
  private atTop = true;
  /** Page count right after a page break that ended a paragraph (see startSection). */
  private breakOpenedPage = -1;
  /** Floating tables (w:tblpPr), as floats anchored to the paragraph that follows them. */
  private tableFloats = new Map<Paragraph, Drawing[]>();
  /** The stand-in drawings of floating tables, rebuilt on every pass. */
  private tableFloatDrawings = new WeakSet<Drawing>();
  private floatedTables = new Set<Table>();
  /** Why the current column started: suppress space-before only after automatic breaks. */
  private topKind: 'hard' | 'soft' | 'section' = 'hard';
  private prevPara: Paragraph | null = null;
  /** Index of the body block being placed (diagnostics). */
  private blockIndex = 0;
  private sectionBlockBase = 0;
  private section!: Section;
  private sectionIndex = 0;
  /** The section after the one being laid out. */
  private nextSection: Section | null = null;
  private sectionStartPage = 0;
  private restartNumber: number | null = null;
  private stats: LayoutStats = { renderedBreaks: 0, matchedBreaks: 0, breakLog: [] };
  private env: BoxEnv;
  private noteNumbers = new Map<string, number>();
  private noteBoxes = new Map<string, BoxResult>();
  private separatorBox: BoxResult | null = null;
  private endnotesUsed: string[] = [];
  private warnings: string[] = [];
  /** Page whose fields are being laid out (null in the body flow, which uses the current page). */
  private fieldPage: { number: number; format: string; total: number | null; sectionTotal: number | null } | null = null;

  constructor(
    private doc: DocxDocument,
    private fonts: FontRegistry,
    private fieldValues: Map<FieldInline, string>,
    /** Floats placed as their page opens, by page index (see placeFloats). */
    private early: Map<Drawing, number>,
  ) {
    this.env = {
      fonts,
      settings: doc.settings,
      compatibilityMode: doc.settings.compatibilityMode,
      fieldText: (f) => this.fieldText(f),
      noteRefText: (type, id) => this.noteRefText(type, id),
      drawInline: (d, x, y) => renderDrawing(d, x, y, this.env),
    };
    this.numberNotes();
  }

  // -------------------------------------------------------------------------------------------
  // Fields and note numbers

  private fieldText(f: FieldInline): string {
    const ctx = this.fieldPage;
    const page = ctx?.number ?? this.page?.number ?? 1;
    const sectFmt = ctx?.format ?? this.section?.pageNumberFormat ?? 'decimal';
    const fmt = f.format ? (FIELD_FORMATS[f.format] ?? sectFmt) : sectFmt;
    if (!ctx) {
      const value = this.fieldValues.get(f) ?? (f.field === 'page' ? formatNumber(page, fmt) : f.cached || '1');
      this.measuredFields.set(f, { value, section: this.section });
      return value;
    }
    if (f.field === 'page') return formatNumber(page, fmt);
    if (f.field === 'numpages') return ctx?.total != null ? formatNumber(ctx.total, f.format ? fmt : 'decimal') : this.fieldValues.get(f) ?? (f.cached || '1');
    if (f.field === 'sectionpages') return ctx?.sectionTotal != null ? formatNumber(ctx.sectionTotal, f.format ? fmt : 'decimal') : this.fieldValues.get(f) ?? (f.cached || '1');
    if (f.field === 'pageref') {
      const at = this.anchorPages.get(f.bookmark ?? '');
      return at ? formatNumber(at.number, at.format) : this.fieldValues.get(f) ?? (f.cached || '1');
    }
    return f.cached;
  }

  private noteRefText(type: 'footnote' | 'endnote', id: string): string {
    const n = this.noteNumbers.get(`${type}:${id}`) ?? 1;
    if (type === 'endnote') return formatNumber(n, 'lowerRoman');
    return formatNumber(n, this.section?.footnoteFormat ?? 'decimal');
  }

  /** Notes are numbered in reference order through the body. */
  private numberNotes(): void {
    let fn = this.doc.finalSection.footnoteStart || 1;
    let en = 1;
    const visit = (blocks: Block[]) => {
      for (const b of blocks) {
        if (b.type === 'table') {
          for (const row of b.rows) for (const cell of row.cells) visit(cell.content);
          continue;
        }
        for (const inl of b.content) {
          if (inl.type !== 'noteRef' || inl.customMark !== undefined) continue;
          const key = `${inl.noteType}:${inl.id}`;
          if (this.noteNumbers.has(key)) continue;
          this.noteNumbers.set(key, inl.noteType === 'footnote' ? fn++ : en++);
        }
      }
    };
    visit(this.doc.body);
  }

  private noteBox(type: 'footnote' | 'endnote', id: string, width: number): BoxResult {
    const key = `${type}:${id}`;
    let box = this.noteBoxes.get(key);
    if (!box) {
      const note = (type === 'footnote' ? this.doc.footnotes : this.doc.endnotes).get(id);
      box = note ? layoutBlocks(note.blocks, width, { ...this.env, noteNumberText: this.noteRefText(type, id) }) : { height: 0, items: [] };
      this.noteBoxes.set(key, box);
    }
    return box;
  }

  private footnoteSeparator(width: number): BoxResult {
    if (!this.separatorBox) {
      const blocks = this.doc.footnoteSeparator;
      this.separatorBox = blocks ? layoutBlocks(blocks, width, this.env) : { height: 12, items: [{ y: 0, height: 12, ops: [{ kind: 'line', x1: 0, y1: 6, x2: 144, y2: 6, width: 0.5, color: '000000' }] }] };
    }
    return this.separatorBox;
  }

  // -------------------------------------------------------------------------------------------
  // Run

  run(): LayoutResult {
    const groups = this.sections();
    for (let i = 0; i < groups.length; i++) {
      const { section, blocks } = groups[i];
      this.nextSection = groups[i + 1]?.section ?? null;
      this.startSection(section, i);
      this.placeBlocks(blocks);
      this.sectionBlockBase += blocks.length;
    }
    this.placeEndnotes();
    this.flushContinuations();
    const pages = this.finish();
    return { pages, stats: this.stats, warnings: this.warnings };
  }

  private sections(): Array<{ section: Section; blocks: Block[] }> {
    const out: Array<{ section: Section; blocks: Block[] }> = [];
    let cur: Block[] = [];
    for (const b of this.doc.body) {
      cur.push(b);
      if (b.type === 'paragraph' && b.sectPr) {
        out.push({ section: b.sectPr, blocks: cur });
        cur = [];
      }
    }
    // A trailing empty group (document ends on a section-break paragraph) adds nothing.
    if (cur.length || out.length === 0) out.push({ section: this.doc.finalSection, blocks: cur });
    return out;
  }

  private startSection(section: Section, index: number): void {
    const prev = this.section;
    this.section = section;
    this.sectionIndex = index;
    this.env.gridPitch = section.gridType === 'lines' || section.gridType === 'linesAndChars' ? section.gridLinePitch : undefined;
    this.env.charGrid = section.gridType === 'linesAndChars' && section.gridCharSpace ? section.gridCharSpace : undefined;
    if (section.pageNumberStart !== null) this.restartNumber = section.pageNumberStart;
    if (index === 0 || !this.page) {
      this.newPage(true);
      return;
    }
    const samePage = prev && prev.pageWidth === section.pageWidth && prev.pageHeight === section.pageHeight;
    if (section.type === 'continuous' && samePage) {
      // The new section continues on this page below the previous section's content.
      const top = this.atTop && this.col === 0 ? this.y : Math.max(this.y, ...this.page.colBottoms);
      this.closeColSet(top);
      this.page.colTop = top;
      // Continuous sections keep the page but bring their own margins and columns.
      this.page.cols = this.columnsFor(section, section.marginLeft + section.gutter, this.page.width - section.marginRight);
      this.page.colBottoms = this.page.cols.map(() => top);
      this.page.colSets.push({ section, cols: this.page.cols, top, bottom: top });
      this.col = 0;
      this.y = top;
      if (this.restartNumber !== null) {
        this.page.number = this.restartNumber;
        this.restartNumber = null;
      }
      this.sectionStartPage = this.pages.length - 1;
      return;
    }
    if (section.type === 'nextColumn' && samePage && this.col + 1 < this.page.cols.length) {
      this.breakColumn('hard');
      return;
    }
    // A page break that ends the previous section's last paragraph keeps the paragraph mark (and
    // its section break) on the page before; the section break then starts the next page itself,
    // so the empty page the break opened is not kept.
    // Bookmarks of skipped empty paragraphs on that page move to the section's first page.
    let anchors: DrawOp[] = [];
    if (this.breakOpenedPage === this.pages.length && this.atTop && this.page.body.every((op) => op.kind === 'anchor') && this.page.floats.length === 0 && this.pages.length > 1) {
      anchors = this.page.body;
      this.pages.pop();
      this.page = this.pages[this.pages.length - 1];
      this.y = Math.max(...this.page.colBottoms);
    }
    // Word keeps odd and even page numbers alternating. An odd- or even-page break that lands on
    // the wrong parity, or (with different odd and even headers) a section restarting its numbers
    // on the parity of the page before, gets an empty page first, printed without header or footer.
    const last = this.page.number;
    const parity = section.type === 'oddPage' ? 1 : section.type === 'evenPage' ? 0 : -1;
    if (this.restartNumber === null) {
      if (parity >= 0 && (last + 1) % 2 !== parity) this.blankPage(last + 1);
    } else if ((parity >= 0 || this.doc.settings.evenAndOddHeaders) && this.restartNumber % 2 === last % 2) this.blankPage(last + 1);
    this.newPage(true);
    this.page.body.push(...anchors);
  }

  // -------------------------------------------------------------------------------------------
  // Pages and columns

  private columnsFor(section: Section, left: number, right: number): Column[] {
    const width = right - left;
    const c = section.columns;
    if (c.count <= 1) return [{ x: left, w: width }];
    const cols: Column[] = [];
    if (!c.equalWidth && c.cols.length) {
      let x = left;
      for (const col of c.cols) {
        cols.push({ x, w: col.width });
        x += col.width + col.space;
      }
      return cols;
    }
    const w = (width - c.space * (c.count - 1)) / c.count;
    for (let i = 0; i < c.count; i++) cols.push({ x: left + i * (w + c.space), w });
    return cols;
  }

  private headerFooterFor(section: Section, number: number, firstOfSection: boolean): { header?: Block[]; footer?: Block[] } {
    if (section.titlePage && firstOfSection) return { header: section.headers.first, footer: section.footers.first };
    if (this.doc.settings.evenAndOddHeaders && number % 2 === 0) return { header: section.headers.even, footer: section.footers.even };
    return { header: section.headers.default, footer: section.footers.default };
  }

  private blankPage(number: number): void {
    const restart = this.restartNumber;
    this.restartNumber = number;
    this.newPage(false);
    this.page.blank = true;
    this.restartNumber = restart;
  }

  private newPage(sectionStart: boolean): void {
    this.flushContinuations();
    this.openPage(sectionStart);
  }

  private openPage(sectionStart: boolean): void {
    if (this.page) this.closeColSet(Math.max(this.y, ...this.page.colBottoms));
    const s = this.section;
    const prevNumber = this.pages.length ? this.pages[this.pages.length - 1].number : 0;
    let number = prevNumber + 1;
    if (this.restartNumber !== null) {
      number = this.restartNumber;
      this.restartNumber = null;
    }
    const firstOfSection = sectionStart;
    if (sectionStart) this.sectionStartPage = this.pages.length;
    let marginLeft = s.marginLeft + s.gutter;
    let marginRight = s.marginRight;
    // Mirrored margins swap inside and outside on even (left-hand) pages.
    if (this.doc.settings.mirrorMargins && this.pages.length % 2 === 1) {
      marginLeft = s.marginRight;
      marginRight = s.marginLeft + s.gutter;
    }
    const width = s.pageWidth;
    const height = s.pageHeight;
    const hf = this.headerFooterFor(s, number, firstOfSection);
    const hfWidth = width - marginLeft - marginRight;
    this.fieldPage = { number, format: s.pageNumberFormat, total: null, sectionTotal: null };
    const frame: PageFrame = { section: s, width, height, marginLeft, marginRight };
    const headerBox = hf.header ? this.hfBox(hf.header, frame, false).box : null;
    const footerBox = hf.footer ? this.hfBox(hf.footer, frame, true).box : null;
    this.fieldPage = null;
    let top = s.marginTop;
    if (!s.marginTopExact && headerBox) top = Math.max(top, s.header + headerBox.height);
    let bottom = height - s.marginBottom;
    if (!s.marginBottomExact && footerBox) bottom = Math.min(bottom, height - s.footer - footerBox.height);
    const cols = this.columnsFor(s, marginLeft, width - marginRight);
    const page: PageState = {
      section: s,
      sectionIndex: this.sectionIndex,
      firstOfSection,
      number,
      width,
      height,
      marginLeft,
      marginRight,
      top,
      bottom,
      cols,
      colTop: top,
      body: [],
      floats: [],
      exclusions: [],
      notes: [],
      overflow: [],
      continuation: false,
      noteHeight: 0,
      colBottoms: cols.map(() => top),
      colSets: [{ section: s, cols, top, bottom: top }],
    };
    this.pages.push(page);
    this.page = page;
    // Objects in the header/footer that wrap text push body text aside.
    for (const [box, y0] of [
      [headerBox, s.header],
      [footerBox, footerBox ? height - s.footer - footerBox.height : 0],
    ] as Array<[BoxResult | null, number]>) {
      if (!box) continue;
      for (const it of box.items)
        for (const op of it.ops) {
          if (op.kind !== 'float' || !op.drawing.anchor) continue;
          const cell = op.cell && { ...op.cell, x: op.cell.x + marginLeft, y: op.cell.y + y0 };
          const pos = this.floatPosition(op.drawing, page, { x: marginLeft, w: hfWidth }, marginLeft + op.x, y0 + op.y, cell);
          const ex = this.exclusionFor(op.drawing, pos.x, pos.y);
          if (ex && ex.y2 > top && ex.y1 < bottom) page.exclusions.push(ex);
        }
    }
    for (const [d, index] of this.early) {
      if (index !== this.pages.length - 1) continue;
      const pos = this.floatPosition(d, page, cols[0], cols[0].x, top);
      page.floats.push({ z: d.anchor!.zIndex, behind: d.anchor!.behindDoc, ops: renderDrawing(d, pos.x, pos.y, this.env) });
      const ex = this.exclusionFor(d, pos.x, pos.y);
      if (ex) page.exclusions.push(ex);
    }
    this.col = 0;
    this.y = top;
    this.atTop = true;
    this.topKind = sectionStart ? 'section' : 'hard';
  }

  private closeColSet(bottom: number): void {
    const set = this.page?.colSets[this.page.colSets.length - 1];
    if (set) set.bottom = Math.max(set.bottom, bottom);
  }

  private get column(): Column {
    return this.page.cols[this.col];
  }

  private get colBottom(): number {
    return this.page.bottom - this.page.noteHeight;
  }

  /** Move to the next column, or a new page after the last one. */
  private breakColumn(kind: 'hard' | 'soft'): void {
    this.page.colBottoms[this.col] = Math.max(this.page.colBottoms[this.col], this.y);
    if (this.col + 1 < this.page.cols.length) {
      this.col++;
      this.y = this.page.colTop;
    } else this.newPage(false);
    this.atTop = true;
    this.topKind = kind;
  }

  private breakPage(kind: 'hard' | 'soft'): void {
    this.page.colBottoms[this.col] = Math.max(this.page.colBottoms[this.col], this.y);
    this.newPage(false);
    this.topKind = kind;
  }

  private markPlaced(bottom: number): void {
    this.atTop = false;
    this.page.colBottoms[this.col] = Math.max(this.page.colBottoms[this.col], bottom);
  }

  // -------------------------------------------------------------------------------------------
  // Checkpoints (for keep-with-next, widow/orphan control and moving floats with their anchors)

  private checkpoint(): Checkpoint {
    const p = this.page;
    return {
      pages: this.pages.length,
      body: p.body.length,
      floats: p.floats.length,
      exclusions: p.exclusions.length,
      notes: p.notes.length,
      overflow: p.overflow.length,
      noteHeight: p.noteHeight,
      col: this.col,
      y: this.y,
      atTop: this.atTop,
      topKind: this.topKind,
      prevPara: this.prevPara,
      colBottoms: p.colBottoms.slice(),
      colTop: p.colTop,
      cols: p.cols,
      colSets: p.colSets.length,
      renderedBreaks: this.stats.renderedBreaks,
      matchedBreaks: this.stats.matchedBreaks,
      breakLogLength: this.stats.breakLog.length,
    };
  }

  private restore(cp: Checkpoint): void {
    this.pages.length = cp.pages;
    const p = this.pages[cp.pages - 1];
    this.page = p;
    p.body.length = cp.body;
    p.floats.length = cp.floats;
    p.exclusions.length = cp.exclusions;
    p.notes.length = cp.notes;
    p.overflow.length = cp.overflow;
    p.noteHeight = cp.noteHeight;
    p.colBottoms = cp.colBottoms.slice();
    p.colTop = cp.colTop;
    p.cols = cp.cols;
    p.colSets.length = cp.colSets;
    this.col = cp.col;
    this.y = cp.y;
    this.atTop = cp.atTop;
    this.topKind = cp.topKind;
    this.prevPara = cp.prevPara;
    this.stats.renderedBreaks = cp.renderedBreaks;
    this.stats.matchedBreaks = cp.matchedBreaks;
    this.stats.breakLog.length = cp.breakLogLength;
  }

  // -------------------------------------------------------------------------------------------
  // Blocks

  private placeBlocks(blocks: Block[]): void {
    const cps: Checkpoint[] = [];
    let i = 0;
    let guard = 0;
    while (i < blocks.length) {
      const block = blocks[i];
      cps[i] = this.checkpoint();
      this.blockIndex = this.sectionBlockBase + i;
      const next = blocks[i + 1] ?? null;
      if (block.type === 'table' && block.props.floating) {
        // Word anchors a floating table to the next paragraph and wraps text around it.
        let k = i + 1;
        while (blocks[k]?.type === 'table' && (blocks[k] as Table).props.floating) k++;
        const anchor = blocks[k];
        const float = anchor?.type === 'paragraph' ? this.tableFloat(block) : null;
        // A floating table too long for the rest of the page breaks across pages like an inline one.
        if (anchor?.type === 'paragraph' && float && this.floatFits(float)) {
          if (!this.tableFloats.has(anchor)) this.tableFloats.set(anchor, []);
          if (!this.floatedTables.has(block)) this.tableFloats.get(anchor)!.push(float);
          this.floatedTables.add(block);
          i++;
          continue;
        }
      }
      const moved = block.type === 'paragraph' ? this.placeParagraph(block, next, blocks[i - 1] ?? null) : this.placeTable(block);
      // Keep with next: when this block's start moved to a new column, take the chain of
      // keep-with-next paragraphs before it along, unless the chain already starts a column.
      if (moved && i > 0 && keepsWithNext(blocks[i - 1]) && guard < 1000) {
        let s = i - 1;
        while (s > 0 && keepsWithNext(blocks[s - 1]) && cps[s - 1].pages === cps[i].pages && cps[s - 1].col === cps[i].col) s--;
        const startCp = cps[s];
        if (!startCp.atTop && startCp.pages === cps[i].pages && startCp.col === cps[i].col) {
          guard++;
          this.restore(startCp);
          this.breakColumn('soft');
          i = s;
          // The chain now starts a column, so it will not be moved again.
          continue;
        }
      }
      i++;
    }
  }

  // -------------------------------------------------------------------------------------------
  // Paragraphs

  /** Returns true when the paragraph's first line was pushed into a new column. */
  private placeParagraph(para: Paragraph, next: Block | null, prevBlock: Block | null): boolean {
    let moved = false;
    // An empty paragraph that only carries a new-page section break, landing on a page that a page
    // break just opened, takes no room there: Word starts the next section on that page instead
    // of leaving it blank (see startSection).
    if (para.sectPr && para.sectPr.type !== 'continuous' && para.content.length === 0 && this.atTop && this.breakOpenedPage === this.pages.length) return false;
    // An empty paragraph that ends a section after other content, before a continuous section
    // break, takes no room either: Word shows only the break. A section of nothing but that
    // paragraph keeps its line.
    if (para.sectPr && prevBlock && this.nextSection?.type === 'continuous' && para.content.every((c) => c.type === 'bookmark' || c.type === 'renderedPageBreak')) {
      const x = this.page.cols[this.col].x;
      for (const c of para.content) if (c.type === 'bookmark') this.page.body.push({ kind: 'anchor', name: c.name, x, y: this.y });
      return false;
    }
    // Likewise before a section that starts a new page: nothing follows on this page, and Word lets
    // the empty paragraph stay here even where it would not fit.
    if (para.sectPr && prevBlock && (this.nextSection?.type === 'nextPage' || this.nextSection?.type === 'oddPage' || this.nextSection?.type === 'evenPage') && para.content.every((c) => c.type === 'bookmark' || c.type === 'renderedPageBreak')) {
      const x = this.page.cols[this.col].x;
      for (const c of para.content) if (c.type === 'bookmark') this.page.body.push({ kind: 'anchor', name: c.name, x, y: this.y });
      // Its space after still sets how much space before the next section's first paragraph keeps.
      this.prevPara = para;
      return false;
    }
    if (para.props.pageBreakBefore && !this.atTop) this.breakPage('hard');
    const layout = new ParagraphLayout(para, this.env);
    const nextPara = next?.type === 'paragraph' ? next : null;
    const prevParaBlock = prevBlock?.type === 'paragraph' ? prevBlock : null;
    const frame = paraFrame(prevParaBlock, para, nextPara);
    const deco = hasDecorations(para.props);
    let attempts = 0;
    restart: for (;;) {
      attempts++;
      const paraCp = this.checkpoint();
      // Spacing before this paragraph.
      const gap = gapBetween(this.prevPara, para);
      if (this.atTop) {
        // Word drops a paragraph's space before at the top of a page, unless the paragraph itself
        // asked for the new page (page break before).
        // Automatic (HTML) space before collapses into the top of the document, as in a header.
        // After a section break, the previous paragraph's space after stays on its page and only
        // what the larger gap adds to it comes before this paragraph.
        if (this.topKind !== 'soft' && !(this.blockIndex === 0 && para.props.spaceBeforeAuto)) {
          if (para.props.pageBreakBefore) this.y += gapBetween(null, para).before;
          else if (this.topKind === 'section') this.y += Math.max(0, gap.after + gap.before - gapBetween(this.prevPara, null).after);
          else this.y += gap.before;
        }
      } else this.y += (this.prevPara ? gap.after : 0) + gap.before;
      this.y += frame.topExtra;
      let floatsPlaced = false;
      let linesHere = 0;
      let startedHere = true;
      let lineCps: Array<{ cp: Checkpoint; snap: ReturnType<ParagraphLayout['snapshot']> }> = [];
      const markHeight = layout.markRun.ascent + layout.markRun.descent;
      let estimate = { h: markHeight, text: markHeight };
      // The gap beside an object where the line just placed continues, and that line's bottom.
      let beside: { left: number; right: number; bottom: number } | null = null;
      let besideBottom = 0;
      for (;;) {
        if (!floatsPlaced) {
          const tables = this.tableFloats.get(para);
          const drawings = tables ? [...tables, ...layout.floatDrawings] : layout.floatDrawings;
          // Before Word 2013, a paragraph whose first line starts right of an object it wraps around
          // measures its own objects' column offsets from where that line starts.
          let lineX: number | undefined;
          if (drawings.length && this.env.compatibilityMode < 15) {
            const base = layout.lineBounds(this.column.w);
            const slot = findSlot(this.page.exclusions, this.y, estimate.h, this.column.x + base.left, this.column.x + base.right, estimate.text);
            if (slot.y <= this.y + EPS && slot.left > this.column.x + base.left + EPS) lineX = slot.left;
          }
          this.placeFloats(drawings, this.column, this.y, lineX);
          floatsPlaced = true;
        }
        const snap = layout.snapshot();
        const lineCp = this.checkpoint();
        const placed = this.layoutLine(layout, estimate, beside ?? undefined);
        beside = null;
        if (!placed) break;
        const { line, y, left } = placed;
        estimate = { h: line.height, text: textHeight(line) };
        const noteH = this.notesHeight(line.notes);
        // The extra space of "multiple" line spacing sits below the text, and Word lets it hang into
        // the bottom margin: only the text of the line has to fit. Above a footer tall enough to push
        // the body up, the text has to clear the footer and the extra space may still reach down to
        // the bottom margin, not into it.
        const marginEdge = this.page!.height - this.page!.section.marginBottom;
        const intoMargin = this.page!.bottom >= marginEdge - EPS;
        const text = textHeight(line);
        const fits = y + text + noteH <= this.colBottom + EPS && (intoMargin || y + line.height + noteH <= marginEdge + EPS);
        if (!fits && !this.atTop) {
          besideBottom = 0;
          layout.restore(snap);
          this.restore(lineCp);
          const canMoveWhole = startedHere && !paraCp.atTop && attempts < 3;
          if (linesHere === 0) {
            // Nothing of this paragraph fits here.
            this.restore(paraCp);
            this.breakColumn('soft');
            layout.reset();
            if (startedHere) moved = true;
            continue restart;
          }
          if (canMoveWhole && para.props.keepLines) {
            this.restore(paraCp);
            this.breakColumn('soft');
            layout.reset();
            moved = true;
            continue restart;
          }
          if (para.props.widowControl) {
            if (linesHere === 1 && canMoveWhole) {
              this.restore(paraCp);
              this.breakColumn('soft');
              layout.reset();
              moved = true;
              continue restart;
            }
            if (line.last && linesHere >= 2) {
              // Widow: take the previous line along to the next page.
              if (linesHere - 1 < 2 && canMoveWhole) {
                this.restore(paraCp);
                this.breakColumn('soft');
                layout.reset();
                moved = true;
                continue restart;
              }
              const prev = lineCps[lineCps.length - 1];
              this.restore(prev.cp);
              layout.restore(prev.snap);
            }
          }
          this.breakColumn('soft');
          startedHere = false;
          linesHere = 0;
          lineCps = [];
          continue;
        }
        // Place the line.
        lineCps.push({ cp: lineCp, snap });
        if (line.renderedPageBreak) {
          this.stats.renderedBreaks++;
          if (this.atTop) this.stats.matchedBreaks++;
          this.stats.breakLog.push({ page: this.pages.length - 1, top: this.atTop, y: Math.round(y), block: this.blockIndex });
        }
        const col = this.column;
        if (deco) {
          const nextGap = frame.groupedNext && nextPara ? gapBetween(para, nextPara) : { after: 0, before: 0 };
          const top = line.first ? y - frame.topExtra : y;
          const bottom = line.last ? y + line.height + frame.bottomExtra + nextGap.after + nextGap.before + shadedSpaceAfter(para, frame, nextPara) : y + line.height;
          this.page.body.push(...paraDecorations(para.props, frame, col.x, col.w, top, bottom, line.first, line.last, line.last ? sideBorderRunOn(para, frame, nextPara) : 0));
        }
        this.page.body.push(...layout.renderLine(line, col.x + left, y));
        const bottom = Math.max(y + line.height, besideBottom);
        // The paragraph's next line may continue beside an object wrapped on both sides.
        const base = layout.lineBounds(col.w);
        const rest = !line.last && line.brk !== 'page' && line.brk !== 'column' ? gapBeside(this.page.exclusions, y, line.height, textHeight(line), col.x + base.left, col.x + base.right, placed.right) : null;
        beside = rest && { ...rest, bottom };
        besideBottom = rest ? bottom : 0;
        this.y = rest ? y : bottom;
        this.markPlaced(bottom);
        this.addNotes(line.notes);
        linesHere++;
        if (line.brk === 'page') {
          this.breakPage('hard');
          // Before Word 2013 (compatibility mode 15), a paragraph that opens with a page break
          // starts on the new page as far as its spacing goes: its space before applies there.
          if (this.env.compatibilityMode < 15 && line.first && !line.last && layout.onlyBreak(line)) this.y += gapBetween(null, para).before;
          if (line.last) this.breakOpenedPage = this.pages.length;
          startedHere = false;
          linesHere = 0;
          lineCps = [];
        } else if (line.brk === 'column') {
          this.breakColumn('hard');
          startedHere = false;
          linesHere = 0;
          lineCps = [];
        }
      }
      break;
    }
    this.y += frame.bottomExtra;
    this.prevPara = para;
    return moved;
  }

  /** Lay out the next line at the current position, flowing around wrapped objects. */
  private layoutLine(layout: ParagraphLayout, estimate: { h: number; text: number }, gap?: { left: number; right: number; bottom: number }): { line: Line; y: number; left: number; right: number } | null {
    const placed = wrapLine(layout, this.page.exclusions, this.y, estimate, this.column, gap);
    return placed && { ...placed, left: 0 };
  }

  /**
   * Lay out a header or footer, flowing its text around its own wrapped floating objects (a logo
   * set to square wrapping pushes the header's text below it). `y0` is the box's top on the page.
   */
  private hfBox(blocks: Block[], page: PageFrame, footer: boolean): { box: BoxResult; y0: number } {
    const s = page.section;
    const width = page.width - page.marginLeft - page.marginRight;
    const top = (box: BoxResult) => (footer ? page.height - s.footer - box.height : s.header);
    const plain = layoutBlocks(blocks, width, this.env);
    if (!plain.items.some((it) => it.ops.some((op) => op.kind === 'float' && op.drawing.anchor && op.drawing.anchor.wrap !== 'none'))) return { box: plain, y0: top(plain) };
    // A footer's top depends on its height: place its objects from the unwrapped height.
    const y0 = top(plain);
    const col = { x: page.marginLeft, w: width };
    const wrap: BoxWrap = (d, y) => {
      const pos = this.floatPosition(d, page, col, page.marginLeft, y0 + y);
      const ex = this.exclusionFor(d, pos.x, pos.y);
      return ex && { ...ex, x1: ex.x1 - page.marginLeft, x2: ex.x2 - page.marginLeft, y1: ex.y1 - y0, y2: ex.y2 - y0 };
    };
    const box = layoutBlocks(blocks, width, this.env, { wrap });
    return { box, y0: top(box) };
  }

  private floatPosition(d: Drawing, page: PageFrame, col: Column, paraX: number, paraY: number, cell?: FloatOp['cell']): { x: number; y: number } {
    const a = d.anchor!;
    const s = page.section;
    const w = d.width;
    const h = d.height;
    // An object laid out in its table cell measures page- and margin-relative positions from the
    // cell (a logo in a header table stays in its cell instead of dropping into the body).
    const inCell = a.layoutInCell ? cell : undefined;
    const h0 = (() => {
      if (inCell && (a.posH.relativeFrom === 'page' || a.posH.relativeFrom === 'margin')) return { base: inCell.x, size: inCell.w };
      switch (a.posH.relativeFrom) {
        case 'page':
          return { base: 0, size: page.width };
        case 'column':
          return { base: col.x, size: col.w };
        case 'character':
          return { base: paraX, size: 0 };
        case 'leftMargin':
        case 'insideMargin':
          return { base: 0, size: page.marginLeft };
        case 'rightMargin':
        case 'outsideMargin':
          return { base: page.width - page.marginRight, size: page.marginRight };
        default:
          return { base: page.marginLeft, size: page.width - page.marginLeft - page.marginRight };
      }
    })();
    let x = h0.base;
    const ha = a.posH.align;
    // Word aligns an object together with its effects (a shadow keeps inside the margin).
    if (ha === 'right' || ha === 'outside') x = h0.base + h0.size - w - d.effectR;
    else if (ha === 'center') x = h0.base + (h0.size - w - d.effectL - d.effectR) / 2 + d.effectL;
    else if (ha === 'left' || ha === 'inside') x = h0.base + d.effectL;
    else if (a.posH.pct !== undefined) x = h0.base + a.posH.pct * h0.size;
    else x = h0.base + (a.posH.offset ?? 0);
    const v0 = (() => {
      if (inCell && (a.posV.relativeFrom === 'page' || a.posV.relativeFrom === 'margin')) return { base: inCell.y, size: inCell.h };
      switch (a.posV.relativeFrom) {
        case 'page':
          return { base: 0, size: page.height };
        case 'paragraph':
        case 'line':
          return { base: paraY, size: 0 };
        case 'topMargin':
          return { base: 0, size: s.marginTop };
        case 'bottomMargin':
          return { base: page.height - s.marginBottom, size: s.marginBottom };
        default:
          return { base: s.marginTop, size: page.height - s.marginTop - s.marginBottom };
      }
    })();
    let y = v0.base;
    const va = a.posV.align;
    if (va === 'bottom' || va === 'outside') y = v0.base + v0.size - h - d.effectB;
    else if (va === 'center') y = v0.base + (v0.size - h - d.effectT - d.effectB) / 2 + d.effectT;
    else if (va === 'top' || va === 'inside') y = v0.base + d.effectT;
    else if (a.posV.pct !== undefined) y = v0.base + a.posV.pct * v0.size;
    else y = v0.base + (a.posV.offset ?? 0);
    // Word keeps a text-wrapping object laid out in a cell from rising above the cell's top.
    if (inCell && a.wrap !== 'none') y = Math.max(y, inCell.y);
    return { x, y };
  }

  private exclusionFor(d: Drawing, x: number, y: number): Exclusion | null {
    const a = d.anchor;
    if (!a || a.wrap === 'none') return null;
    if (a.polygon && !d.rotation && !d.flipH && !d.flipV) {
      // Tight and through wrapping follow the object's wrap outline beside each line.
      const x1 = x - d.effectL - a.distL;
      const y1 = y - d.effectT - a.distT;
      const pts = a.polygon.map(([px, py]): [number, number] => [x + px * d.width - x1, y + py * d.height - y1]);
      const poly = { pts, padL: a.distL, padT: a.distT, padR: a.distR, padB: a.distB };
      return { x1, y1, x2: x + d.width + d.effectR + a.distR, y2: y + d.height + d.effectB + a.distB, wrap: a.wrap, side: a.wrapText, poly };
    }
    // Word lays out an object turned nearer a quarter turn than upright in its upright box turned
    // about the centre (width and height swapped); the effect extents grow that box.
    const turn = (((d.rotation % 180) + 180) % 180);
    const swap = turn >= 45 && turn < 135;
    const w = swap ? d.height : d.width;
    const h = swap ? d.width : d.height;
    x += (d.width - w) / 2;
    y += (d.height - h) / 2;
    return {
      x1: x - d.effectL - a.distL,
      y1: y - d.effectT - a.distT,
      x2: x + w + d.effectR + a.distR,
      y2: y + h + d.effectB + a.distB,
      wrap: a.wrap,
      side: a.wrapText,
    };
  }

  /**
   * Word keeps an object that text wraps around, placed below its paragraph, on its anchor's page:
   * one that would run past the bottom margin (before Word 2013, the page's edge) moves up to end
   * there. A top-and-bottom object moved up over its own paragraph's line pushes that line, and so
   * the paragraph, to the next page.
   */
  private keptOnPage(d: Drawing, y: number): number {
    const a = d.anchor!;
    if (d.hidden || a.wrap === 'none' || (a.posV.relativeFrom !== 'paragraph' && a.posV.relativeFrom !== 'line') || a.posV.align || a.posV.pct !== undefined) return y;
    const limit = this.env.compatibilityMode < 15 ? this.page.height : this.page.height - this.page.section.marginBottom;
    return Math.min(y, Math.max(limit - d.height - d.effectB, this.page.colTop));
  }

  /** Whether a float placed from the current position ends above the bottom of the page body. */
  private floatFits(d: Drawing): boolean {
    const pos = this.floatPosition(d, this.page, this.column, this.column.x, this.y);
    return pos.y + d.height <= this.colBottom + EPS;
  }

  /** A floating table as a borderless text box positioned by its w:tblpPr, wrapping text on both sides. */
  private tableFloat(table: Table): Drawing {
    const f = table.props.floating!;
    // Inside the box the table starts at the box's left edge: its own indent and alignment give way to
    // the float position (before Word 2013 the indent measured the first cell's text, not its border).
    const first = table.rows[0]?.cells[0];
    const indent = this.env.compatibilityMode < 15 ? (first?.margins.left ?? table.props.cellMargins.left) : 0;
    const inner: Table = { ...table, props: { ...table.props, floating: null, indent, align: 'left' } };
    const width = measureTable(inner, this.column.w, this.env).width;
    const height = layoutBlocks([inner], width, this.env).height;
    const horz = f.horzAnchor === 'page' ? 'page' : f.horzAnchor === 'margin' ? 'margin' : 'column';
    const vert = f.vertAnchor === 'page' ? 'page' : f.vertAnchor === 'margin' ? 'margin' : 'paragraph';
    const xAlign = f.xAlign && f.xAlign !== 'inline' ? f.xAlign : undefined;
    const yAlign = f.yAlign && f.yAlign !== 'inline' ? f.yAlign : undefined;
    const drawing: Drawing = {
      kind: 'shape',
      inline: false,
      width,
      height,
      effectL: 0,
      effectT: 0,
      effectR: 0,
      effectB: 0,
      rotation: 0,
      flipH: false,
      flipV: false,
      geometry: 'rect',
      fill: null,
      line: null,
      textBox: { blocks: [inner], insetL: 0, insetT: 0, insetR: 0, insetB: 0, vAlign: 'top', autoFit: false, unclipped: true },
      hidden: false,
      anchor: {
        behindDoc: false,
        wrap: 'square',
        wrapText: 'bothSides',
        distT: f.topFromText,
        distB: f.bottomFromText,
        distL: f.leftFromText,
        distR: f.rightFromText,
        posH: { relativeFrom: horz, align: xAlign, offset: f.x },
        posV: { relativeFrom: vert, align: yAlign, offset: f.y },
        layoutInCell: true,
        allowOverlap: false,
        zIndex: 0,
      },
    };
    this.tableFloatDrawings.add(drawing);
    return drawing;
  }

  private placeFloats(drawings: Drawing[], col: Column, paraY: number, lineX?: number): void {
    const index = this.pages.length - 1;
    for (const d of drawings) {
      if (!d.anchor || d.hidden) continue;
      const earlyPage = this.early.get(d);
      if (earlyPage === index) continue;
      if (earlyPage !== undefined) {
        // The anchor moved off the page the float was placed on early: place it with its anchor.
        this.early.delete(d);
        this.earlyChanged = true;
      }
      const h = d.anchor.posH;
      const fromLine = lineX !== undefined && h.relativeFrom === 'column' && !h.align && h.pct === undefined;
      const pos = this.floatPosition(d, this.page, fromLine ? { ...col, x: lineX } : col, col.x, paraY);
      pos.y = this.keptOnPage(d, pos.y);
      this.page.floats.push({ z: d.anchor.zIndex, behind: d.anchor.behindDoc, ops: renderDrawing(d, pos.x, pos.y, this.env) });
      const ex = this.exclusionFor(d, pos.x, pos.y);
      if (!ex) continue;
      this.page.exclusions.push(ex);
      // Word lays out a page with all its floats known, so text above the anchor wraps around an
      // object positioned against the page or margin too. Lay out again with it placed as the page opens.
      const fixed = !this.tableFloatDrawings.has(d) && PAGE_FIXED_H.has(d.anchor.posH.relativeFrom) && PAGE_FIXED_V.has(d.anchor.posV.relativeFrom);
      if (fixed && earlyPage === undefined && ex.y1 < paraY - 0.5 && ex.y2 > this.page.colTop && (paraY > this.page.colTop + 0.5 || this.col > 0)) {
        this.early.set(d, index);
        this.earlyChanged = true;
      }
    }
  }

  /** Floats inside table cells and header boxes, already translated to page coordinates. */
  private extractFloats(ops: DrawOp[], page: PageState, target: DrawOp[], floats: FloatEntry[]): void {
    for (const op of ops) {
      if (op.kind !== 'float') {
        target.push(op);
        continue;
      }
      const d = op.drawing;
      if (!d.anchor || d.hidden) continue;
      const pos = this.floatPosition(d, page, { x: op.x, w: op.areaWidth }, op.x, op.y, op.cell);
      floats.push({ z: d.anchor.zIndex, behind: d.anchor.behindDoc, ops: renderDrawing(d, pos.x, pos.y, this.env) });
    }
  }

  // -------------------------------------------------------------------------------------------
  // Footnotes

  private notesHeight(notes: Line['notes']): number {
    let h = 0;
    const width = this.footnoteWidth();
    for (const n of notes) {
      if (n.noteType !== 'footnote') continue;
      if (this.page.notes.some((p) => p.id === n.id)) continue;
      if (this.page.noteHeight === 0 && h === 0) h += this.footnoteSeparator(width).height;
      const box = this.noteBox('footnote', n.id, width);
      const first = box.items[0];
      if (first) h += first.y + first.height;
    }
    return h;
  }

  private footnoteWidth(): number {
    return this.page.width - this.page.marginLeft - this.page.marginRight;
  }

  private addNotes(notes: NonNullable<BoxItem['notes']>): void {
    const width = this.footnoteWidth();
    for (const n of notes) {
      if (n.noteType === 'endnote') {
        if (!this.endnotesUsed.includes(n.id)) this.endnotesUsed.push(n.id);
        continue;
      }
      if (this.page.notes.some((p) => p.id === n.id) || this.page.overflow.some((p) => p.id === n.id)) continue;
      const box = this.noteBox('footnote', n.id, width);
      if (box.items.length) this.reserveNote({ id: n.id, box, start: 0, end: box.items.length, offset: 0 });
    }
  }

  private noteSeparator(): BoxResult {
    if (this.page.continuation && this.doc.footnoteContinuationSeparator) {
      return layoutBlocks(this.doc.footnoteContinuationSeparator, this.footnoteWidth(), this.env);
    }
    return this.footnoteSeparator(this.footnoteWidth());
  }

  private reserveNote(piece: FootnotePiece): void {
    const page = this.page;
    if (page.overflow.length) {
      page.overflow.push(piece);
      return;
    }
    const separator = page.notes.length ? 0 : this.noteSeparator().height;
    const available = page.bottom - Math.max(...page.colBottoms) - page.noteHeight - separator;
    const { box, start, offset } = piece;
    let end = start;
    let height = 0;
    while (end < piece.end) {
      const item = box.items[end];
      const nextHeight = item.y + item.height - offset;
      if (nextHeight > available + EPS) break;
      height = nextHeight;
      end++;
    }
    if (end === start && page.continuation && !page.notes.length) {
      throw new Error('A footnote line or object is taller than the printable page');
    }
    if (end > start) {
      page.notes.push({ ...piece, end });
      page.noteHeight += separator + height;
    }
    if (end < piece.end) page.overflow.push({ ...piece, start: end, offset: box.items[end].y });
  }

  /** Keep overflow with its reference section even when the next body page restarts numbering. */
  private flushContinuations(): void {
    if (!this.page?.overflow.length) return;
    const { section, sectionIndex, restartNumber } = this;
    this.section = this.page.section;
    this.sectionIndex = this.page.sectionIndex;
    this.restartNumber = null;
    while (this.page.overflow.length) {
      const remaining = this.page.overflow;
      this.openPage(false);
      this.page.continuation = true;
      for (const piece of remaining) this.reserveNote(piece);
    }
    this.section = section;
    this.sectionIndex = sectionIndex;
    this.restartNumber = restartNumber;
  }

  private placeEndnotes(): void {
    if (!this.endnotesUsed.length) return;
    const col = this.column;
    const items: BoxItem[] = [];
    const sep = this.doc.endnoteSeparator ? layoutBlocks(this.doc.endnoteSeparator, col.w, this.env) : this.footnoteSeparator(col.w);
    items.push(...sep.items);
    let base = sep.height;
    for (const id of this.endnotesUsed) {
      const box = this.noteBox('endnote', id, col.w);
      for (const it of box.items) items.push({ ...it, y: it.y + base, ops: translateOps(it.ops, 0, base) });
      base += box.height;
    }
    if (!this.atTop && this.prevPara) this.y += spaceAfter(this.prevPara.props);
    let origin = this.y - (items[0]?.y ?? 0);
    for (const it of items) {
      if (this.y + it.height > this.colBottom + EPS && !this.atTop) {
        this.breakColumn('soft');
        origin = this.y - it.y;
      }
      const c = this.column;
      this.page.body.push(...translateOps(it.ops, c.x, origin));
      this.y = origin + it.y + it.height;
      this.markPlaced(this.y);
    }
  }

  // -------------------------------------------------------------------------------------------
  // Tables

  private placeTable(table: Table): boolean {
    let col = this.column;
    let tl = measureTable(table, col.w, this.env);
    // Space after the previous paragraph still applies above the table.
    if (!this.atTop && this.prevPara) this.y += spaceAboveTable(this.prevPara, table);
    this.prevPara = null;
    const merge: MergeProgress = new Map();
    let pieces: RowPiece[] = [];
    let moved = false;
    let rowsOnPage = 0;
    const flush = () => {
      if (!pieces.length) return;
      const seg = renderSegment(tl, pieces, col.x, merge);
      this.extractFloats(seg.ops, this.page, this.page.body, this.page.floats);
      this.addNotes(seg.notes);
      pieces = [];
    };
    const headerHasFields = tl.rows.slice(0, tl.headerRows).some((row) =>
      row.cells.some((cell) => cell.box.items.some((item) => item.ops.some((op) => op.kind === 'glyphs' && op.field))),
    );
    const newColumn = (rowIndex: number) => {
      flush();
      this.breakColumn('soft');
      const c = this.column;
      if (Math.abs(c.w - col.w) > EPS) {
        col = c;
        tl = measureTable(table, col.w, this.env);
      }
      col = c;
      rowsOnPage = 0;
      if (rowIndex < tl.headerRows) return;
      if (headerHasFields) {
        // A repeated PAGE field has a different value on each page. Shape this copy
        // in its own page context so it cannot overwrite the first header's field.
        this.fieldPage = { number: this.page.number, format: this.section.pageNumberFormat, total: null, sectionTotal: null };
        const header = measureTable({ ...table, rows: table.rows.slice(0, tl.headerRows) }, col.w, this.env);
        this.fieldPage = null;
        const repeats = header.rows.map((row, r) => {
          const y = this.y;
          this.y += row.height;
          return { r, y, height: row.height };
        });
        const seg = renderSegment(header, repeats, col.x, new Map());
        for (const op of seg.ops) if (op.kind === 'glyphs') delete op.field;
        this.extractFloats(seg.ops, this.page, this.page.body, this.page.floats);
        this.addNotes(seg.notes);
        return;
      }
      for (let h = 0; h < tl.headerRows; h++) {
        pieces.push({ r: h, y: this.y, height: tl.rows[h].height });
        this.y += tl.rows[h].height;
      }
    };
    const progress = new Map<CellLayout, number>();
    for (let r = 0; r < tl.rows.length; r++) {
      const row = tl.rows[r];
      let first = true;
      for (let guard = 0; ; guard++) {
        if (guard === 500) throw new Error('Table row exceeds the pagination limit');
        const remainingH = this.rowRemainingHeight(tl, r, progress);
        const fitsWhole = this.y + remainingH <= this.colBottom + EPS;
        if (fitsWhole || (this.atTop && rowsOnPage === 0 && !this.canSplit(tl, r))) {
          const slices = this.sliceRest(tl, r, progress);
          pieces.push({ r, y: this.y, height: remainingH, slices });
          this.y += remainingH;
          this.markPlaced(this.y);
          rowsOnPage++;
          break;
        }
        const emptyHere = this.atTop;
        // A row whose content fits here but whose minimum height does not moves on whole: Word
        // splits a row only to break its content.
        const started = row.cells.some((c) => (progress.get(c) ?? 0) > 0);
        const onlyHeightOverflows = this.y + this.rowRemainingHeight(tl, r, progress, true) + (started ? 0 : row.borderTop) <= this.colBottom + EPS;
        if (!emptyHere && (row.row.cantSplit || !this.canSplit(tl, r) || onlyHeightOverflows)) {
          if (first && r === 0) moved = true;
          newColumn(r);
          first = false;
          continue;
        }
        // Split the row at the column bottom.
        const split = this.splitRow(tl, r, progress, this.colBottom - this.y, emptyHere);
        if (!split) {
          if (first && r === 0) moved = true;
          newColumn(r);
          first = false;
          continue;
        }
        pieces.push({ r, y: this.y, height: split.height, slices: split.slices });
        this.y += split.height;
        this.markPlaced(this.y);
        rowsOnPage++;
        newColumn(r);
        first = false;
      }
    }
    flush();
    this.atTop = false;
    return moved;
  }

  private canSplit(tl: TableLayout, r: number): boolean {
    const row = tl.rows[r];
    // Turned text runs the full height of its row, so the row moves to the next page whole.
    if (row.row.cantSplit || row.row.heightRule === 'exact' || row.cells.some((c) => c.turned)) return false;
    return row.cells.some((c) => c.rowSpan === 1 && !c.cell.vMerge && c.box.items.length > 1);
  }

  /** Height of what is left of row r after earlier split pieces; with `content`, of its content alone. */
  private rowRemainingHeight(tl: TableLayout, r: number, progress: Map<CellLayout, number>, content = false): number {
    const row = tl.rows[r];
    if (!content && !row.cells.some((c) => (progress.get(c) ?? 0) > 0)) return row.height;
    let h = 0;
    for (const c of row.cells) {
      if (c.cell.vMerge) continue;
      const from = progress.get(c) ?? 0;
      const items = c.box.items;
      if (from >= items.length) continue;
      const shift = items[from].y;
      const last = items[items.length - 1];
      h = Math.max(h, last.y + last.height - shift + c.cell.margins.top + c.cell.margins.bottom);
    }
    return h;
  }

  private sliceRest(tl: TableLayout, r: number, progress: Map<CellLayout, number>): RowPiece['slices'] {
    const row = tl.rows[r];
    if (!row.cells.some((c) => (progress.get(c) ?? 0) > 0)) return undefined;
    const slices = new Map<CellLayout, { from: number; to: number; shift: number }>();
    for (const c of row.cells) {
      const from = progress.get(c) ?? 0;
      const items = c.box.items;
      slices.set(c, { from: Math.min(from, items.length), to: items.length, shift: from < items.length ? items[from].y : 0 });
      progress.set(c, items.length);
    }
    return slices;
  }

  private splitRow(tl: TableLayout, r: number, progress: Map<CellLayout, number>, avail: number, force: boolean): { height: number; slices: Map<CellLayout, { from: number; to: number; shift: number }> } | null {
    const row = tl.rows[r];
    const slices = new Map<CellLayout, { from: number; to: number; shift: number }>();
    let height = 0;
    let advanced = false;
    for (const c of row.cells) {
      const items = c.box.items;
      const from = progress.get(c) ?? 0;
      const shift = from < items.length ? items[from].y : 0;
      // A piece that starts the row has its content below the row's top border.
      const lead = from === 0 ? row.borderTop : 0;
      const inner = avail - lead - c.cell.margins.top - c.cell.margins.bottom;
      let to = from;
      if (!c.cell.vMerge) {
        while (to < items.length && items[to].y - shift + items[to].height <= inner + EPS) to++;
        const fits = to;
        while (to > from && to < items.length && items[to].keepWithPrev) to--;
        // Before Word 2013 (compatibility mode 14 and earlier) keeping lines together never empties a
        // cell's piece: Word leaves the first line of a long statement at the foot of the page. Word
        // 2013 moves the row to the next page instead.
        if (to === from) to = force ? Math.max(fits, Math.min(from + 1, items.length)) : this.env.compatibilityMode < 15 ? fits : from;
      }
      if (to > from) advanced = true;
      slices.set(c, { from, to, shift });
      if (to > from) height = Math.max(height, lead + items[to - 1].y + items[to - 1].height - shift + c.cell.margins.top + c.cell.margins.bottom);
    }
    if (!advanced) return null;
    for (const [c, s] of slices) progress.set(c, s.to);
    return { height: Math.min(Math.max(height, 0), Math.max(avail, height)), slices };
  }

  // -------------------------------------------------------------------------------------------
  // Finishing pages: headers, footers, footnotes, borders

  private finish(): Page[] {
    if (this.page) this.closeColSet(Math.max(this.y, ...this.page.colBottoms));
    const total = this.pages.length;
    const sectionTotals = new Map<Section, number>();
    for (const p of this.pages) {
      for (const section of new Set(p.colSets.map((set) => set.section))) {
        sectionTotals.set(section, (sectionTotals.get(section) ?? 0) + 1);
      }
    }
    for (const p of this.pages) {
      for (const op of p.body) if (op.kind === 'anchor' && !this.anchorPages.has(op.name)) this.anchorPages.set(op.name, { number: p.number, format: p.section.pageNumberFormat });
    }
    const out: Page[] = [];
    for (const p of this.pages) {
      const s = p.section;
      const hfOps: DrawOp[] = [];
      const floats = p.floats.slice();
      const hf = p.blank ? {} : this.headerFooterFor(s, p.number, p.firstOfSection);
      const hfWidth = p.width - p.marginLeft - p.marginRight;
      this.fieldPage = { number: p.number, format: s.pageNumberFormat, total, sectionTotal: sectionTotals.get(p.section) ?? null };
      for (const [blocks, footer] of [[hf.header, false], [hf.footer, true]] as const) {
        if (!blocks) continue;
        const { box, y0 } = this.hfBox(blocks, p, footer);
        for (const it of box.items) this.extractFloats(translateOps(it.ops, p.marginLeft, y0), p, hfOps, floats);
      }
      this.fieldPage = null;

      // Footnotes sit at the bottom of the body area.
      const body = p.body.slice();
      if (p.notes.length) {
        let y = p.bottom - p.noteHeight;
        const sep = p.continuation && this.doc.footnoteContinuationSeparator
          ? layoutBlocks(this.doc.footnoteContinuationSeparator, hfWidth, this.env)
          : this.footnoteSeparator(hfWidth);
        for (const it of sep.items) body.push(...translateOps(it.ops, p.marginLeft, y));
        y += sep.height;
        for (const n of p.notes) {
          for (let i = n.start; i < n.end; i++) {
            this.extractFloats(translateOps(n.box.items[i].ops, p.marginLeft, y - n.offset), p, body, floats);
          }
          const last = n.box.items[n.end - 1];
          y += last.y + last.height - n.offset;
        }
      }

      // Vertical alignment of the page's text (single-section pages only).
      if (s.vAlign !== 'top' && p.colSets.length === 1 && !p.notes.length) {
        const contentBottom = Math.max(...p.colBottoms);
        const free = p.bottom - contentBottom;
        if (free > 0) {
          const dy = s.vAlign === 'bottom' ? free : free / 2;
          const shifted = translateOps(body, 0, dy);
          body.length = 0;
          body.push(...shifted);
        }
      }

      // Column separator lines.
      for (const set of p.colSets) {
        if (!set.section.columns.separator || set.cols.length < 2) continue;
        const bottom = Math.max(set.bottom, set.top);
        for (let i = 1; i < set.cols.length; i++) {
          const x = (set.cols[i - 1].x + set.cols[i - 1].w + set.cols[i].x) / 2;
          body.push({ kind: 'line', x1: x, y1: set.top, x2: x, y2: bottom, width: 0.5, color: '000000' });
        }
      }

      const behind: DrawOp[] = [];
      const front: DrawOp[] = [];
      if (this.doc.background && this.doc.settings.displayBackgroundShape) behind.push({ kind: 'rect', x: 0, y: 0, w: p.width, h: p.height, fill: this.doc.background });
      floats.sort((a, b) => a.z - b.z);
      for (const f of floats) (f.behind ? behind : front).push(...f.ops);
      const borders = p.blank ? [] : this.pageBorders(p);
      if (borders.length) (s.pageBorders?.zOrder === 'back' ? behind : front).push(...borders);
      for (const layer of [body, behind, front]) {
        for (const op of layer) {
          if (op.kind !== 'glyphs' || !op.field) continue;
          const measured = this.measuredFields.get(op.field);
          if (!measured) continue;
          this.fieldPage = { number: p.number, format: measured.section.pageNumberFormat, total, sectionTotal: sectionTotals.get(measured.section) ?? 1 };
          const value = this.fieldText(op.field);
          this.fieldPage = null;
          if (value !== measured.value) {
            this.fieldValues.set(op.field, value);
            this.fieldsChanged = true;
          }
        }
      }
      out.push({ width: p.width, height: p.height, behind, headerFooter: hfOps, body, front, section: p.sectionIndex, number: p.number, numberFormat: s.pageNumberFormat });
    }
    return out;
  }

  private pageBorders(p: PageState): DrawOp[] {
    const pb = p.section.pageBorders;
    if (!pb) return [];
    if (pb.display === 'firstPage' && !p.firstOfSection) return [];
    if (pb.display === 'notFirstPage' && p.firstOfSection) return [];
    const ops: DrawOp[] = [];
    const s = p.section;
    let x1: number, y1: number, x2: number, y2: number;
    if (pb.offsetFrom === 'page') {
      x1 = (pb.left?.space ?? 24) + (pb.left?.width ?? 0) / 2;
      y1 = (pb.top?.space ?? 24) + (pb.top?.width ?? 0) / 2;
      x2 = p.width - (pb.right?.space ?? 24) - (pb.right?.width ?? 0) / 2;
      y2 = p.height - (pb.bottom?.space ?? 24) - (pb.bottom?.width ?? 0) / 2;
    } else {
      x1 = p.marginLeft - (pb.left?.space ?? 4) - (pb.left?.width ?? 0) / 2;
      y1 = s.marginTop - (pb.top?.space ?? 1) - (pb.top?.width ?? 0) / 2;
      x2 = p.width - p.marginRight + (pb.right?.space ?? 4) + (pb.right?.width ?? 0) / 2;
      y2 = p.height - s.marginBottom + (pb.bottom?.space ?? 1) + (pb.bottom?.width ?? 0) / 2;
    }
    const edge = (b: typeof pb.top, ax: number, ay: number, bx: number, by: number) => {
      if (!b) return;
      ops.push({ kind: 'line', x1: ax, y1: ay, x2: bx, y2: by, width: b.width, color: b.color ?? '000000', cap: 'square' });
    };
    edge(pb.top, x1, y1, x2, y1);
    edge(pb.bottom, x1, y2, x2, y2);
    edge(pb.left, x1, y1, x1, y2);
    edge(pb.right, x2, y1, x2, y2);
    return ops;
  }
}

function keepsWithNext(b: Block): boolean {
  return b.type === 'paragraph' && b.props.keepNext;
}
