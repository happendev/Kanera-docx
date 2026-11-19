import type { AnchorProps, Block, DocxDocument, Drawing, FieldInline, Paragraph, Section, Table } from '../docx/model.js';
import { formatNumber } from '../docx/numbering.js';
import type { FontRegistry } from '../fonts/registry.js';
import {
  gapBetween,
  hasDecorations,
  layoutBlocks,
  measureTable,
  paraDecorations,
  paraFrame,
  renderSegment,
  spaceAfter,
  type BoxEnv,
  type BoxItem,
  type BoxResult,
  type CellLayout,
  type MergeProgress,
  type RowPiece,
  type TableLayout,
} from './box.js';
import { renderDrawing } from './drawing.js';
import { ParagraphLayout, type Line } from './paragraph.js';
import { translateOps, type DrawOp, type Page } from './types.js';

const EPS = 0.01;
/** Narrowest gap beside a wrapped object that still receives text. */
const MIN_WRAP_WIDTH = 18;

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

interface Exclusion {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  wrap: AnchorProps['wrap'];
  side: AnchorProps['wrapText'];
}

interface FloatEntry {
  z: number;
  behind: boolean;
  ops: DrawOp[];
}

interface Column {
  x: number;
  w: number;
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
  notes: Array<{ id: string; noteType: 'footnote'; box: BoxResult }>;
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
  noteHeight: number;
  col: number;
  y: number;
  atTop: boolean;
  topKind: 'hard' | 'soft';
  prevPara: Paragraph | null;
  colBottoms: number[];
  colTop: number;
  cols: Column[];
  colSets: number;
  stats: LayoutStats;
}

const FIELD_FORMATS: Record<string, string> = {
  roman: 'lowerRoman',
  ROMAN: 'upperRoman',
  alphabetic: 'lowerLetter',
  ALPHABETIC: 'upperLetter',
  Arabic: 'decimal',
  arabic: 'decimal',
};

export function layoutDocument(doc: DocxDocument, fonts: FontRegistry): LayoutResult {
  return new Paginator(doc, fonts).run();
}

class Paginator {
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
  private floatedTables = new Set<Table>();
  /** Why the current column started: suppress space-before only after automatic breaks. */
  private topKind: 'hard' | 'soft' = 'hard';
  private prevPara: Paragraph | null = null;
  /** Index of the body block being placed (diagnostics). */
  private blockIndex = 0;
  private sectionBlockBase = 0;
  private section!: Section;
  private sectionIndex = 0;
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
    if (f.field === 'page') return formatNumber(page, fmt);
    if (f.field === 'numpages') return ctx?.total != null ? formatNumber(ctx.total, f.format ? fmt : 'decimal') : f.cached || '1';
    if (f.field === 'sectionpages') return ctx?.sectionTotal != null ? formatNumber(ctx.sectionTotal, f.format ? fmt : 'decimal') : f.cached || '1';
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
      this.startSection(section, i);
      this.placeBlocks(blocks);
      this.sectionBlockBase += blocks.length;
    }
    this.placeEndnotes();
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
    if (this.breakOpenedPage === this.pages.length && this.atTop && this.page.body.length === 0 && this.page.floats.length === 0 && this.pages.length > 1) {
      this.pages.pop();
      this.page = this.pages[this.pages.length - 1];
      this.y = Math.max(...this.page.colBottoms);
    }
    this.newPage(true);
    if (section.type === 'evenPage' || section.type === 'oddPage') {
      const wantOdd = section.type === 'oddPage';
      if ((this.page.number % 2 === 1) !== wantOdd) {
        // Word inserts a blank page to reach the right parity.
        this.newPage(true);
      }
    }
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

  private newPage(sectionStart: boolean): void {
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
    const headerBox = hf.header ? layoutBlocks(hf.header, hfWidth, this.env) : null;
    const footerBox = hf.footer ? layoutBlocks(hf.footer, hfWidth, this.env) : null;
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
          const pos = this.floatPosition(op.drawing, page, { x: marginLeft, w: hfWidth }, marginLeft + op.x, y0 + op.y);
          const ex = this.exclusionFor(op.drawing, pos.x, pos.y);
          if (ex && ex.y2 > top && ex.y1 < bottom) page.exclusions.push(ex);
        }
    }
    this.col = 0;
    this.y = top;
    this.atTop = true;
    this.topKind = 'hard';
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
      stats: { ...this.stats, breakLog: this.stats.breakLog.slice() },
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
    this.stats = { ...cp.stats, breakLog: cp.stats.breakLog.slice() };
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
        if (this.topKind === 'hard') this.y += gap.before;
      } else this.y += (this.prevPara ? gap.after : 0) + gap.before;
      this.y += frame.topExtra;
      let floatsPlaced = false;
      let linesHere = 0;
      let startedHere = true;
      let lineCps: Array<{ cp: Checkpoint; snap: ReturnType<ParagraphLayout['snapshot']> }> = [];
      let estimate = layout.markRun.ascent + layout.markRun.descent;
      for (;;) {
        if (!floatsPlaced) {
          const tables = this.tableFloats.get(para);
          this.placeFloats(tables ? [...tables, ...layout.floatDrawings] : layout.floatDrawings, this.column, this.y);
          floatsPlaced = true;
        }
        const snap = layout.snapshot();
        const lineCp = this.checkpoint();
        const placed = this.layoutLine(layout, estimate);
        if (!placed) break;
        const { line, y, left } = placed;
        estimate = line.height;
        const noteH = this.notesHeight(line.notes);
        // The extra space of "multiple" line spacing sits below the text, and Word lets it hang into
        // the bottom margin: only the text of the line has to fit. A footer tall enough to push the
        // body up is not margin, so above one the whole line box has to fit.
        const intoMargin = this.page!.bottom >= this.page!.height - this.page!.section.marginBottom - EPS;
        const need = intoMargin ? Math.min(line.height, line.baseline + line.descent) : line.height;
        const fits = y + need + noteH <= this.colBottom + EPS;
        if (!fits && !this.atTop) {
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
          const bottom = line.last ? y + line.height + frame.bottomExtra + nextGap.after + nextGap.before : y + line.height;
          this.page.body.push(...paraDecorations(para.props, frame, col.x, col.w, top, bottom, line.first, line.last));
        }
        this.page.body.push(...layout.renderLine(line, col.x + left, y));
        this.addNotes(line.notes);
        this.y = y + line.height;
        this.markPlaced(this.y);
        linesHere++;
        if (line.brk === 'page') {
          this.breakPage('hard');
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

  /**
   * Lay out the next line at the current position, flowing around wrapped objects. `left` is the
   * extra offset applied when rendering (always 0: bounds are passed in column coordinates).
   */
  private layoutLine(layout: ParagraphLayout, estimate: number): { line: Line; y: number; left: number } | null {
    const col = this.column;
    let y = this.y;
    let est = Math.max(estimate, 1);
    for (let iter = 0; iter < 50; iter++) {
      const base = layout.lineBounds(col.w);
      const slot = this.findSlot(y, est, col.x + base.left, col.x + base.right);
      const snap = layout.snapshot();
      const line = layout.nextLine(slot.left - col.x, slot.right - col.x, col.w);
      if (!line) return null;
      if (!this.page.exclusions.length || line.height <= est + EPS) return { line, y: slot.y, left: 0 };
      const check = this.findSlot(slot.y, line.height, col.x + base.left, col.x + base.right);
      if (Math.abs(check.y - slot.y) < EPS && Math.abs(check.left - slot.left) < EPS && Math.abs(check.right - slot.right) < EPS) return { line, y: slot.y, left: 0 };
      layout.restore(snap);
      est = line.height;
      y = slot.y;
    }
    const line = layout.nextLine(layout.lineBounds(col.w).left, layout.lineBounds(col.w).right, col.w);
    return line ? { line, y, left: 0 } : null;
  }

  /** The text interval for a line of height h at y between L and R (page x), avoiding exclusions. */
  private findSlot(y: number, h: number, L: number, R: number): { y: number; left: number; right: number } {
    const ex = this.page.exclusions;
    if (!ex.length) return { y, left: L, right: R };
    for (let iter = 0; iter < 100; iter++) {
      const hits = ex.filter((e) => e.y1 < y + h - EPS && e.y2 > y + EPS && e.x1 < R - EPS && e.x2 > L + EPS);
      if (!hits.length) return { y, left: L, right: R };
      const block = hits.filter((e) => e.wrap === 'topAndBottom');
      if (block.length) {
        y = Math.max(...block.map((e) => e.y2));
        continue;
      }
      let intervals: Array<[number, number]> = [[L, R]];
      for (const e of hits) {
        const next: Array<[number, number]> = [];
        for (const [a, b] of intervals) {
          let leftPart: [number, number] | null = e.x1 > a ? [a, Math.min(b, e.x1)] : null;
          let rightPart: [number, number] | null = e.x2 < b ? [Math.max(a, e.x2), b] : null;
          if (e.side === 'left') rightPart = null;
          else if (e.side === 'right') leftPart = null;
          else if (e.side === 'largest') {
            const lw = leftPart ? leftPart[1] - leftPart[0] : 0;
            const rw = rightPart ? rightPart[1] - rightPart[0] : 0;
            if (lw >= rw) rightPart = null;
            else leftPart = null;
          }
          if (leftPart && leftPart[1] - leftPart[0] > EPS) next.push(leftPart);
          if (rightPart && rightPart[1] - rightPart[0] > EPS) next.push(rightPart);
        }
        intervals = next;
      }
      const fit = intervals.find(([a, b]) => b - a >= MIN_WRAP_WIDTH);
      if (fit) return { y, left: fit[0], right: fit[1] };
      y = Math.min(...hits.map((e) => e.y2));
    }
    return { y, left: L, right: R };
  }

  // -------------------------------------------------------------------------------------------
  // Floating objects

  private floatPosition(d: Drawing, page: PageState, col: Column, paraX: number, paraY: number): { x: number; y: number } {
    const a = d.anchor!;
    const s = page.section;
    const w = d.width;
    const h = d.height;
    const h0 = (() => {
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
    if (ha === 'right' || ha === 'outside') x = h0.base + h0.size - w;
    else if (ha === 'center') x = h0.base + (h0.size - w) / 2;
    else if (ha === 'left' || ha === 'inside') x = h0.base;
    else if (a.posH.pct !== undefined) x = h0.base + a.posH.pct * h0.size;
    else x = h0.base + (a.posH.offset ?? 0);
    const v0 = (() => {
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
    if (va === 'bottom' || va === 'outside') y = v0.base + v0.size - h;
    else if (va === 'center') y = v0.base + (v0.size - h) / 2;
    else if (va === 'top' || va === 'inside') y = v0.base;
    else if (a.posV.pct !== undefined) y = v0.base + a.posV.pct * v0.size;
    else y = v0.base + (a.posV.offset ?? 0);
    return { x, y };
  }

  private exclusionFor(d: Drawing, x: number, y: number): Exclusion | null {
    const a = d.anchor;
    if (!a || a.wrap === 'none') return null;
    return {
      x1: x - d.effectL - a.distL,
      y1: y - d.effectT - a.distT,
      x2: x + d.width + d.effectR + a.distR,
      y2: y + d.height + d.effectB + a.distB,
      wrap: a.wrap,
      side: a.wrapText,
    };
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
    return {
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
  }

  private placeFloats(drawings: Drawing[], col: Column, paraY: number): void {
    for (const d of drawings) {
      if (!d.anchor || d.hidden) continue;
      const pos = this.floatPosition(d, this.page, col, col.x, paraY);
      this.page.floats.push({ z: d.anchor.zIndex, behind: d.anchor.behindDoc, ops: renderDrawing(d, pos.x, pos.y, this.env) });
      const ex = this.exclusionFor(d, pos.x, pos.y);
      if (ex) this.page.exclusions.push(ex);
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
      const pos = this.floatPosition(d, page, { x: op.x, w: op.areaWidth }, op.x, op.y);
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
      h += this.noteBox('footnote', n.id, width).height;
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
      if (this.page.notes.some((p) => p.id === n.id)) continue;
      if (this.page.notes.length === 0) this.page.noteHeight += this.footnoteSeparator(width).height;
      const box = this.noteBox('footnote', n.id, width);
      this.page.notes.push({ id: n.id, noteType: 'footnote', box });
      this.page.noteHeight += box.height;
    }
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
      for (const it of box.items) items.push({ ...it, y: it.y + base });
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
    if (!this.atTop && this.prevPara) this.y += spaceAfter(this.prevPara.props);
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
    const newColumn = () => {
      flush();
      this.breakColumn('soft');
      const c = this.column;
      if (Math.abs(c.w - col.w) > EPS) {
        col = c;
        tl = measureTable(table, col.w, this.env);
      }
      col = c;
      rowsOnPage = 0;
      // Header rows repeat at the top of each page.
      for (let h = 0; h < tl.headerRows; h++) {
        pieces.push({ r: h, y: this.y, height: tl.rows[h].height });
        this.y += tl.rows[h].height;
      }
    };
    const progress = new Map<CellLayout, number>();
    for (let r = 0; r < tl.rows.length; r++) {
      const row = tl.rows[r];
      let first = true;
      for (let guard = 0; guard < 500; guard++) {
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
        const emptyHere = this.atTop || rowsOnPage === 0;
        if (!emptyHere && (row.row.cantSplit || !this.canSplit(tl, r))) {
          if (first && r === 0) moved = true;
          newColumn();
          first = false;
          continue;
        }
        // Split the row at the column bottom.
        const split = this.splitRow(tl, r, progress, this.colBottom - this.y, emptyHere);
        if (!split) {
          if (first && r === 0) moved = true;
          newColumn();
          first = false;
          continue;
        }
        pieces.push({ r, y: this.y, height: split.height, slices: split.slices });
        this.y += split.height;
        this.markPlaced(this.y);
        rowsOnPage++;
        newColumn();
        first = false;
      }
    }
    flush();
    this.atTop = false;
    return moved;
  }

  private canSplit(tl: TableLayout, r: number): boolean {
    const row = tl.rows[r];
    if (row.row.cantSplit || row.row.heightRule === 'exact') return false;
    return row.cells.some((c) => c.rowSpan === 1 && !c.cell.vMerge && c.box.items.length > 1);
  }

  /** Height of what is left of row r after earlier split pieces. */
  private rowRemainingHeight(tl: TableLayout, r: number, progress: Map<CellLayout, number>): number {
    const row = tl.rows[r];
    if (!row.cells.some((c) => (progress.get(c) ?? 0) > 0)) return row.height;
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
      const inner = avail - c.cell.margins.top - c.cell.margins.bottom;
      let to = from;
      if (!c.cell.vMerge) {
        while (to < items.length && items[to].y - shift + items[to].height <= inner + EPS) to++;
        const fits = to;
        while (to > from && to < items.length && items[to].keepWithPrev) to--;
        if (to === from) to = force ? Math.max(fits, Math.min(from + 1, items.length)) : from;
      }
      if (to > from) advanced = true;
      slices.set(c, { from, to, shift });
      if (to > from) height = Math.max(height, items[to - 1].y + items[to - 1].height - shift + c.cell.margins.top + c.cell.margins.bottom);
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
    const sectionTotals = new Map<number, number>();
    for (const p of this.pages) sectionTotals.set(p.sectionIndex, (sectionTotals.get(p.sectionIndex) ?? 0) + 1);
    const out: Page[] = [];
    for (const p of this.pages) {
      const s = p.section;
      const hfOps: DrawOp[] = [];
      const floats = p.floats.slice();
      const hf = this.headerFooterFor(s, p.number, p.firstOfSection);
      const hfWidth = p.width - p.marginLeft - p.marginRight;
      this.fieldPage = { number: p.number, format: s.pageNumberFormat, total, sectionTotal: sectionTotals.get(p.sectionIndex) ?? null };
      if (hf.header) {
        const box = layoutBlocks(hf.header, hfWidth, this.env);
        for (const it of box.items) this.extractFloats(translateOps(it.ops, p.marginLeft, s.header), p, hfOps, floats);
      }
      if (hf.footer) {
        const box = layoutBlocks(hf.footer, hfWidth, this.env);
        const y0 = p.height - s.footer - box.height;
        for (const it of box.items) this.extractFloats(translateOps(it.ops, p.marginLeft, y0), p, hfOps, floats);
      }
      this.fieldPage = null;

      // Footnotes sit at the bottom of the body area.
      const body = p.body.slice();
      if (p.notes.length) {
        let y = p.bottom - p.noteHeight;
        const sep = this.footnoteSeparator(hfWidth);
        for (const it of sep.items) body.push(...translateOps(it.ops, p.marginLeft, y));
        y += sep.height;
        for (const n of p.notes) {
          for (const it of n.box.items) this.extractFloats(translateOps(it.ops, p.marginLeft, y), p, body, floats);
          y += n.box.height;
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
      const borders = this.pageBorders(p);
      if (borders.length) (s.pageBorders?.zOrder === 'back' ? behind : front).push(...borders);
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
