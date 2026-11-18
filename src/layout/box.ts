import type { Block, Border, Paragraph, ParaProps, Table, TableCell, TableRow } from '../docx/model.js';
import { ParagraphLayout, type Line, type ParagraphEnv } from './paragraph.js';
import { translateOps, type DrawOp } from './types.js';

/** A horizontal slice of laid-out content inside a box; pagination may split between items. */
export interface BoxItem {
  y: number;
  height: number;
  ops: DrawOp[];
  notes?: Line['notes'];
  /** Forced break after this item (page/column breaks inside table cells are ignored by Word). */
}

export interface BoxResult {
  height: number;
  items: BoxItem[];
}

export interface BoxEnv extends ParagraphEnv {
  compatibilityMode: number;
}

// ---------------------------------------------------------------------------------------------
// Paragraph spacing and borders

export function spaceBefore(p: ParaProps): number {
  return p.spaceBeforeAuto ? 14 : p.spaceBefore;
}
export function spaceAfter(p: ParaProps): number {
  return p.spaceAfterAuto ? 14 : p.spaceAfter;
}

/** Contextual spacing: no space between paragraphs of the same style when either asks for it. */
export function gapBetween(prev: Paragraph | null, next: Paragraph | null, prevAfterOverride?: number): { after: number; before: number } {
  let after = prev ? (prevAfterOverride ?? spaceAfter(prev.props)) : 0;
  let before = next ? spaceBefore(next.props) : 0;
  if (prev && next && prev.props.styleId === next.props.styleId) {
    if (prev.props.contextualSpacing) after = 0;
    if (next.props.contextualSpacing) before = 0;
  }
  // HTML auto spacing collapses like CSS margins.
  if (prev?.props.spaceAfterAuto && next?.props.spaceBeforeAuto) return { after: Math.max(after, before), before: 0 };
  return { after, before };
}

function sameBorder(a: Border | null, b: Border | null): boolean {
  if (!a || !b) return a === b;
  return a.style === b.style && a.width === b.width && a.color === b.color && a.space === b.space;
}

/** Paragraphs with identical borders and indents share one border box. */
export function bordersGroup(a: Paragraph | null, b: Paragraph | null): boolean {
  if (!a || !b) return false;
  const pa = a.props;
  const pb = b.props;
  const hasDeco = !!(pa.borders.top || pa.borders.bottom || pa.borders.left || pa.borders.right || pa.borders.between || pa.shading);
  if (!hasDeco) return false;
  return (
    sameBorder(pa.borders.top, pb.borders.top) &&
    sameBorder(pa.borders.bottom, pb.borders.bottom) &&
    sameBorder(pa.borders.left, pb.borders.left) &&
    sameBorder(pa.borders.right, pb.borders.right) &&
    sameBorder(pa.borders.between, pb.borders.between) &&
    pa.shading === pb.shading &&
    Math.abs(pa.indLeft - pb.indLeft) < 0.01 &&
    Math.abs(pa.indRight - pb.indRight) < 0.01
  );
}

export interface ParaFrame {
  groupedPrev: boolean;
  groupedNext: boolean;
  /** Space taken above the first line by the top (or between) border. */
  topExtra: number;
  bottomExtra: number;
}

export function paraFrame(prev: Paragraph | null, para: Paragraph, next: Paragraph | null): ParaFrame {
  const b = para.props.borders;
  const groupedPrev = bordersGroup(prev, para);
  const groupedNext = bordersGroup(para, next);
  const topB = groupedPrev ? b.between : b.top;
  const topExtra = topB ? topB.space + topB.width : 0;
  const bottomExtra = !groupedNext && b.bottom ? b.bottom.space + b.bottom.width : 0;
  return { groupedPrev, groupedNext, topExtra, bottomExtra };
}

export function hasDecorations(p: ParaProps): boolean {
  const b = p.borders;
  return !!(b.top || b.bottom || b.left || b.right || b.between || p.shading);
}

/** Border and shading ops for a vertical slice [top, bottom] of a paragraph. */
export function paraDecorations(p: ParaProps, frame: ParaFrame, x0: number, areaWidth: number, top: number, bottom: number, firstSlice: boolean, lastSlice: boolean): DrawOp[] {
  const ops: DrawOp[] = [];
  const b = p.borders;
  const left = x0 + p.indLeft - (b.left ? b.left.space + b.left.width : 0);
  const right = x0 + areaWidth - p.indRight + (b.right ? b.right.space + b.right.width : 0);
  if (p.shading) ops.push({ kind: 'rect', x: left, y: top, w: right - left, h: bottom - top, fill: p.shading });
  const line = (br: Border, x1: number, y1: number, x2: number, y2: number) => {
    const color = br.color ?? '000000';
    if (br.style === 'double') {
      const w = br.width / 3;
      const horiz = y1 === y2;
      const off = w;
      ops.push({ kind: 'line', x1: x1 - (horiz ? 0 : off), y1: y1 - (horiz ? off : 0), x2: x2 - (horiz ? 0 : off), y2: y2 - (horiz ? off : 0), width: w, color });
      ops.push({ kind: 'line', x1: x1 + (horiz ? 0 : off), y1: y1 + (horiz ? off : 0), x2: x2 + (horiz ? 0 : off), y2: y2 + (horiz ? off : 0), width: w, color });
    } else ops.push({ kind: 'line', x1, y1, x2, y2, width: br.width, color, dash: dashFor(br) });
  };
  if (b.left) line(b.left, left + b.left.width / 2, top, left + b.left.width / 2, bottom);
  if (b.right) line(b.right, right - b.right.width / 2, top, right - b.right.width / 2, bottom);
  if (firstSlice) {
    const tb = frame.groupedPrev ? b.between : b.top;
    if (tb) line(tb, left, top + tb.width / 2, right, top + tb.width / 2);
  }
  if (lastSlice && !frame.groupedNext && b.bottom) line(b.bottom, left, bottom - b.bottom.width / 2, right, bottom - b.bottom.width / 2);
  return ops;
}

export function dashFor(b: Border): number[] | undefined {
  const w = Math.max(b.width, 0.5);
  if (b.style === 'dotted') return [w, w * 2];
  if (b.style === 'dashed' || b.style === 'dashSmallGap') return [w * 3, w * 2];
  if (b.style === 'dotDash') return [w * 3, w * 2, w, w * 2];
  if (b.style === 'dotDotDash') return [w * 3, w * 2, w, w * 2, w, w * 2];
  return undefined;
}

// ---------------------------------------------------------------------------------------------
// Box layout (no pagination): table cells, headers/footers, footnotes, text boxes

export function layoutBlocks(blocks: Block[], width: number, env: BoxEnv): BoxResult {
  const items: BoxItem[] = [];
  let y = 0;
  let prevPara: Paragraph | null = null;
  let pendingAfter = 0;
  for (let bi = 0; bi < blocks.length; bi++) {
    const block = blocks[bi];
    if (block.type === 'paragraph') {
      const next = blocks[bi + 1]?.type === 'paragraph' ? (blocks[bi + 1] as Paragraph) : null;
      const gap = gapBetween(prevPara, block);
      y += prevPara ? gap.after + gap.before : gap.before;
      const frame = paraFrame(prevPara, block, next);
      const layout = new ParagraphLayout(block, env);
      const deco = hasDecorations(block.props);
      const lines: Line[] = [];
      let line: Line | null;
      y += frame.topExtra;
      while ((line = layoutNextLine(layout, width))) {
        const ops = layout.renderLine(line, 0, y);
        const item: BoxItem = { y, height: line.height, ops, notes: line.notes.length ? line.notes : undefined };
        // Floats are anchored to the paragraph, so they are positioned from its first line.
        if (lines.length === 0) for (const f of layout.floatDrawings) ops.push({ kind: 'float', drawing: f, x: 0, y, areaWidth: width });
        lines.push(line);
        items.push(item);
        y += line.height;
      }
      if (deco && items.length) {
        // Decorations are attached per line so the box can split between lines.
        const first = items.length - lines.length;
        const nextGap = frame.groupedNext && next ? gapBetween(block, next) : { after: 0, before: 0 };
        for (let k = first; k < items.length; k++) {
          const it = items[k];
          const isFirst = k === first;
          const isLast = k === items.length - 1;
          const top = isFirst ? it.y - frame.topExtra : it.y;
          const bottom = isLast ? it.y + it.height + frame.bottomExtra + nextGap.after + nextGap.before : it.y + it.height;
          it.ops.unshift(...paraDecorations(block.props, frame, 0, width, top, bottom, isFirst, isLast));
        }
      }
      y += frame.bottomExtra;
      pendingAfter = spaceAfter(block.props);
      prevPara = block;
    } else {
      // Space after the previous paragraph still applies before a table.
      if (prevPara) y += gapBetween(prevPara, null).after;
      const t = layoutTableInBox(block, width, env, y);
      items.push(...t.items);
      y = t.bottom;
      prevPara = null;
      pendingAfter = 0;
    }
  }
  if (prevPara) y += pendingAfter;
  return { height: y, items };
}

/** Next line using the paragraph's indents (no floats inside boxes). */
export function layoutNextLine(layout: ParagraphLayout, width: number): Line | null {
  if (layout.done) return null;
  const b = layout.lineBounds(width);
  return layout.nextLine(b.left, b.right, width);
}

// ---------------------------------------------------------------------------------------------
// Tables

export interface CellLayout {
  cell: TableCell;
  /** Index of the first grid column. */
  col: number;
  x: number;
  width: number;
  box: BoxResult;
  /** Rows spanned by a vertical merge (1 when not merged). */
  rowSpan: number;
}

export interface RowLayout {
  row: TableRow;
  cells: CellLayout[];
  height: number;
  /** Width of the row's top border, which content sits below. */
  borderTop: number;
}

export interface TableLayout {
  x: number;
  width: number;
  colX: number[];
  rows: RowLayout[];
  headerRows: number;
}

function gridWidths(table: Table, avail: number): number[] {
  let grid = table.grid.slice();
  const nCols = grid.length;
  let sum = grid.reduce((a, b) => a + b, 0);
  const w = table.props.width;
  if (nCols === 0) return grid;
  if (sum <= 0) {
    const target = w.type === 'dxa' && w.value > 0 ? w.value : w.type === 'pct' ? w.value * avail : avail;
    return grid.map(() => target / nCols);
  }
  // Fill zero-width columns from the remaining space.
  if (grid.some((g) => g <= 0)) {
    const zeros = grid.filter((g) => g <= 0).length;
    const fill = Math.max(10, (avail - sum) / zeros);
    grid = grid.map((g) => (g <= 0 ? fill : g));
    sum = grid.reduce((a, b) => a + b, 0);
  }
  if (w.type === 'pct' && w.value > 0) {
    const target = w.value * avail;
    if (Math.abs(target - sum) > 1) grid = grid.map((g) => (g * target) / sum);
  }
  return grid;
}

export function measureTable(table: Table, avail: number, env: BoxEnv): TableLayout {
  const grid = gridWidths(table, avail);
  const colX: number[] = [0];
  for (const g of grid) colX.push(colX[colX.length - 1] + g);
  const width = colX[colX.length - 1];
  let x = table.props.indent;
  if (table.props.align === 'center') x = (avail - width) / 2;
  else if (table.props.align === 'right') x = avail - width;
  else if (env.compatibilityMode < 15) {
    // Before Word 2013 the table indent positions the first cell's text, not its border.
    const firstCell = table.rows[0]?.cells[0];
    x -= firstCell ? firstCell.margins.left : table.props.cellMargins.left;
  }

  const rows: RowLayout[] = [];
  for (const row of table.rows) {
    const cells: CellLayout[] = [];
    let col = row.gridBefore;
    for (const cell of row.cells) {
      const c0 = Math.min(col, grid.length);
      const c1 = Math.min(col + cell.gridSpan, grid.length);
      const cx = colX[c0];
      const cw = Math.max(0, colX[c1] - cx);
      const inner = Math.max(1, cw - cell.margins.left - cell.margins.right);
      const box = cell.vMerge === 'continue' ? { height: 0, items: [] } : layoutBlocks(cell.content, inner, env);
      cells.push({ cell, col, x: cx, width: cw, box, rowSpan: 1 });
      col += cell.gridSpan;
    }
    let h = 0;
    // Word places cell content below the row's top border, so border width adds to the height.
    let borderTop = 0;
    for (const c of cells) borderTop = Math.max(borderTop, c.cell.borders.top?.width ?? 0);
    for (const c of cells) if (c.cell.vMerge !== 'restart' && c.cell.vMerge !== 'continue') h = Math.max(h, c.box.height + c.cell.margins.top + c.cell.margins.bottom);
    h += borderTop;
    // Rows with only merged cells still need room for their own (empty) content.
    if (cells.every((c) => c.cell.vMerge)) for (const c of cells) if (c.cell.vMerge === 'continue') h = Math.max(h, c.cell.margins.top + c.cell.margins.bottom);
    if (cells.length === 0) h = 0;
    if (row.heightRule === 'exact' && row.height > 0) h = row.height;
    else if (row.height > 0) h = Math.max(h, row.height);
    rows.push({ row, cells, height: h, borderTop });
  }
  // The table's bottom border sits below the last row's content.
  const last = rows[rows.length - 1];
  if (last) {
    let bb = 0;
    for (const c of last.cells) bb = Math.max(bb, c.cell.borders.bottom?.width ?? 0);
    if (last.row.heightRule !== 'exact') last.height += bb;
  }
  // Vertical merges: the restart cell spans following rows whose cell at the same column continues.
  for (let r = 0; r < rows.length; r++) {
    for (const c of rows[r].cells) {
      if (c.cell.vMerge !== 'restart') continue;
      let span = 1;
      while (r + span < rows.length && rows[r + span].cells.some((o) => o.col === c.col && o.cell.vMerge === 'continue')) span++;
      c.rowSpan = span;
      const need = c.box.height + c.cell.margins.top + c.cell.margins.bottom;
      let have = 0;
      for (let k = 0; k < span; k++) have += rows[r + k].height;
      if (need > have && rows[r + span - 1].row.heightRule !== 'exact') rows[r + span - 1].height += need - have;
    }
  }
  let headerRows = 0;
  while (headerRows < table.rows.length && table.rows[headerRows].isHeader) headerRows++;
  if (headerRows === table.rows.length) headerRows = 0;
  return { x, width, colX, rows, headerRows };
}

/** A row (or part of a split row) placed at `y`. */
export interface RowPiece {
  r: number;
  y: number;
  height: number;
  /** For a row split across pages: the item range drawn per cell, and how far content shifts up. */
  slices?: Map<CellLayout, { from: number; to: number; shift: number }>;
}

/** Progress through vertically merged cells whose content continues across pages. */
export type MergeProgress = Map<CellLayout, { from: number; shift: number }>;

function findRestart(tl: TableLayout, r: number, col: number): { r: number; cell: CellLayout } | null {
  for (let k = r - 1; k >= 0; k--) {
    const c = tl.rows[k].cells.find((o) => o.col === col);
    if (!c) return null;
    if (c.cell.vMerge === 'restart') return { r: k, cell: c };
    if (c.cell.vMerge !== 'continue') return null;
  }
  return null;
}

function drawEdge(out: DrawOp[], br: Border | null, x1: number, y1: number, x2: number, y2: number): void {
  if (!br || br.width <= 0) return;
  const color = br.color ?? '000000';
  if (br.style === 'double') {
    const w = Math.max(0.25, br.width / 3);
    const horiz = y1 === y2;
    out.push({ kind: 'line', x1: x1 - (horiz ? 0 : w), y1: y1 - (horiz ? w : 0), x2: x2 - (horiz ? 0 : w), y2: y2 - (horiz ? w : 0), width: w, color });
    out.push({ kind: 'line', x1: x1 + (horiz ? 0 : w), y1: y1 + (horiz ? w : 0), x2: x2 + (horiz ? 0 : w), y2: y2 + (horiz ? w : 0), width: w, color });
  } else out.push({ kind: 'line', x1, y1, x2, y2, width: br.width, color, dash: dashFor(br), cap: 'square' });
}

/**
 * Draw a run of row pieces that share a page. `ox` is the x of the layout area. Vertically merged
 * cells are drawn once across the pieces they span; their content continues on later segments
 * through `merge`.
 */
export function renderSegment(tl: TableLayout, pieces: RowPiece[], ox: number, merge: MergeProgress): { ops: DrawOp[]; notes: NonNullable<BoxItem['notes']> } {
  const back: DrawOp[] = [];
  const content: DrawOp[] = [];
  const borders: DrawOp[] = [];
  const notes: NonNullable<BoxItem['notes']> = [];
  for (let pi = 0; pi < pieces.length; pi++) {
    const p = pieces[pi];
    const rl = tl.rows[p.r];
    for (const c of rl.cells) {
      let cell = c;
      let spanFrom = p.r;
      if (c.cell.vMerge === 'continue') {
        const rs = findRestart(tl, p.r, c.col);
        // Drawn by the restart cell when the merge started earlier in this segment.
        const prev = pieces[pi - 1];
        if (rs && prev && prev.r >= rs.r && prev.r < p.r) continue;
        if (rs) {
          cell = rs.cell;
          spanFrom = rs.r;
        }
      }
      const spanEnd = cell.rowSpan > 1 ? spanFrom + cell.rowSpan : p.r + 1;
      let height = p.height;
      let last = pi;
      while (last + 1 < pieces.length && pieces[last + 1].r < spanEnd && pieces[last + 1].r > pieces[last].r) {
        last++;
        height += pieces[last].height;
      }
      const isMerged = cell.rowSpan > 1;
      const x = ox + tl.x + cell.x;
      const top = p.y;
      if (cell.cell.shading) back.push({ kind: 'rect', x, y: top, w: cell.width, h: height, fill: cell.cell.shading });
      const m = cell.cell.margins;
      const innerX = x + m.left;
      const innerH = height - m.top - m.bottom;
      const items = cell.box.items;
      let from = 0;
      let to = items.length;
      let shift = 0;
      let align = true;
      const slice = !isMerged ? p.slices?.get(cell) : undefined;
      if (slice) {
        ({ from, to, shift } = slice);
        align = false;
      } else if (isMerged) {
        const st = merge.get(cell) ?? { from: 0, shift: 0 };
        from = st.from;
        shift = st.shift;
        to = from;
        while (to < items.length && items[to].y - shift + items[to].height <= innerH + 0.01) to++;
        // The span ends in this segment: draw whatever is left even if it overflows.
        if (last === pieces.length - 1 && spanEnd - 1 > pieces[last].r) {
          if (to === from && to < items.length) to++;
        } else to = items.length;
        merge.set(cell, { from: to, shift: to < items.length ? items[to].y : 0 });
        align = from === 0 && to === items.length;
      }
      let dy = top + m.top - shift + (from === 0 ? tl.rows[spanFrom].borderTop : 0);
      if (align && from === 0 && to === items.length) {
        const free = innerH - cell.box.height;
        if (free > 0 && cell.cell.vAlign === 'center') dy += free / 2;
        else if (free > 0 && cell.cell.vAlign === 'bottom') dy += free;
      }
      for (let k = from; k < to; k++) {
        content.push(...translateOps(items[k].ops, innerX, dy));
        if (items[k].notes) notes.push(...items[k].notes!);
      }
      const b = cell.cell.borders;
      const lastRow = tl.rows[Math.min(spanEnd, tl.rows.length) - 1];
      const bottomCell = lastRow?.cells.find((o) => o.col === cell.col);
      const bottomB = bottomCell ? bottomCell.cell.borders.bottom : b.bottom;
      drawEdge(borders, b.top, x, top, x + cell.width, top);
      drawEdge(borders, bottomB, x, top + height, x + cell.width, top + height);
      drawEdge(borders, b.left, x, top, x, top + height);
      drawEdge(borders, b.right, x + cell.width, top, x + cell.width, top + height);
      drawEdge(borders, b.tl2br, x, top, x + cell.width, top + height);
      drawEdge(borders, b.tr2bl, x + cell.width, top, x, top + height);
    }
  }
  return { ops: [...back, ...content, ...borders], notes };
}

/** Tables inside boxes: one item per group of rows tied together by vertical merges. */
function layoutTableInBox(table: Table, width: number, env: BoxEnv, y0: number): { items: BoxItem[]; bottom: number } {
  const tl = measureTable(table, width, env);
  const items: BoxItem[] = [];
  const merge: MergeProgress = new Map();
  let y = y0;
  let r = 0;
  while (r < tl.rows.length) {
    let end = r + 1;
    for (let k = r; k < end; k++) for (const c of tl.rows[k].cells) if (c.rowSpan > 1) end = Math.max(end, k + c.rowSpan);
    const pieces: RowPiece[] = [];
    let gy = y;
    for (let k = r; k < end && k < tl.rows.length; k++) {
      pieces.push({ r: k, y: gy, height: tl.rows[k].height });
      gy += tl.rows[k].height;
    }
    const seg = renderSegment(tl, pieces, 0, merge);
    items.push({ y, height: gy - y, ops: seg.ops, notes: seg.notes.length ? seg.notes : undefined });
    y = gy;
    r = end;
  }
  return { items, bottom: y };
}
