import type { AnchorProps, Block, Border, Drawing, FrameProps, Paragraph, ParaProps, Table, TableCell, TableRow } from '../docx/model.js';
import { ParagraphLayout, type Line, type ParagraphEnv } from './paragraph.js';
import { translateOps, type DrawOp } from './types.js';

/** Start of an unconsumed line or nested table row in the original cell. */
export interface CellCursor {
  block: number;
  line?: ReturnType<ParagraphLayout['snapshot']>;
  row?: number;
}

/** A horizontal slice of laid-out content inside a box; pagination may split between items. */
export interface BoxItem {
  cursor?: CellCursor;
  y: number;
  height: number;
  ops: DrawOp[];
  notes?: Line['notes'];
  /** A table row split may not fall just before this line (widow/orphan control, keep lines together). */
  keepWithPrev?: boolean;
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
  // Automatic (HTML) spacing leaves no gap between items of one list, as between HTML <li>
  // elements: only the space around the list remains.
  if (prev?.numbering && prev.numbering.numId === next?.numbering?.numId) {
    if (prev.props.spaceAfterAuto) after = 0;
    if (next.props.spaceBeforeAuto) before = 0;
  }
  // Word does not add one paragraph's space after to the next one's space before: the larger of
  // the two separates them (measured against Word's own PDF output and its recorded page breaks).
  if (prev && next) return { after: Math.max(after, before), before: 0 };
  return { after, before };
}

/**
 * Space left after a paragraph that a table follows. Contextual spacing reaches into the table: a
 * paragraph of the same style opening its first cell drops it (as in Word's PDFs).
 */
export function spaceAboveTable(prev: Paragraph, table: Table): number {
  const first = table.rows[0]?.cells[0]?.content[0];
  if (prev.props.contextualSpacing && first?.type === 'paragraph' && first.props.styleId === prev.props.styleId) return 0;
  return spaceAfter(prev.props);
}

function sameBorder(a: Border | null, b: Border | null): boolean {
  if (!a || !b) return a === b;
  return a.style === b.style && a.width === b.width && a.color === b.color && a.space === b.space;
}

/** Paragraphs with identical borders and indents share one border box. */
function boxLeft(p: ParaProps): number {
  return p.indLeft + Math.min(0, p.indFirstLine);
}

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
    // The box spans from the leftmost of the left and first-line indents, so a hanging-indent
    // paragraph joins a plain one whose text starts where its first line does.
    Math.abs(boxLeft(pa) - boxLeft(pb)) < 0.01 &&
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

/**
 * Space below a paragraph's last line that its shading covers. Word closes the gap between two
 * shaded paragraphs, even in different colours, with the upper one's shading. Above unshaded text
 * or a table the space after stays bare, as it does where a bottom border closes the box first or
 * contextual spacing drops the space. A grouped paragraph's shading already runs on into the next.
 */
export function shadedSpaceAfter(para: Paragraph, frame: ParaFrame, next: Paragraph | null): number {
  const p = para.props;
  if (!p.shading || !next?.props.shading || frame.groupedNext || p.borders.bottom) return 0;
  if (p.contextualSpacing && next.props.styleId === p.styleId) return 0;
  return spaceAfter(p);
}

/**
 * How far a paragraph's side borders run on below the rest of its box. With no bottom border to
 * close the box, Word carries them down through the space after, which the shading leaves bare
 * unless the next paragraph is shaded too.
 */
export function sideBorderRunOn(para: Paragraph, frame: ParaFrame, next: Paragraph | null): number {
  const p = para.props;
  if (!(p.borders.left || p.borders.right) || p.borders.bottom || frame.groupedNext) return 0;
  if (p.contextualSpacing && next?.props.styleId === p.styleId) return 0;
  return spaceAfter(p) - shadedSpaceAfter(para, frame, next);
}

export function hasDecorations(p: ParaProps): boolean {
  const b = p.borders;
  return !!(b.top || b.bottom || b.left || b.right || b.between || p.shading);
}

/** Word's paragraph box, shading and borders alike, reaches this far past the text on either side. */
const BOX_PAD = 1.5;

/** Border and shading ops for a vertical slice [top, bottom] of a paragraph. */
export function paraDecorations(p: ParaProps, frame: ParaFrame, x0: number, areaWidth: number, top: number, bottom: number, firstSlice: boolean, lastSlice: boolean, sideRunOn = 0): DrawOp[] {
  const ops: DrawOp[] = [];
  const b = p.borders;
  const left = x0 + p.indLeft - (b.left ? b.left.space + b.left.width : 0) - BOX_PAD;
  const right = x0 + areaWidth - p.indRight + (b.right ? b.right.space + b.right.width : 0) + BOX_PAD;
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
  const sideBottom = lastSlice ? bottom + sideRunOn : bottom;
  if (b.left) line(b.left, left + b.left.width / 2, top, left + b.left.width / 2, sideBottom);
  if (b.right) line(b.right, right - b.right.width / 2, top, right - b.right.width / 2, sideBottom);
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
// Text wrapping around floating objects

/** The area a wrapped floating object keeps text out of. */
export interface Exclusion {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  wrap: AnchorProps['wrap'];
  side: AnchorProps['wrapText'];
  /**
   * Tight and through wrapping: the object's outline, as points from (x1, y1), and the distances
   * kept from text around it. Text then wraps to the outline's extent across each line; x1 to y2
   * stay the object's box.
   */
  poly?: { pts: Array<[number, number]>; padL: number; padT: number; padR: number; padB: number };
}

const EPS = 0.01;
/** Narrowest gap beside a wrapped object that still receives text. */
const MIN_WRAP_WIDTH = 18;

/**
 * The text interval for a line of height h at y between L and R, avoiding exclusions. Outlines
 * meet only the line's text, `text` high (Word's tight wrapping ignores the extra space of
 * "multiple" line spacing below it); boxes meet the whole line.
 */
export function findSlot(ex: Exclusion[], y: number, h: number, L: number, R: number, text = h): { y: number; left: number; right: number } {
  if (!ex.length) return { y, left: L, right: R };
  for (let iter = 0; iter < 100; iter++) {
    const hits = bandHits(ex, y, h, text, L, R);
    if (!hits.length) return { y, left: L, right: R };
    const block = hits.filter((e) => e.wrap === 'topAndBottom');
    if (block.length) {
      y = Math.max(...block.map((e) => e.y2));
      continue;
    }
    const fit = gaps(hits, L, R).find(([a, b]) => b - a >= MIN_WRAP_WIDTH);
    if (fit) return { y, left: fit[0], right: fit[1] };
    y = Math.min(...hits.map((e) => e.next ?? e.y2));
  }
  return { y, left: L, right: R };
}

/**
 * The next gap right of x on the line band from y to y + h. Word fills each gap that objects
 * wrapped on both sides leave beside a line, left to right, before it moves down a line.
 */
export function gapBeside(ex: Exclusion[], y: number, h: number, text: number, L: number, R: number, x: number): { left: number; right: number } | null {
  const hits = bandHits(ex, y, h, text, L, R);
  if (!hits.length || hits.some((e) => e.wrap === 'topAndBottom')) return null;
  const fit = gaps(hits, L, R).find(([a, b]) => a > x + EPS && b - a >= MIN_WRAP_WIDTH);
  return fit ? { left: fit[0], right: fit[1] } : null;
}

/** The exclusions a line at y, h high with text `text` high, between L and R meets. */
function bandHits(ex: Exclusion[], y: number, h: number, text: number, L: number, R: number): Array<Exclusion & { next?: number }> {
  return ex
    .map((e) => (e.poly ? outlineBand(e, y, h, text) : e))
    .filter((e): e is Exclusion & { next?: number } => !!e && e.y1 < y + h - EPS && e.y2 > y + EPS && e.x1 < R - EPS && e.x2 > L + EPS);
}

/** The parts of L to R that the side-wrapped exclusions leave to text. */
function gaps(hits: Exclusion[], L: number, R: number): Array<[number, number]> {
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
  return intervals;
}

/**
 * The box an outline-wrapped object keeps text out of beside a line at y, h high with text `text`
 * high: the outline's horizontal extent across the text. Null where the text misses the outline.
 * A line that fits beside none of it moves down a line (`next`), as Word skips a line there.
 */
function outlineBand(e: Exclusion, y: number, h: number, text: number): (Exclusion & { next: number }) | null {
  const { pts, padL, padT, padR, padB } = e.poly!;
  // The distances above and below the outline widen the band it is measured across.
  const ya = y - padB;
  const yb = y + text + padT;
  let lo = Infinity;
  let hi = -Infinity;
  for (let i = 0; i < pts.length; i++) {
    const [px, py] = [e.x1 + pts[i][0], e.y1 + pts[i][1]];
    const [qx, qy] = [e.x1 + pts[(i + 1) % pts.length][0], e.y1 + pts[(i + 1) % pts.length][1]];
    if (Math.max(py, qy) < ya || Math.min(py, qy) > yb) continue;
    if (Math.abs(qy - py) < EPS) {
      lo = Math.min(lo, px, qx);
      hi = Math.max(hi, px, qx);
      continue;
    }
    // The part of the edge inside the band.
    const ta = Math.max(0, Math.min(1, (ya - py) / (qy - py)));
    const tb = Math.max(0, Math.min(1, (yb - py) / (qy - py)));
    for (const t of [ta, tb]) {
      const x = px + (qx - px) * t;
      lo = Math.min(lo, x);
      hi = Math.max(hi, x);
    }
  }
  if (lo > hi) return null;
  return { x1: lo - padL, x2: hi + padR, y1: y, y2: y + text, wrap: e.wrap, side: e.side, next: y + Math.max(h, 1) };
}

/** A line's text height, without the extra space "multiple" line spacing adds below. */
export function textHeight(line: Line): number {
  return Math.min(line.height, line.baseline + line.descent);
}

/**
 * The next line of a paragraph at or below y in the column (x, w), flowing around the exclusions.
 * The line's height decides which objects it meets, so a line taller than estimated is retried.
 * With `gap`, the line continues one at y in that gap beside an object (`bottom` is the earlier
 * part's bottom). `right` is the right edge of the space the line had.
 */
export function wrapLine(
  layout: ParagraphLayout,
  ex: Exclusion[],
  y: number,
  estimate: { h: number; text: number },
  col: { x: number; w: number },
  gap?: { left: number; right: number; bottom: number },
): { line: Line; y: number; right: number } | null {
  let est = { h: Math.max(estimate.h, 1), text: Math.max(estimate.text, 1) };
  const y0 = y;
  for (let iter = 0; iter < 50; iter++) {
    const base = layout.lineBounds(col.w);
    const L = gap ? Math.max(gap.left, col.x + base.left) : col.x + base.left;
    const R = gap ? Math.min(gap.right, col.x + base.right) : col.x + base.right;
    const slot = findSlot(ex, y, est.h, L, R, est.text);
    // A gap beside an earlier part of the line holds only on that line: text that no longer fits
    // there goes below the line.
    if (gap && slot.y > y + EPS) {
      y = gap.bottom;
      gap = undefined;
      continue;
    }
    const snap = layout.snapshot();
    // A first line pushed right by an object keeps its first-line indent from the object's edge.
    let left = slot.left;
    if (slot.left > col.x + base.left + EPS) left = Math.min(slot.right, slot.left + Math.max(0, base.left - layout.para.props.indLeft));
    const line = layout.nextLine(left - col.x, slot.right - col.x, col.w);
    if (!line) return null;
    // A line holding nothing but a page or column break takes no room beside objects: it stays
    // where it is rather than move below one (and push its break onto the next page).
    if (slot.y > y0 + EPS && !gap && (line.brk === 'page' || line.brk === 'column') && line.contentRight <= line.left + EPS) {
      layout.restore(snap);
      const bounds = layout.lineBounds(col.w);
      const kept = layout.nextLine(bounds.left, bounds.right, col.w);
      return kept && { line: kept, y: y0, right: col.x + bounds.right };
    }
    const text = textHeight(line);
    // A line moved down past an outline moved a line at a time, so it searches again with its
    // real height.
    const moved = slot.y > y0 + EPS && !gap && Math.abs(line.height - est.h) > EPS && ex.some((e) => e.poly);
    if (!ex.length || (!moved && line.height <= est.h + EPS && text <= est.text + EPS)) return { line, y: slot.y, right: slot.right };
    if (!moved) {
      const check = findSlot(ex, slot.y, line.height, L, R, text);
      if (Math.abs(check.y - slot.y) < EPS && Math.abs(check.left - slot.left) < EPS && Math.abs(check.right - slot.right) < EPS) return { line, y: slot.y, right: slot.right };
    }
    layout.restore(snap);
    est = { h: line.height, text };
    y = moved ? y0 : slot.y;
  }
  const bounds = layout.lineBounds(col.w);
  const line = layout.nextLine(bounds.left, bounds.right, col.w);
  return line ? { line, y, right: col.x + bounds.right } : null;
}

// ---------------------------------------------------------------------------------------------
// Box layout (no pagination): table cells, headers/footers, footnotes, text boxes

/**
 * The area a floating object anchored to a paragraph at box y keeps text out of, in box
 * coordinates; null when it doesn't wrap text.
 */
export type BoxWrap = (d: Drawing, y: number) => Exclusion | null;

/**
 * Lay out blocks in a box. With `wrap`, text flows around the box's own floating objects; `cell`
 * marks a table cell's content.
 */
export function layoutBlocks(blocks: Block[], width: number, env: BoxEnv, opts: { wrap?: BoxWrap; cell?: boolean; trackCursor?: boolean; start?: CellCursor } = {}): BoxResult {
  const wrap = opts.wrap;
  const items: BoxItem[] = [];
  const exclusions: Exclusion[] = [];
  // Top-and-bottom objects placed on the page rather than from their paragraph.
  const banners: Exclusion[] = [];
  let y = 0;
  let prevPara: Paragraph | null = null;
  let pendingAfter = 0;
  let pendingFrame: TextFrame | null = null;
  let inFlow = false;
  // A text frame takes no room in the flow: it sits relative to the next paragraph's top.
  const placeFrame = (top: number) => {
    if (!pendingFrame) return;
    const f = pendingFrame;
    const fy = top + (f.props.y ?? 0);
    items.push({ y: fy, height: 0, ops: translateOps(f.box.items.flatMap((it) => it.ops), f.x, fy) });
    if (f.props.wrap !== 'none' && f.props.wrap !== 'through') {
      exclusions.push({ x1: f.x - f.props.hSpace, x2: f.x + f.width + f.props.hSpace, y1: fy - f.props.vSpace, y2: fy + f.box.height + f.props.vSpace, wrap: f.props.wrap === 'notBeside' ? 'topAndBottom' : 'square', side: 'bothSides' });
    }
    pendingFrame = null;
  };
  for (let bi = opts.start?.block ?? 0; bi < blocks.length; bi++) {
    const block = blocks[bi];
    if (block.type === 'paragraph' && !opts.cell && isBoxFrame(block)) {
      let end = bi + 1;
      while (end < blocks.length && blocks[end].type === 'paragraph' && sameFrame((blocks[end] as Paragraph).props.frame, block.props.frame)) end++;
      pendingFrame = layoutFrame(blocks.slice(bi, end) as Paragraph[], width, env);
      bi = end - 1;
      continue;
    }
    // A framed paragraph is out of the flow wherever its frame sits; an empty one draws nothing
    // (a footer of framed page numbers takes no height).
    if (block.type === 'paragraph' && !opts.cell && block.props.frame && !block.props.frame.dropCap && block.content.length === 0 && !block.numbering) continue;
    if (block.type === 'paragraph') {
      const next = blocks[bi + 1]?.type === 'paragraph' ? (blocks[bi + 1] as Paragraph) : null;
      const gap = gapBetween(prevPara, block);
      // Automatic (HTML) space before collapses into the top of a header, cell or text box.
      y += prevPara ? gap.after + gap.before : !inFlow && block.props.spaceBeforeAuto ? 0 : gap.before;
      inFlow = true;
      const frame = paraFrame(prevPara, block, next);
      const layout = new ParagraphLayout(block, env);
      const resume = bi === opts.start?.block ? opts.start.line : undefined;
      if (resume) layout.restore(resume);
      const continuing = !!resume && !resume.firstLine;
      const deco = hasDecorations(block.props);
      const lines: Line[] = [];
      let line: Line | null;
      if (continuing) y = 0;
      else y += frame.topExtra;
      // Floats are anchored to the paragraph, so they are positioned from its top.
      const paraY = y;
      placeFrame(paraY);
      if (wrap) {
        for (const f of layout.floatDrawings) {
          const ex = f.hidden ? null : wrap(f, paraY);
          if (!ex) continue;
          exclusions.push(ex);
          const v = f.anchor!.posV.relativeFrom;
          if (ex.wrap === 'topAndBottom' && v !== 'paragraph' && v !== 'line') banners.push(ex);
        }
      }
      let estimate = layout.markRun.ascent + layout.markRun.descent;
      for (;;) {
        const snap = exclusions.length || opts.trackCursor ? layout.snapshot() : null;
        line = layoutNextLine(layout, width);
        if (!line) break;
        // Lines with text flow around the objects; Word leaves an empty line where it is (a logo
        // over a header's blank paragraphs doesn't make the header taller), unless a banner placed
        // on the page and wrapped above and below covers it.
        const around = line.contentRight > line.left ? exclusions : banners;
        if (snap && around.length) {
          layout.restore(snap);
          const placed = wrapLine(layout, around, y, { h: estimate, text: estimate }, { x: 0, w: width });
          if (!placed) break;
          line = placed.line;
          y = placed.y;
        }
        estimate = line.height;
        const ops = layout.renderLine(line, 0, y);
        const item: BoxItem = { y, height: line.height, ops, notes: line.notes.length ? line.notes : undefined };
        if (opts.trackCursor && snap) item.cursor = { block: bi, line: snap };
        if (!continuing && lines.length === 0) for (const f of layout.floatDrawings) ops.push({ kind: 'float', drawing: f, x: 0, y: paraY, areaWidth: width });
        lines.push(line);
        items.push(item);
        y += line.height;
      }
      // Word applies widow/orphan control and keep-lines-together when it splits a table row across
      // pages: a two- or three-line paragraph in a cell moves whole rather than leave one line behind.
      const n = lines.length;
      for (let k = 1; k < n; k++) {
        if (block.props.keepLines || (block.props.widowControl && (k === 1 || k === n - 1))) items[items.length - n + k].keepWithPrev = true;
      }
      if (deco && items.length) {
        // Decorations are attached per line so the box can split between lines.
        const first = items.length - lines.length;
        const nextGap = frame.groupedNext && next ? gapBetween(block, next) : { after: 0, before: 0 };
        for (let k = first; k < items.length; k++) {
          const it = items[k];
          const isFirst = !continuing && k === first;
          const isLast = k === items.length - 1;
          const top = isFirst ? it.y - frame.topExtra : it.y;
          const bottom = isLast ? it.y + it.height + frame.bottomExtra + nextGap.after + nextGap.before + shadedSpaceAfter(block, frame, next) : it.y + it.height;
          it.ops.unshift(...paraDecorations(block.props, frame, 0, width, top, bottom, isFirst, isLast, sideBorderRunOn(block, frame, next)));
        }
      }
      y += frame.bottomExtra;
      pendingAfter = spaceAfter(block.props);
      prevPara = block;
    } else {
      // Space after the previous paragraph still applies before a table.
      if (prevPara) y += spaceAboveTable(prevPara, block);
      placeFrame(y);
      inFlow = true;
      const t = layoutTableInBox(block, width, env, y, !!opts.cell, opts.trackCursor ? bi : undefined, bi === opts.start?.block ? opts.start.row : undefined);
      items.push(...t.items);
      y = t.bottom;
      prevPara = null;
      pendingAfter = 0;
    }
  }
  placeFrame(y);
  // Automatic (HTML) space after collapses into the bottom of a table cell. So does contextual
  // spacing's in the default paragraph style: Word compares a cell's last paragraph with the
  // row-end mark, which has that style (a List Paragraph or custom style keeps its space).
  const contextualEnd = !!prevPara?.props.contextualSpacing && prevPara.props.styleId === env.settings.defaultParagraphStyle;
  if (prevPara && !(opts.cell && (prevPara.props.spaceAfterAuto || contextualEnd))) y += pendingAfter;
  return { height: y, items };
}

interface TextFrame {
  props: FrameProps;
  /** Left edge in box coordinates. */
  x: number;
  width: number;
  box: BoxResult;
}

/**
 * A text frame (w:framePr) positioned against the box and the following paragraph, like Word's page
 * number frames in headers and footers. Other frames (drop caps, inline frames, frames placed on
 * the page) flow as ordinary paragraphs.
 */
function isBoxFrame(p: Paragraph): boolean {
  const f = p.props.frame;
  return !!f && !f.dropCap && f.yAlign !== 'inline' && f.vAnchor === 'text' && (f.hAnchor === 'margin' || f.hAnchor === 'text');
}

function sameFrame(a: FrameProps | null, b: FrameProps | null): boolean {
  if (!a || !b) return false;
  const keys = Object.keys(a) as Array<keyof FrameProps>;
  return keys.length === Object.keys(b).length && keys.every((k) => a[k] === b[k]);
}

/** Lays out a frame's paragraphs; a frame without a width is as wide as its widest line. */
function layoutFrame(paras: Paragraph[], width: number, env: BoxEnv): TextFrame {
  const props = paras[0].props.frame!;
  let w = props.width && props.width > 0 ? props.width : 0;
  if (!w) {
    for (const p of paras) {
      const layout = new ParagraphLayout(p, env);
      for (let line = layoutNextLine(layout, width); line; line = layoutNextLine(layout, width)) w = Math.max(w, line.contentRight + width - line.right);
    }
    w = Math.min(Math.max(w, 1), width);
  }
  const box = layoutBlocks(paras.map((p) => ({ ...p, props: { ...p.props, frame: null } })), w, env);
  const align = props.xAlign;
  const x = align === 'right' || align === 'outside' ? width - w : align === 'center' ? (width - w) / 2 : align ? 0 : props.x ?? 0;
  return { props, x, width: w, box };
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
  /** Text turned to run along the row (w:textDirection), and the line length it was laid out at. */
  turned?: { dir: TurnedText; length: number };
}

/**
 * Text that runs up the cell (btLr), down it from the right edge (tbRl), or down it from the left
 * edge (tbLrV). The glyphs of the last two face the same way; only where the lines start differs.
 */
type TurnedText = 'up' | 'down' | 'downFromLeft';

function turnedText(dir: string): TurnedText | undefined {
  // Transitional names, then their Strict equivalents.
  if (dir === 'btLr' || dir === 'lr') return 'up';
  if (dir === 'tbRl' || dir === 'tbRlV' || dir === 'rl' || dir === 'rlV') return 'down';
  if (dir === 'tbLrV' || dir === 'lrV') return 'downFromLeft';
  return undefined;
}

export interface RowLayout {
  row: TableRow;
  cells: CellLayout[];
  height: number;
  /** Width of the row's top border, which content sits below. */
  borderTop: number;
  /** Cell spacing above and below the row's cells. */
  spaceTop: number;
  spaceBottom: number;
  /** Room below the table's last row for its bottom border, which lies on the content's edge. */
  borderBottom?: number;
}

export interface TableLayout {
  x: number;
  width: number;
  colX: number[];
  rows: RowLayout[];
  headerRows: number;
  /** A right-to-left table (w:bidiVisual): its first column on the right, cells' left and right sides their start and end. */
  rtl?: boolean;
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

/**
 * How far below the top of a cell's content its wrapped objects laid out in the cell reach: Word
 * grows the row to hold them. Objects aligned within the cell, and those in nested tables, don't count.
 */
function inCellFloatsBottom(box: BoxResult, marginTop: number): number {
  let bottom = 0;
  for (const it of box.items) {
    for (const op of it.ops) {
      if (op.kind !== 'float' || op.cell) continue;
      const d = op.drawing;
      const a = d.anchor;
      if (!a || d.hidden || !a.layoutInCell || a.wrap === 'none' || a.posV.align || a.posV.pct !== undefined) continue;
      const rel = a.posV.relativeFrom;
      const top = rel === 'paragraph' || rel === 'line' ? op.y : rel === 'page' || rel === 'margin' ? -marginTop : null;
      if (top === null) continue;
      // A text-wrapping object cannot rise above its cell, so one raised past the top pushes the
      // row down by as much.
      bottom = Math.max(bottom, Math.max(-marginTop, top + (a.posV.offset ?? 0)) + d.height + d.effectB + a.distB);
    }
  }
  return bottom;
}

/** Word's widest table column: 22 inches. */
const MAX_COLUMN = 1584;

/** Narrowest and widest a cell's content can be laid out without breaking words (`hi` 0 with `loOnly`). */
function cellContentWidths(cell: TableCell, env: BoxEnv, loOnly = false): { lo: number; hi: number } {
  let lo = 0;
  let hi = 0;
  const wide = 2 * MAX_COLUMN;
  for (const b of cell.content) {
    if (b.type === 'table') {
      // A nested autofit table shrinks to fit its cell (see measureTable), so only a fixed one
      // holds the column open.
      const w = b.grid.reduce((a, g) => a + g, 0) + Math.max(0, b.props.indent);
      if (b.props.layout === 'fixed') lo = Math.max(lo, w);
      hi = Math.max(hi, w);
      continue;
    }
    const layout = new ParagraphLayout(b, env);
    lo = Math.max(lo, layout.minContentWidth());
    if (loOnly) continue;
    for (let line = layoutNextLine(layout, wide); line; line = layoutNextLine(layout, wide)) hi = Math.max(hi, line.contentRight + wide - line.right);
  }
  return { lo, hi: Math.max(lo, hi) };
}

/**
 * The grid Word lays an autofit table of automatic width out on when it sizes the columns from
 * their content (it does for files other tools wrote; Word's own files carry the grid it computed).
 * Each column gets at least its longest word and at most its longest line, or a cell's preferred
 * width in place of that line; when the widest layout doesn't fit, the room above the minimums is
 * shared in proportion to how much more each column wants.
 */
function autofitGrid(table: Table, maxWidth: number, env: BoxEnv): number[] {
  const n = table.grid.length;
  const sp = table.rows[0]?.cellSpacing ?? 0;
  const min = new Array<number>(n).fill(0);
  const max = new Array<number>(n).fill(0);
  const spans: Array<{ from: number; to: number; min: number; max: number }> = [];
  for (const row of table.rows) {
    let col = row.gridBefore;
    for (const cell of row.cells) {
      const from = col;
      const to = Math.min(n, col + cell.gridSpan);
      col += cell.gridSpan;
      if (from >= n || cell.vMerge === 'continue') continue;
      const { lo, hi } = cellContentWidths(cell, env);
      const sides = cell.margins.left + cell.margins.right;
      const pref = cell.width.type === 'dxa' && cell.width.value > 0 ? cell.width.value : 0;
      const cellMin = lo + sides;
      const cellMax = Math.max(cellMin, pref || hi + sides);
      if (to - from === 1) {
        min[from] = Math.max(min[from], cellMin);
        max[from] = Math.max(max[from], cellMax);
      } else spans.push({ from, to, min: cellMin, max: cellMax });
    }
  }
  // A spanning cell widens the columns it covers evenly where they fall short of it.
  for (const s of spans) {
    const inner = 2 * sp * (s.to - s.from - 1);
    for (const arr of [min, max]) {
      let have = inner;
      for (let i = s.from; i < s.to; i++) have += arr[i];
      const need = arr === min ? s.min : s.max;
      if (need > have) for (let i = s.from; i < s.to; i++) arr[i] += (need - have) / (s.to - s.from);
    }
  }
  for (let i = 0; i < n; i++) max[i] = Math.min(MAX_COLUMN, Math.max(max[i], min[i]));
  // Cells sit twice the spacing apart and from the table's edges.
  const room = maxWidth - 2 * sp * (n + 1);
  const sumMin = min.reduce((a, b) => a + b, 0);
  const sumMax = max.reduce((a, b) => a + b, 0);
  let boxes = max;
  if (sumMax > room) boxes = sumMin >= room ? min : min.map((m, i) => m + ((max[i] - m) * (room - sumMin)) / (sumMax - sumMin));
  return boxes.map((b, i) => b + (i === 0 ? 2 * sp : sp) + (i === n - 1 ? 2 * sp : sp));
}

/** An autofit table's grid with each column widened to its cells' longest word. */
function widenToWords(table: Table, grid: number[], env: BoxEnv): number[] {
  const out = grid.slice();
  for (const row of table.rows) {
    let col = row.gridBefore;
    const sp = row.cellSpacing;
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const from = col;
      col += cell.gridSpan;
      if (cell.gridSpan !== 1 || from >= out.length || cell.vMerge === 'continue' || turnedText(cell.textDirection)) continue;
      const lead = ci === 0 ? 2 * sp : sp;
      const trail = ci === row.cells.length - 1 ? 2 * sp : sp;
      const need = cellContentWidths(cell, env, true).lo + cell.margins.left + cell.margins.right + lead + trail;
      if (need > out[from]) out[from] = need;
    }
  }
  return out;
}

/**
 * Whether a row another tool wrote gives every column a fixed preferred width (one cell per column)
 * that disagrees with the grid. Word sizes such a table's columns from the cells, not from the grid
 * saved with it. A row Word saved (it carries revision ids), or a grid that cells only partly pin
 * down (automatic or spanning cells), holds the columns Word computed.
 */
function gridContradicted(table: Table): boolean {
  return table.rows.some((row) => {
    if (row.stamped || row.gridBefore || row.cells.length !== table.grid.length) return false;
    let differs = false;
    for (let i = 0; i < row.cells.length; i++) {
      const cell = row.cells[i];
      if (cell.gridSpan !== 1 || cell.width.type !== 'dxa' || cell.width.value <= 0) return false;
      if (Math.abs(cell.width.value - table.grid[i]) > 1) differs = true;
    }
    return differs;
  });
}

export function measureTable(table: Table, avail: number, env: BoxEnv, nested = false, starts?: Map<TableCell, CellCursor>): TableLayout {
  // Word snaps lines in table cells to the document grid only when the document asks it to.
  const cellEnv = env.gridPitch && !env.settings.adjustLineHeightInTable ? { ...env, gridPitch: undefined } : env;
  let grid = gridWidths(table, avail);
  // A nested table stays inside its cell's text area: Word shrinks an autofit one that is too wide
  // and moves one that would start left of the text (a negative indent) back in.
  const total = grid.reduce((a, b) => a + b, 0);
  if (nested && table.props.layout !== 'fixed' && total > avail) grid = grid.map((g) => (g * avail) / total);
  let x = table.props.indent;
  const aligned = table.props.align === 'center' || table.props.align === 'right';
  const spacing = table.rows[0]?.cellSpacing ?? 0;
  if (!aligned && env.compatibilityMode < 15 && !(spacing > 0 && env.compatibilityMode < 14)) {
    // Before Word 2013 the table indent positions the first cell's text area, not its border, so
    // the table moves out by the cell margin and (Word 2010) the cell spacing. Word 2007 leaves
    // tables with spaced cells at the indent.
    const firstCell = table.rows[0]?.cells[0];
    x -= (firstCell ? firstCell.margins.left : table.props.cellMargins.left) + 2 * spacing;
  }
  if (nested) x = Math.max(0, x);
  // An autofit table may reach the right margin; in Word 2010 mode and before, its last cell's text
  // does, so that cell's margin and spacing go past the margin.
  const lastRow = table.rows[table.rows.length - 1];
  const lastCell = lastRow?.cells[lastRow.cells.length - 1];
  const past = env.compatibilityMode < 15 && lastCell ? lastCell.margins.right + 2 * lastRow.cellSpacing : 0;
  const maxWidth = (aligned ? avail : avail - x) + past;
  if ((env.settings.generated || gridContradicted(table)) && !nested && table.props.layout === 'autofit' && table.props.width.type === 'auto' && !table.props.floating && grid.length > 0) {
    grid = autofitGrid(table, maxWidth, env);
  } else if (!nested && table.props.layout === 'autofit') {
    // Word widens a column to hold its longest word while the table still fits; past that it
    // breaks the word.
    const wide = widenToWords(table, grid, cellEnv);
    if (wide.reduce((a, b) => a + b, 0) <= maxWidth + 0.01) grid = wide;
  }
  const colX: number[] = [0];
  for (const g of grid) colX.push(colX[colX.length - 1] + g);
  const width = colX[colX.length - 1];
  if (table.props.align === 'center') x = (avail - width) / 2;
  else if (table.props.align === 'right') x = avail - width;
  // A right-to-left table starts at the right margin, its indent taken from there.
  const rtl = table.props.bidiVisual;
  if (rtl && table.props.align !== 'center') x = avail - width - x;

  const rows: RowLayout[] = [];
  for (const row of table.rows) {
    const cells: CellLayout[] = [];
    const contents: Block[][] = [];
    let col = row.gridBefore;
    // Cell spacing insets every cell of the row on all sides (Word's HTML-style cellspacing).
    // Word's spacing value is half the gap: adjacent cells sit twice it apart, and the outermost
    // cells twice it in from the table's edges.
    const sp = row.cellSpacing;
    const ri = rows.length;
    const spaceTop = ri === 0 ? 2 * sp : sp;
    const spaceBottom = ri === table.rows.length - 1 ? 2 * sp : sp;
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const c0 = Math.min(col, grid.length);
      const c1 = Math.min(col + cell.gridSpan, grid.length);
      const lead = ci === 0 ? 2 * sp : sp;
      const trail = ci === row.cells.length - 1 ? 2 * sp : sp;
      const cx = colX[c0] + lead;
      const cw = Math.max(0, colX[c1] - colX[c0] - lead - trail);
      const inner = Math.max(1, cw - cell.margins.left - cell.margins.right - (sp > 0 ? sideBorders(cell) : 0));
      // A continued merge has no content of its own. hideMark hides the cell's empty paragraph
      // marks: a cell of nothing but empty paragraphs takes no room (a row of them collapses to its
      // margins), and an empty last paragraph after text adds nothing below it.
      const isEmpty = (b: Block) => b.type === 'paragraph' && b.content.length === 0 && !b.numbering;
      let content = cell.content;
      if (cell.hideMark) {
        if (content.every(isEmpty)) content = [];
        else if (isEmpty(content[content.length - 1])) content = content.slice(0, -1);
      }
      // The empty paragraph that must close a cell after a nested table takes no room in Word: the
      // cell ends at the nested table's bottom.
      if (content.length > 1 && isEmpty(content[content.length - 1]) && content[content.length - 2].type === 'table') content = content.slice(0, -1);
      const empty = cell.vMerge === 'continue' || content.length === 0;
      // Turned text runs along the row: Word grows the row to fit its longest line rather than
      // wrapping it, unless the row's height is exact.
      const dir = empty ? undefined : turnedText(cell.textDirection);
      let length = 0;
      if (dir) {
        const exact = row.heightRule === 'exact' && row.height > 0;
        length = Math.max(1, exact ? row.height - cell.margins.top - cell.margins.bottom : Math.min(MAX_COLUMN, cellContentWidths({ ...cell, content }, cellEnv).hi));
      }
      let box = empty ? { height: 0, items: [] } : layoutBlocks(content, dir ? length : inner, cellEnv, { cell: true, trackCursor: !!starts, start: starts?.get(cell) });
      const floats = dir ? 0 : inCellFloatsBottom(box, cell.margins.top);
      if (floats > box.height) box = { ...box, height: floats };
      cells.push({ cell, col, x: cx, width: cw, box, rowSpan: 1, turned: dir ? { dir, length } : undefined });
      contents.push(content);
      col += cell.gridSpan;
    }
    let h = 0;
    // Word places cell content below the row's top border, so border width adds to the height.
    // Adjacent rows share one border: the bottom border of the cell above counts too, but only by
    // what it is wider than the cell's top margin, which it sits in.
    let borderTop = 0;
    const above = rows[ri - 1]?.cells;
    for (const c of cells) {
      const up = above?.find((o) => o.col === c.col)?.cell.borders.bottom?.width ?? 0;
      borderTop = Math.max(borderTop, c.cell.borders.top?.width ?? 0, up - c.cell.margins.top);
    }
    for (const c of cells) if (c.cell.vMerge !== 'restart' && c.cell.vMerge !== 'continue') h = Math.max(h, (c.turned ? c.turned.length : c.box.height) + c.cell.margins.top + c.cell.margins.bottom);
    h += borderTop + spaceTop + spaceBottom;
    // A spaced cell keeps its bottom border inside it too (the last row's is added below).
    if (sp > 0 && ri < table.rows.length - 1) h += Math.max(0, ...cells.map((c) => c.cell.borders.bottom?.width ?? 0));
    // Rows with only merged cells still need room for their own (empty) content.
    if (cells.every((c) => c.cell.vMerge)) for (const c of cells) if (c.cell.vMerge === 'continue') h = Math.max(h, c.cell.margins.top + c.cell.margins.bottom);
    if (cells.length === 0) h = 0;
    if (row.heightRule === 'exact' && row.height > 0) {
      // An exact height leaves out the row's bottom margin, which Word adds below it.
      let bottom = 0;
      for (const c of cells) bottom = Math.max(bottom, c.cell.margins.bottom);
      h = row.height + bottom;
    }
    // A minimum row height leaves out the row's top border and the cells' top and bottom margins,
    // as content heights do.
    else if (row.height > 0) {
      let margins = 0;
      for (const c of cells) margins = Math.max(margins, c.cell.margins.top + c.cell.margins.bottom);
      h = Math.max(h, row.height + margins + borderTop);
    }
    // Turned lines span the whole row, so centred or right-aligned text moves along it.
    cells.forEach((c, i) => {
      const length = h - borderTop - spaceTop - spaceBottom - c.cell.margins.top - c.cell.margins.bottom;
      if (!c.turned || row.heightRule === 'exact' || length <= c.turned.length + EPS) return;
      c.box = layoutBlocks(contents[i], length, cellEnv, { cell: true });
      c.turned.length = length;
    });
    rows.push({ row, cells, height: h, borderTop, spaceTop, spaceBottom });
  }
  // The table's bottom border sits below the last row's content.
  const last = rows[rows.length - 1];
  if (last) {
    let bb = 0;
    for (const c of last.cells) bb = Math.max(bb, c.cell.borders.bottom?.width ?? 0);
    if (last.row.heightRule !== 'exact') {
      last.height += bb;
      last.borderBottom = bb;
    }
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
  // Laid out from the start, a right-to-left table's cells are then mirrored across it.
  if (rtl) for (const r of rows) for (const c of r.cells) c.x = width - c.x - c.width;
  return { x, width, colX, rows, headerRows, ...(rtl && { rtl }) };
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

/** The border Word draws where two cells' borders meet: the wider, or the first when they tie. */
function heavier(a: Border | null, b: Border | null | undefined): Border | null {
  return b && (!a || b.width > a.width) ? b : a;
}

/** Widths of a cell's left and right borders. */
function sideBorders(cell: TableCell): number {
  return (cell.borders.left?.width ?? 0) + (cell.borders.right?.width ?? 0);
}

function drawEdge(out: DrawOp[], br: Border | null, x1: number, y1: number, x2: number, y2: number): void {
  if (!br || br.width <= 0) return;
  const color = br.color ?? '000000';
  const horiz = y1 === y2;
  // A vertical border straddles its edge. A horizontal one hangs from it: Word starts it at most
  // 0.375pt above the edge and runs it down into the room the row below (or, for the table's
  // bottom, the last row) keeps for it.
  if (horiz) {
    const shift = (br.width - Math.min(br.width, 0.75)) / 2;
    y1 += shift;
    y2 += shift;
  }
  if (br.style === 'double') {
    const w = Math.max(0.25, br.width / 3);
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
      // Spaced cells sit inside their rows, clear of the spacing above the first and below the last.
      const top = p.y + tl.rows[spanFrom].spaceTop;
      height -= tl.rows[spanFrom].spaceTop + tl.rows[Math.min(spanEnd, tl.rows.length) - 1].spaceBottom;
      if (cell.cell.shading) back.push({ kind: 'rect', x, y: top, w: cell.width, h: height, fill: cell.cell.shading });
      const m = cell.cell.margins;
      // A right-to-left table's cells start on the right: their left margin and border are on that side.
      const side = (o: CellLayout, which: 'left' | 'right'): Border | null => o.cell.borders[!!tl.rtl === (which === 'left') ? 'right' : 'left'];
      // Spaced cells hold their borders inside them, so text starts past the left border.
      const spaced = tl.rows[spanFrom].row.cellSpacing > 0;
      const innerX = x + (tl.rtl ? m.right : m.left) + (spaced ? (side(cell, 'left')?.width ?? 0) : 0);
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
        // Word centres and bottom-aligns content in the space between the row's borders.
        const ruled = tl.rows[spanFrom].borderTop + (tl.rows[Math.min(spanEnd, tl.rows.length) - 1]?.borderBottom ?? 0);
        const free = innerH - ruled - cell.box.height;
        if (free > 0 && cell.cell.vAlign === 'center') dy += free / 2;
        else if (free > 0 && cell.cell.vAlign === 'bottom') dy += free;
      }
      const area = { x: innerX, y: top, w: cell.width - m.left - m.right - (spaced ? sideBorders(cell.cell) : 0), h: height };
      if (cell.turned) {
        // The content box, laid out along the row, turns a quarter into the cell: text running up
        // starts at the cell's bottom with its lines' tops to the left; text running down starts at
        // the top with them to the right. Vertical alignment moves the lines across the cell.
        const free = area.w - cell.box.height;
        const off = free > 0 && cell.cell.vAlign === 'center' ? free / 2 : free > 0 && cell.cell.vAlign === 'bottom' ? free : 0;
        const y0 = top + m.top + tl.rows[spanFrom].borderTop;
        const dir = cell.turned.dir;
        const matrix: [number, number, number, number, number, number] =
          dir === 'up' ? [0, -1, 1, 0, innerX + off, top + height - m.bottom] : [0, 1, -1, 0, dir === 'down' ? innerX + area.w - off : innerX + off + cell.box.height, y0];
        content.push({ kind: 'transform', matrix });
        for (const it of items) {
          content.push(...it.ops);
          if (it.notes) notes.push(...it.notes);
        }
        content.push({ kind: 'restore' });
        from = to;
      }
      for (let k = from; k < to; k++) {
        // Objects laid out in the cell position from it (the innermost cell of nested tables).
        for (const op of translateOps(items[k].ops, innerX, dy)) content.push(op.kind === 'float' && !op.cell ? { ...op, cell: area } : op);
        if (items[k].notes) notes.push(...items[k].notes!);
      }
      const b = cell.cell.borders;
      const lastRow = tl.rows[Math.min(spanEnd, tl.rows.length) - 1];
      const bottomCell = lastRow?.cells.find((o) => o.col === cell.col);
      // Cells that touch share an edge, and Word draws only the heavier of their two borders there.
      const touching = tl.rows[spanFrom].row.cellSpacing <= 0;
      const own = tl.rows[spanFrom].cells;
      const above = touching && spanFrom > 0 ? tl.rows[spanFrom - 1].cells.find((o) => o.col === cell.col) : undefined;
      const below = touching ? tl.rows[spanEnd]?.cells.find((o) => o.col === cell.col) : undefined;
      const before = touching ? own.find((o) => o.col + o.cell.gridSpan === cell.col) : undefined;
      const after = touching ? own.find((o) => o.col === cell.col + cell.cell.gridSpan) : undefined;
      const left = tl.rtl ? after : before;
      const right = tl.rtl ? before : after;
      const bottomB = heavier(bottomCell ? bottomCell.cell.borders.bottom : b.bottom, below?.cell.borders.top);
      const yb = top + height - (lastRow?.borderBottom ?? 0);
      if (spaced) {
        // Each spaced cell draws its whole border inside its own box.
        const hang = (br: Border | null) => (br ? br.width / 2 - (br.width - Math.min(br.width, 0.75)) / 2 : 0);
        const half = (br: Border | null) => (br ? br.width / 2 : 0);
        drawEdge(borders, b.top, x, top + hang(b.top), x + cell.width, top + hang(b.top));
        drawEdge(borders, bottomB, x, top + height - half(bottomB) * 2 + hang(bottomB), x + cell.width, top + height - half(bottomB) * 2 + hang(bottomB));
        const bl = side(cell, 'left');
        const br = side(cell, 'right');
        drawEdge(borders, bl, x + half(bl), top, x + half(bl), top + height);
        drawEdge(borders, br, x + cell.width - half(br), top, x + cell.width - half(br), top + height);
      } else {
        drawEdge(borders, heavier(b.top, above?.cell.borders.bottom), x, top, x + cell.width, top);
        drawEdge(borders, bottomB, x, yb, x + cell.width, yb);
        drawEdge(borders, heavier(side(cell, 'left'), left && side(left, 'right')), x, top, x, top + height);
        drawEdge(borders, heavier(side(cell, 'right'), right && side(right, 'left')), x + cell.width, top, x + cell.width, top + height);
      }
      drawEdge(borders, b.tl2br, x, top, x + cell.width, top + height);
      drawEdge(borders, b.tr2bl, x + cell.width, top, x, top + height);
    }
  }
  return { ops: [...back, ...content, ...borders], notes };
}

/** Tables inside boxes: one item per group of rows tied together by vertical merges. */
function layoutTableInBox(table: Table, width: number, env: BoxEnv, y0: number, nested: boolean, block?: number, firstRow = 0): { items: BoxItem[]; bottom: number } {
  const tl = measureTable(table, width, env, nested);
  const items: BoxItem[] = [];
  const merge: MergeProgress = new Map();
  let y = y0;
  let r = firstRow;
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
    items.push({ cursor: block === undefined ? undefined : { block, row: r }, y, height: gy - y, ops: seg.ops, notes: seg.notes.length ? seg.notes : undefined });
    y = gy;
    r = end;
  }
  return { items, bottom: y };
}
