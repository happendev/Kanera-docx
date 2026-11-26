import { attr, child, children, type XmlElement } from '../util/xml.js';
import type { Border, CellBorders, CellMargins, Color, ParaBorders, ParaProps, RunProps, TabStop, TableWidth, FrameProps } from './model.js';
import { themeColorValue, themeFont, type Theme } from './theme.js';
import { applyWordTintShade, blend, HIGHLIGHT_COLORS, hexColor, measure, num, onOff, parseBorder, twip, twipAttr } from './values.js';

export interface PropContext {
  theme: Theme;
  colorMapping?: Record<string, string>;
}

export type RunDelta = Partial<RunProps>;

export interface TabOp {
  pos: number;
  type: TabStop['type'] | 'clear';
  leader: TabStop['leader'];
}

export type ParaDelta = Partial<Omit<ParaProps, 'tabs' | 'borders' | 'spaceBeforeLines' | 'spaceAfterLines'>> & {
  tabOps?: TabOp[];
  borders?: Partial<ParaBorders>;
  numId?: string;
  ilvl?: number;
  /** Indents in characters (w:leftChars and the like), or null where points replace them. */
  indLeftChars?: number | null;
  indFirstLineChars?: number | null;
  /** Spacing in lines (w:beforeLines, w:afterLines), or null where points replace it. */
  spaceBeforeLines?: number | null;
  spaceAfterLines?: number | null;
};

export const TOGGLES = ['bold', 'italic', 'boldCs', 'italicCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'hidden', 'outline', 'emboss', 'imprint', 'shadow'] as const;

const TAB_TYPES: readonly TabOp['type'][] = ['clear', 'left', 'center', 'right', 'decimal', 'bar', 'num'];
const TAB_LEADERS: readonly TabStop['leader'][] = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'];
const PARA_BORDER_SIDES: readonly (keyof ParaBorders)[] = ['top', 'bottom', 'left', 'right', 'between', 'bar'];
const TABLE_BORDER_SIDES: readonly (keyof TableBordersDelta)[] = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];

export function themeColorFn(ctx: PropContext): (name: string) => string | undefined {
  return (name) => themeColorValue(ctx.theme, name, ctx.colorMapping);
}

/** Resolve a w:shd element to a single background colour (patterns blended like Word renders them). */
export function parseShading(el: XmlElement | undefined, ctx: PropContext): Color | undefined {
  if (!el) return undefined;
  const val = el.attrs['w:val'] ?? 'clear';
  const tc = themeColorFn(ctx);
  let fill = hexColor(el.attrs['w:fill']);
  if ((fill === undefined || fill === null) && el.attrs['w:themeFill']) {
    const base = tc(el.attrs['w:themeFill']);
    if (base) fill = applyWordTintShade(base, el.attrs['w:themeFillTint'], el.attrs['w:themeFillShade']);
  }
  let color = hexColor(el.attrs['w:color']);
  if ((color === undefined || color === null) && el.attrs['w:themeColor']) {
    const base = tc(el.attrs['w:themeColor']);
    if (base) color = applyWordTintShade(base, el.attrs['w:themeTint'], el.attrs['w:themeShade']);
  }
  if (val === 'nil') return null;
  if (val === 'clear') return fill ?? null;
  if (val === 'solid') return color ?? '000000';
  const m = /^pct(\d+)$/.exec(val);
  if (m) {
    const pct = parseInt(m[1], 10) / 100;
    return blend(color ?? '000000', fill ?? 'FFFFFF', pct);
  }
  // Hatch patterns: approximate with a 25% blend.
  if (fill || color) return blend(color ?? '000000', fill ?? 'FFFFFF', 0.25);
  return null;
}

export function parseRunDelta(rPr: XmlElement | undefined, ctx: PropContext): RunDelta {
  const d: RunDelta = {};
  if (!rPr) return d;
  for (const el of children(rPr)) {
    switch (el.name) {
      case 'w:rStyle':
        d.styleId = el.attrs['w:val'];
        break;
      case 'w:rFonts': {
        const a = el.attrs;
        const pick = (themeAttr: string, nameAttr: string) => (a[themeAttr] ? themeFont(ctx.theme, a[themeAttr]) : undefined) ?? a[nameAttr];
        const ascii = pick('w:asciiTheme', 'w:ascii');
        const hAnsi = pick('w:hAnsiTheme', 'w:hAnsi');
        const ea = pick('w:eastAsiaTheme', 'w:eastAsia');
        const cs = pick('w:cstheme', 'w:cs');
        if (ascii) d.fontAscii = ascii;
        if (hAnsi) d.fontHAnsi = hAnsi;
        if (ea) d.fontEastAsia = ea;
        if (cs) d.fontCs = cs;
        if (a['w:hint']) d.hint = a['w:hint'];
        break;
      }
      case 'w:lang':
        if (el.attrs['w:eastAsia']) d.eastAsiaLang = el.attrs['w:eastAsia'];
        break;
      case 'w:b':
        d.bold = onOff(el);
        break;
      case 'w:bCs':
        d.boldCs = onOff(el);
        break;
      case 'w:i':
        d.italic = onOff(el);
        break;
      case 'w:iCs':
        d.italicCs = onOff(el);
        break;
      case 'w:caps':
        d.caps = onOff(el);
        break;
      case 'w:smallCaps':
        d.smallCaps = onOff(el);
        break;
      case 'w:strike':
        d.strike = onOff(el);
        break;
      case 'w:dstrike':
        d.dstrike = onOff(el);
        break;
      case 'w:vanish':
        d.hidden = onOff(el);
        break;
      case 'w:outline':
        d.outline = onOff(el);
        break;
      case 'w:emboss':
        d.emboss = onOff(el);
        break;
      case 'w:imprint':
        d.imprint = onOff(el);
        break;
      case 'w:shadow':
        d.shadow = onOff(el);
        break;
      case 'w:color': {
        let c = hexColor(el.attrs['w:val']);
        const tc = el.attrs['w:themeColor'];
        if (tc && (c === undefined || el.attrs['w:val'] === undefined)) {
          const base = themeColorFn(ctx)(tc);
          if (base) c = applyWordTintShade(base, el.attrs['w:themeTint'], el.attrs['w:themeShade']);
        }
        if (c !== undefined) d.color = c;
        break;
      }
      case 'w:spacing': {
        const v = twipAttr(el, 'w:val');
        if (v !== undefined) d.spacing = v;
        break;
      }
      case 'w:snapToGrid':
        d.snapToGrid = onOff(el);
        break;
      case 'w:w': {
        const v = num(el.attrs['w:val']?.replace('%', ''));
        if (v !== undefined) d.scale = v;
        break;
      }
      case 'w:kern': {
        const v = num(el.attrs['w:val']);
        if (v !== undefined) d.kern = v / 2;
        break;
      }
      case 'w:position': {
        const v = measure(el.attrs['w:val'], (n) => n / 2);
        if (v !== undefined) d.position = v;
        break;
      }
      case 'w:sz': {
        const v = measure(el.attrs['w:val'], (n) => n / 2);
        if (v !== undefined) d.size = v;
        break;
      }
      case 'w:szCs': {
        const v = measure(el.attrs['w:val'], (n) => n / 2);
        if (v !== undefined) d.sizeCs = v;
        break;
      }
      case 'w:highlight': {
        const v = el.attrs['w:val'];
        d.highlight = v && v !== 'none' ? (Object.hasOwn(HIGHLIGHT_COLORS, v) ? HIGHLIGHT_COLORS[v] : hexColor(v) ?? null) : null;
        break;
      }
      case 'w:u': {
        // A w:u without w:val only sets the underline colour; Word draws no underline for it.
        const v = el.attrs['w:val'];
        if (v !== undefined) d.underline = v === 'none' ? null : v;
        const c = hexColor(el.attrs['w:color']);
        if (c !== undefined) d.underlineColor = c;
        else if (el.attrs['w:themeColor']) {
          const base = themeColorFn(ctx)(el.attrs['w:themeColor']);
          if (base) d.underlineColor = applyWordTintShade(base, el.attrs['w:themeTint'], el.attrs['w:themeShade']);
        }
        break;
      }
      case 'w:shd': {
        const s = parseShading(el, ctx);
        if (s !== undefined) d.shading = s;
        break;
      }
      case 'w:bdr': {
        const b = parseBorder(el, themeColorFn(ctx));
        if (b !== undefined) d.border = b;
        break;
      }
      case 'w:vertAlign': {
        const v = el.attrs['w:val'];
        d.vertAlign = v === 'superscript' || v === 'subscript' ? v : 'baseline';
        break;
      }
      case 'w:rtl':
        d.rtl = onOff(el);
        break;
      case 'w:cs':
        d.cs = onOff(el);
        break;
    }
  }
  return d;
}

const JC: Record<string, ParaProps['align']> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  both: 'both',
  justify: 'both',
  distribute: 'distribute',
  thaiDistribute: 'distribute',
  mediumKashida: 'both',
  highKashida: 'both',
  lowKashida: 'both',
  numTab: 'left',
};

export function parseParaDelta(pPr: XmlElement | undefined, ctx: PropContext): ParaDelta {
  const d: ParaDelta = {};
  if (!pPr) return d;
  const tc = themeColorFn(ctx);
  for (const el of children(pPr)) {
    switch (el.name) {
      case 'w:pStyle':
        d.styleId = el.attrs['w:val'];
        break;
      case 'w:keepNext':
        d.keepNext = onOff(el);
        break;
      case 'w:keepLines':
        d.keepLines = onOff(el);
        break;
      case 'w:pageBreakBefore':
        d.pageBreakBefore = onOff(el);
        break;
      case 'w:widowControl':
        d.widowControl = onOff(el);
        break;
      case 'w:contextualSpacing':
        d.contextualSpacing = onOff(el);
        break;
      case 'w:bidi':
        d.bidi = onOff(el);
        break;
      case 'w:snapToGrid':
        d.snapToGrid = onOff(el);
        break;
      case 'w:autoSpaceDE':
        d.autoSpaceDE = onOff(el);
        break;
      case 'w:autoSpaceDN':
        d.autoSpaceDN = onOff(el);
        break;
      case 'w:overflowPunct':
        d.overflowPunct = onOff(el);
        break;
      case 'w:suppressAutoHyphens':
        d.suppressAutoHyphens = onOff(el);
        break;
      case 'w:jc': {
        const v = JC[el.attrs['w:val'] ?? 'left'];
        if (v) d.align = v;
        break;
      }
      case 'w:textAlignment':
        d.textAlignment = el.attrs['w:val'];
        break;
      case 'w:outlineLvl': {
        const v = num(el.attrs['w:val']);
        if (v !== undefined) d.outlineLevel = Number.isInteger(v) && v >= 0 && v < 9 ? v : null;
        break;
      }
      case 'w:numPr': {
        const id = child(el, 'w:numId')?.attrs['w:val'];
        const lvl = num(child(el, 'w:ilvl')?.attrs['w:val']);
        if (id !== undefined) d.numId = id;
        if (lvl !== undefined) d.ilvl = lvl;
        break;
      }
      case 'w:spacing': {
        const a = el.attrs;
        const before = twipAttr(el, 'w:before');
        const after = twipAttr(el, 'w:after');
        if (before !== undefined) [d.spaceBefore, d.spaceBeforeLines] = [before, null];
        if (after !== undefined) [d.spaceAfter, d.spaceAfterLines] = [after, null];
        // beforeLines/afterLines are hundredths of a line; the parser sizes the line once it knows
        // the section's grid. A zero, which WPS writes beside every point value, leaves the points
        // in force.
        const bl = num(a['w:beforeLines']);
        const al = num(a['w:afterLines']);
        if (bl) d.spaceBeforeLines = bl / 100;
        if (al) d.spaceAfterLines = al / 100;
        if (a['w:beforeAutospacing'] !== undefined) d.spaceBeforeAuto = onOff({ name: '', attrs: { 'w:val': a['w:beforeAutospacing'] }, children: [] });
        if (a['w:afterAutospacing'] !== undefined) d.spaceAfterAuto = onOff({ name: '', attrs: { 'w:val': a['w:afterAutospacing'] }, children: [] });
        if (a['w:line'] !== undefined) {
          const rule = a['w:lineRule'];
          const raw = num(a['w:line']) ?? 240;
          d.lineRule = rule === 'exact' || rule === 'atLeast' ? rule : 'auto';
          d.line = d.lineRule === 'auto' ? raw / 240 : twip(raw);
          if (d.lineRule !== 'auto' && raw < 0) {
            d.lineRule = 'exact';
            d.line = twip(-raw);
          }
        }
        break;
      }
      case 'w:ind': {
        const a = el.attrs;
        const left = twipAttr(el, 'w:left') ?? twipAttr(el, 'w:start');
        const right = twipAttr(el, 'w:right') ?? twipAttr(el, 'w:end');
        const first = twipAttr(el, 'w:firstLine');
        const hanging = twipAttr(el, 'w:hanging');
        if (left !== undefined) [d.indLeft, d.indLeftChars] = [left, null];
        if (right !== undefined) d.indRight = right;
        if (first !== undefined) [d.indFirstLine, d.indFirstLineChars] = [first, null];
        if (hanging !== undefined) [d.indFirstLine, d.indFirstLineChars] = [-hanging, null];
        // Character-unit indents (hundredths of a character) win over the points beside them; a
        // zero leaves the points in force. The character size is known once the paragraph is.
        const lc = num(a['w:leftChars'] ?? a['w:startChars']);
        if (lc) d.indLeftChars = lc / 100;
        const hc = num(a['w:hangingChars']);
        if (hc) d.indFirstLineChars = -hc / 100;
        const fc = num(a['w:firstLineChars']);
        if (fc) d.indFirstLineChars = fc / 100;
        break;
      }
      case 'w:tabs': {
        d.tabOps = children(el, 'w:tab').map((t) => {
          const val = t.attrs['w:val'];
          const type = val === 'start' ? 'left' : val === 'end' ? 'right' : TAB_TYPES.find((type) => type === val) ?? 'left';
          const leader = TAB_LEADERS.find((leader) => leader === t.attrs['w:leader']) ?? 'none';
          return { pos: twipAttr(t, 'w:pos') ?? 0, type, leader };
        });
        break;
      }
      case 'w:pBdr': {
        const b: Partial<ParaBorders> = {};
        for (const side of children(el)) {
          const key = borderSide(side, PARA_BORDER_SIDES);
          if (key === undefined) continue;
          const border = parseBorder(side, tc);
          if (border !== undefined) b[key] = border;
        }
        d.borders = b;
        break;
      }
      case 'w:shd': {
        const s = parseShading(el, ctx);
        if (s !== undefined) d.shading = s;
        break;
      }
      case 'w:framePr':
        d.frame = parseFrame(el);
        break;
    }
  }
  return d;
}

function parseFrame(el: XmlElement): FrameProps {
  const a = el.attrs;
  return {
    width: twipAttr(el, 'w:w'),
    height: twipAttr(el, 'w:h'),
    hRule: a['w:hRule'],
    x: twipAttr(el, 'w:x'),
    y: twipAttr(el, 'w:y'),
    xAlign: a['w:xAlign'],
    yAlign: a['w:yAlign'],
    hAnchor: a['w:hAnchor'],
    vAnchor: a['w:vAnchor'],
    wrap: a['w:wrap'],
    hSpace: twipAttr(el, 'w:hSpace') ?? 0,
    vSpace: twipAttr(el, 'w:vSpace') ?? 0,
    dropCap: a['w:dropCap'],
    lines: num(a['w:lines']),
  };
}

// ---------------------------------------------------------------------------------------------
// Merging

function copyDefined<T extends object>(target: T, source: Partial<T>): T {
  for (const key in source) {
    if (!Object.hasOwn(source, key)) continue;
    const value = source[key];
    if (value !== undefined) target[key] = value;
  }
  return target;
}

export function mergeRun(base: RunDelta, over: RunDelta): RunDelta {
  return copyDefined({ ...base }, over);
}

export function mergePara(base: ParaDelta, over: ParaDelta): ParaDelta {
  const out = copyDefined({ ...base }, over);
  if (over.tabOps !== undefined) out.tabOps = [...(base.tabOps ?? []), ...over.tabOps];
  if (over.borders !== undefined) out.borders = { ...base.borders, ...over.borders };
  return out;
}

/** Apply toggle-property semantics: style levels XOR, direct formatting is absolute. */
export function combineStyleRuns(levels: RunDelta[]): RunDelta {
  let out: RunDelta = {};
  for (const level of levels) {
    const next = mergeRun(out, level);
    for (const t of TOGGLES) {
      if (level[t] !== undefined && out[t] !== undefined) next[t] = out[t] !== level[t];
    }
    out = next;
  }
  return out;
}

export const DEFAULT_RUN: RunProps = {
  fontAscii: 'Times New Roman',
  fontHAnsi: 'Times New Roman',
  fontEastAsia: 'Times New Roman',
  fontCs: 'Times New Roman',
  size: 10,
  sizeCs: 10,
  bold: false,
  italic: false,
  boldCs: false,
  italicCs: false,
  underline: null,
  underlineColor: null,
  strike: false,
  dstrike: false,
  color: null,
  highlight: null,
  shading: null,
  vertAlign: 'baseline',
  caps: false,
  smallCaps: false,
  spacing: 0,
  scale: 100,
  position: 0,
  hidden: false,
  kern: 0,
  rtl: false,
  cs: false,
  border: null,
  outline: false,
  emboss: false,
  imprint: false,
  shadow: false,
};

export function finalizeRun(d: RunDelta): RunProps {
  const r = copyDefined({ ...DEFAULT_RUN }, d);
  // A font set only for ASCII applies to hAnsi too when hAnsi is unset (Word fills them together).
  if (d.fontAscii && !d.fontHAnsi) r.fontHAnsi = d.fontAscii;
  return r;
}

export const DEFAULT_PARA: ParaProps = {
  styleId: '',
  align: 'left',
  indLeft: 0,
  indRight: 0,
  indFirstLine: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  spaceBeforeAuto: false,
  spaceAfterAuto: false,
  line: 1,
  lineRule: 'auto',
  contextualSpacing: false,
  keepNext: false,
  keepLines: false,
  pageBreakBefore: false,
  // Word applies widow/orphan control unless a style or paragraph turns it off.
  widowControl: true,
  tabs: [],
  borders: { top: null, bottom: null, left: null, right: null, between: null, bar: null },
  shading: null,
  outlineLevel: null,
  bidi: false,
  snapToGrid: true,
  autoSpaceDE: true,
  autoSpaceDN: true,
  overflowPunct: true,
  suppressAutoHyphens: false,
  textAlignment: 'auto',
  frame: null,
};

/** `charSize` is the font size a character-unit indent counts in: the paragraph mark's. */
export function finalizePara(d: ParaDelta, charSize: number): ParaProps {
  const { tabOps, borders, numId, ilvl, indLeftChars, indFirstLineChars, spaceBeforeLines, spaceAfterLines, ...props } = d;
  const p = copyDefined({ ...DEFAULT_PARA, borders: { ...DEFAULT_PARA.borders, ...borders } }, props);
  // Lines are 12pt until the parser knows the section's grid (resolveLineSpacing).
  if (spaceBeforeLines) [p.spaceBefore, p.spaceBeforeLines] = [spaceBeforeLines * 12, spaceBeforeLines];
  if (spaceAfterLines) [p.spaceAfter, p.spaceAfterLines] = [spaceAfterLines * 12, spaceAfterLines];
  if (indLeftChars) p.indLeft = indLeftChars * charSize;
  if (indFirstLineChars) p.indFirstLine = indFirstLineChars * charSize;
  if (p.bidi) {
    // In a right-to-left paragraph "left" and "right" name the start and end sides: its start
    // is on the right.
    if (p.align === 'left') p.align = 'right';
    else if (p.align === 'right') p.align = 'left';
    [p.indLeft, p.indRight] = [p.indRight, p.indLeft];
  }
  // Tab stops: later definitions replace earlier ones at the same position; "clear" removes.
  const stops = new Map<number, TabStop>();
  for (const op of tabOps ?? []) {
    const key = Math.round(op.pos * 20);
    if (op.type === 'clear') stops.delete(key);
    else stops.set(key, { pos: op.pos, type: op.type, leader: op.leader });
  }
  // A "clear" in direct formatting for a stop defined in a style also removes it, even if the
  // positions were written with slightly different rounding.
  p.tabs = [...stops.values()].sort((a, b) => a.pos - b.pos);
  return p;
}

// ---------------------------------------------------------------------------------------------
// Table properties

export interface TableBordersDelta {
  top?: Border | null;
  bottom?: Border | null;
  left?: Border | null;
  right?: Border | null;
  insideH?: Border | null;
  insideV?: Border | null;
  tl2br?: Border | null;
  tr2bl?: Border | null;
}

export interface TblDelta {
  width?: TableWidth;
  align?: 'left' | 'center' | 'right';
  indent?: number;
  layout?: 'fixed' | 'autofit';
  cellMargins?: Partial<CellMargins>;
  cellSpacing?: number;
  borders?: TableBordersDelta;
  shading?: Color;
  rowBandSize?: number;
  colBandSize?: number;
  styleId?: string;
  look?: TableLook;
  bidiVisual?: boolean;
  floating?: XmlElement;
}

export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

export interface TcDelta {
  width?: TableWidth;
  gridSpan?: number;
  vMerge?: 'restart' | 'continue';
  hMerge?: 'restart' | 'continue';
  borders?: TableBordersDelta;
  shading?: Color;
  margins?: Partial<CellMargins>;
  vAlign?: 'top' | 'center' | 'bottom';
  noWrap?: boolean;
  hideMark?: boolean;
  textDirection?: string;
}

export interface TrDelta {
  height?: number;
  heightRule?: 'auto' | 'atLeast' | 'exact';
  cantSplit?: boolean;
  isHeader?: boolean;
  gridBefore?: number;
  gridAfter?: number;
  jc?: 'left' | 'center' | 'right';
  cellSpacing?: number;
}

export function parseWidth(el: XmlElement | undefined): TableWidth | undefined {
  if (!el) return undefined;
  const type = el.attrs['w:type'];
  const raw = el.attrs['w:w'] ?? '0';
  if (type === 'pct') {
    // Either fiftieths of a percent ("5000") or an explicit percentage ("100%").
    const v = raw.endsWith('%') ? (num(raw.slice(0, -1)) ?? 0) / 100 : (num(raw) ?? 0) / 5000;
    return { type, value: v };
  }
  if (type === 'auto' || type === 'nil') return { type, value: 0 };
  return { type: 'dxa', value: measure(raw) ?? 0 };
}

function borderSide<T extends string>(side: XmlElement, sides: readonly T[]): T | undefined {
  if (!side.name.startsWith('w:')) return undefined;
  const name = side.name.slice(2);
  const key = name === 'start' ? 'left' : name === 'end' ? 'right' : name;
  return sides.find((side) => side === key);
}

export function parseTableBorders(el: XmlElement | undefined, ctx: PropContext): TableBordersDelta | undefined {
  if (!el) return undefined;
  const tc = themeColorFn(ctx);
  const out: TableBordersDelta = {};
  for (const side of children(el)) {
    const key = borderSide(side, TABLE_BORDER_SIDES);
    if (key === undefined) continue;
    const b = parseBorder(side, tc);
    if (b !== undefined) out[key] = b;
  }
  return out;
}

function parseMargins(el: XmlElement | undefined): Partial<CellMargins> | undefined {
  if (!el) return undefined;
  const out: Partial<CellMargins> = {};
  for (const side of children(el)) {
    let key = side.name.slice(2);
    if (key === 'start') key = 'left';
    if (key === 'end') key = 'right';
    const w = parseWidth(side);
    if (w && (key === 'top' || key === 'left' || key === 'bottom' || key === 'right')) out[key] = w.type === 'dxa' ? w.value : 0;
  }
  return out;
}

export function parseTblDelta(tblPr: XmlElement | undefined, ctx: PropContext): TblDelta {
  const d: TblDelta = {};
  if (!tblPr) return d;
  for (const el of children(tblPr)) {
    switch (el.name) {
      case 'w:tblStyle':
        d.styleId = el.attrs['w:val'];
        break;
      case 'w:tblW':
        d.width = parseWidth(el);
        break;
      case 'w:jc': {
        const v = el.attrs['w:val'];
        d.align = v === 'center' ? 'center' : v === 'right' || v === 'end' ? 'right' : 'left';
        break;
      }
      case 'w:tblInd': {
        const w = parseWidth(el);
        if (w) d.indent = w.type === 'dxa' ? w.value : 0;
        break;
      }
      case 'w:tblLayout':
        d.layout = el.attrs['w:type'] === 'fixed' ? 'fixed' : 'autofit';
        break;
      case 'w:tblCellMar':
        d.cellMargins = parseMargins(el);
        break;
      case 'w:tblCellSpacing': {
        const w = parseWidth(el);
        if (w) d.cellSpacing = w.type === 'dxa' ? w.value : 0;
        break;
      }
      case 'w:tblBorders':
        d.borders = parseTableBorders(el, ctx);
        break;
      case 'w:shd': {
        const s = parseShading(el, ctx);
        if (s !== undefined) d.shading = s;
        break;
      }
      case 'w:tblStyleRowBandSize':
        d.rowBandSize = num(el.attrs['w:val']);
        break;
      case 'w:tblStyleColBandSize':
        d.colBandSize = num(el.attrs['w:val']);
        break;
      case 'w:tblLook':
        d.look = parseLook(el);
        break;
      case 'w:bidiVisual':
        d.bidiVisual = onOff(el);
        break;
      case 'w:tblpPr':
        d.floating = el;
        break;
    }
  }
  return d;
}

function parseLook(el: XmlElement): TableLook {
  const a = el.attrs;
  const v = a['w:val'] ? parseInt(a['w:val'], 16) : undefined;
  const flag = (name: string, bit: number, dflt: boolean) => {
    if (a[name] !== undefined) return a[name] === '1' || a[name] === 'true' || a[name] === 'on';
    if (v !== undefined) return (v & bit) !== 0;
    return dflt;
  };
  return {
    firstRow: flag('w:firstRow', 0x20, true),
    lastRow: flag('w:lastRow', 0x40, false),
    firstColumn: flag('w:firstColumn', 0x80, true),
    lastColumn: flag('w:lastColumn', 0x100, false),
    noHBand: flag('w:noHBand', 0x200, false),
    noVBand: flag('w:noVBand', 0x400, true),
  };
}

export function parseTcDelta(tcPr: XmlElement | undefined, ctx: PropContext): TcDelta {
  const d: TcDelta = {};
  if (!tcPr) return d;
  for (const el of children(tcPr)) {
    switch (el.name) {
      case 'w:tcW':
        d.width = parseWidth(el);
        break;
      case 'w:gridSpan':
        d.gridSpan = num(el.attrs['w:val']) ?? 1;
        break;
      case 'w:vMerge':
        d.vMerge = el.attrs['w:val'] === 'restart' ? 'restart' : 'continue';
        break;
      case 'w:hMerge':
        d.hMerge = el.attrs['w:val'] === 'restart' ? 'restart' : 'continue';
        break;
      case 'w:tcBorders':
        d.borders = parseTableBorders(el, ctx);
        break;
      case 'w:shd': {
        const s = parseShading(el, ctx);
        if (s !== undefined) d.shading = s;
        break;
      }
      case 'w:tcMar':
        d.margins = parseMargins(el);
        break;
      case 'w:vAlign': {
        const v = el.attrs['w:val'];
        d.vAlign = v === 'center' ? 'center' : v === 'bottom' ? 'bottom' : 'top';
        break;
      }
      case 'w:noWrap':
        d.noWrap = onOff(el);
        break;
      case 'w:hideMark':
        d.hideMark = onOff(el);
        break;
      case 'w:textDirection':
        d.textDirection = el.attrs['w:val'];
        break;
    }
  }
  return d;
}

export function parseTrDelta(trPr: XmlElement | undefined): TrDelta {
  const d: TrDelta = {};
  if (!trPr) return d;
  for (const el of children(trPr)) {
    switch (el.name) {
      case 'w:trHeight': {
        d.height = twipAttr(el, 'w:val') ?? 0;
        const rule = el.attrs['w:hRule'];
        d.heightRule = rule === 'exact' ? 'exact' : rule === 'auto' ? 'auto' : 'atLeast';
        break;
      }
      case 'w:cantSplit':
        d.cantSplit = onOff(el);
        break;
      case 'w:tblHeader':
        d.isHeader = onOff(el);
        break;
      case 'w:gridBefore':
        d.gridBefore = num(el.attrs['w:val']);
        break;
      case 'w:gridAfter':
        d.gridAfter = num(el.attrs['w:val']);
        break;
      case 'w:jc': {
        const v = el.attrs['w:val'];
        d.jc = v === 'center' ? 'center' : v === 'right' || v === 'end' ? 'right' : 'left';
        break;
      }
      case 'w:tblCellSpacing': {
        const w = parseWidth(el);
        if (w) d.cellSpacing = w.type === 'dxa' ? w.value : 0;
        break;
      }
    }
  }
  return d;
}

export function mergeTbl(a: TblDelta, b: TblDelta): TblDelta {
  const out = copyDefined({ ...a }, b);
  if (b.borders !== undefined) out.borders = { ...a.borders, ...b.borders };
  if (b.cellMargins !== undefined) out.cellMargins = { ...a.cellMargins, ...b.cellMargins };
  return out;
}

export function mergeTc(a: TcDelta, b: TcDelta): TcDelta {
  const out = copyDefined({ ...a }, b);
  if (b.borders !== undefined) out.borders = { ...a.borders, ...b.borders };
  if (b.margins !== undefined) out.margins = { ...a.margins, ...b.margins };
  return out;
}

export function mergeTr(a: TrDelta, b: TrDelta): TrDelta {
  return copyDefined({ ...a }, b);
}

export type { CellBorders };
export { attr };
