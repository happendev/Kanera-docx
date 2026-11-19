import { child, children, type XmlElement } from '../util/xml.js';
import { parseParaDelta, parseRunDelta, type ParaDelta, type PropContext, type RunDelta } from './props.js';
import { num } from './values.js';

export interface Level {
  ilvl: number;
  start: number;
  numFmt: string;
  lvlText: string;
  jc: 'left' | 'center' | 'right';
  suffix: 'tab' | 'space' | 'nothing';
  isLgl: boolean;
  restart?: number;
  pStyle?: string;
  pPr: ParaDelta;
  rPr: RunDelta;
  picBulletId?: string;
}

interface AbstractNum {
  id: string;
  levels: Map<number, Level>;
  numStyleLink?: string;
  styleLink?: string;
}

interface NumInstance {
  id: string;
  abstractId: string;
  overrides: Map<number, { start?: number; level?: Level }>;
}

function parseLevel(el: XmlElement, ctx: PropContext): Level {
  const val = (name: string) => child(el, name)?.attrs['w:val'];
  const jc = val('w:lvlJc');
  const suff = val('w:suff');
  return {
    ilvl: num(el.attrs['w:ilvl']) ?? 0,
    start: num(val('w:start')) ?? 0,
    numFmt: val('w:numFmt') ?? 'decimal',
    lvlText: val('w:lvlText') ?? '',
    jc: jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : 'left',
    suffix: suff === 'space' ? 'space' : suff === 'nothing' ? 'nothing' : 'tab',
    isLgl: !!child(el, 'w:isLgl') && child(el, 'w:isLgl')?.attrs['w:val'] !== '0',
    restart: num(val('w:lvlRestart')),
    pStyle: val('w:pStyle'),
    pPr: parseParaDelta(child(el, 'w:pPr'), ctx),
    rPr: parseRunDelta(child(el, 'w:rPr'), ctx),
    picBulletId: val('w:lvlPicBulletId'),
  };
}

export class Numbering {
  private abstracts = new Map<string, AbstractNum>();
  private nums = new Map<string, NumInstance>();
  /** Picture bullets: id -> relationship id of the image. */
  readonly picBullets = new Map<string, string>();
  // Counter state per abstract numbering definition (lists sharing a definition continue).
  private counters = new Map<string, Array<number | undefined>>();
  private startedNums = new Set<string>();

  constructor(
    root: XmlElement | undefined,
    ctx: PropContext,
    private styleNumId: (styleId: string) => string | undefined,
  ) {
    if (!root) return;
    for (const pb of children(root, 'w:numPicBullet')) {
      const id = pb.attrs['w:numPicBulletId'];
      const imagedata = findDeep(pb, 'v:imagedata');
      const blip = findDeep(pb, 'a:blip');
      const rid = imagedata?.attrs['r:id'] ?? blip?.attrs['r:embed'];
      if (id !== undefined && rid) this.picBullets.set(id, rid);
    }
    for (const a of children(root, 'w:abstractNum')) {
      const id = a.attrs['w:abstractNumId'] ?? '';
      const levels = new Map<number, Level>();
      for (const l of children(a, 'w:lvl')) {
        const lvl = parseLevel(l, ctx);
        levels.set(lvl.ilvl, lvl);
      }
      this.abstracts.set(id, {
        id,
        levels,
        numStyleLink: child(a, 'w:numStyleLink')?.attrs['w:val'],
        styleLink: child(a, 'w:styleLink')?.attrs['w:val'],
      });
    }
    for (const n of children(root, 'w:num')) {
      const id = n.attrs['w:numId'] ?? '';
      const overrides = new Map<number, { start?: number; level?: Level }>();
      for (const o of children(n, 'w:lvlOverride')) {
        const ilvl = num(o.attrs['w:ilvl']) ?? 0;
        const so = child(o, 'w:startOverride');
        const lvl = child(o, 'w:lvl');
        overrides.set(ilvl, {
          start: so ? num(so.attrs['w:val']) : undefined,
          level: lvl ? parseLevel(lvl, ctx) : undefined,
        });
      }
      this.nums.set(id, { id, abstractId: child(n, 'w:abstractNumId')?.attrs['w:val'] ?? '', overrides });
    }
  }

  /** The abstract definition for a num, following numStyleLink indirection through numbering styles. */
  private abstractFor(numId: string, depth = 0): AbstractNum | undefined {
    const inst = this.nums.get(numId);
    if (!inst) return undefined;
    const abs = this.abstracts.get(inst.abstractId);
    if (abs?.numStyleLink && depth < 4) {
      const linked = this.styleNumId(abs.numStyleLink);
      if (linked && linked !== numId) return this.abstractFor(linked, depth + 1) ?? abs;
    }
    return abs;
  }

  level(numId: string, ilvl: number): Level | undefined {
    const inst = this.nums.get(numId);
    if (!inst) return undefined;
    const over = inst.overrides.get(ilvl)?.level;
    if (over) return over;
    return this.abstractFor(numId)?.levels.get(ilvl);
  }

  /** Is numbering actually applied (numId 0 or an unknown id removes numbering)? */
  exists(numId: string | undefined): boolean {
    return numId !== undefined && numId !== '0' && this.nums.has(numId);
  }

  /** Advance the counters for a numbered paragraph and return its label text. */
  next(numId: string, ilvl: number): string | null {
    const inst = this.nums.get(numId);
    const abs = this.abstractFor(numId);
    if (!inst || !abs) return null;
    const key = abs.id;
    let counters = this.counters.get(key);
    if (!counters) this.counters.set(key, (counters = new Array(9).fill(undefined)));
    if (!this.startedNums.has(numId)) {
      // A num with start overrides restarts those levels the first time it is used.
      this.startedNums.add(numId);
      for (const [l, o] of inst.overrides) if (o.start !== undefined) counters[l] = undefined;
    }
    const lvl = this.level(numId, ilvl);
    if (!lvl) return null;
    const startFor = (l: number) => inst.overrides.get(l)?.start ?? this.level(numId, l)?.start ?? 0;
    counters[ilvl] = counters[ilvl] === undefined ? startFor(ilvl) : (counters[ilvl] as number) + 1;
    // Deeper levels restart when a shallower level advances (unless lvlRestart says otherwise).
    for (let d = ilvl + 1; d < 9; d++) {
      const restart = this.level(numId, d)?.restart;
      if (restart === 0) continue;
      if (restart !== undefined && ilvl >= restart) continue;
      counters[d] = undefined;
    }
    if (lvl.numFmt === 'none') return lvl.lvlText.replace(/%\d/g, '');
    if (lvl.numFmt === 'bullet') return lvl.lvlText;
    return lvl.lvlText.replace(/%([1-9])/g, (_, d: string) => {
      const l = parseInt(d, 10) - 1;
      const value = counters![l] ?? startFor(l);
      const fmt = lvl.isLgl ? 'decimal' : (this.level(numId, l)?.numFmt ?? 'decimal');
      return formatNumber(value, fmt);
    });
  }
}

function findDeep(el: XmlElement, name: string): XmlElement | undefined {
  for (const c of el.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) return c;
    const d = findDeep(c, name);
    if (d) return d;
  }
  return undefined;
}

const ROMAN: Array<[number, string]> = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
];

function roman(n: number): string {
  if (n <= 0) return String(n);
  let s = '';
  for (const [v, r] of ROMAN) {
    while (n >= v) {
      s += r;
      n -= v;
    }
  }
  return s;
}

function letters(n: number): string {
  if (n <= 0) return String(n);
  // Word repeats the letter: a..z, aa..zz, aaa..
  const ch = String.fromCharCode(97 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

function cardinal(n: number): string {
  if (n < 20) return ONES[n] ?? String(n);
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? '-' + ONES[n % 10] : '');
  if (n < 1000) return ONES[Math.floor(n / 100)] + ' hundred' + (n % 100 ? ' ' + cardinal(n % 100) : '');
  return String(n);
}

/**
 * East Asian counting numerals below ten thousand: 二十一, 一百零五. `zero` fills a skipped place
 * (empty in Japanese); `legal` keeps the leading one before ten (壹拾).
 */
function counting(n: number, digits: string, units: string, zero: string, legal = false): string {
  if (n <= 0 || n >= 10000) return String(n);
  const places = [Math.floor(n / 1000), Math.floor(n / 100) % 10, Math.floor(n / 10) % 10, n % 10];
  const unit = [units[2], units[1], units[0], ''];
  let s = '';
  let gap = false;
  for (let i = 0; i < 4; i++) {
    const d = places[i];
    if (d === 0) {
      if (s) gap = true;
      continue;
    }
    if (gap) s += zero;
    gap = false;
    // 十 alone stands for ten at the start of a number (十一, not 一十一); Japanese also drops
    // the one before 百 and 千.
    const bare = d === 1 && i < 3 && !legal && (zero === '' || (i === 2 && !s));
    s += bare ? unit[i] : digits[d] + unit[i];
  }
  return s;
}

function ordinalSuffix(n: number): string {
  const m100 = n % 100;
  if (m100 >= 11 && m100 <= 13) return 'th';
  return ['th', 'st', 'nd', 'rd'][n % 10] ?? 'th';
}

export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'decimal':
      return String(n);
    case 'decimalZero':
      return n < 10 ? `0${n}` : String(n);
    case 'upperRoman':
      return roman(n).toUpperCase();
    case 'lowerRoman':
      return roman(n);
    case 'upperLetter':
      return letters(n).toUpperCase();
    case 'lowerLetter':
      return letters(n);
    case 'ordinal':
      return `${n}${ordinalSuffix(n)}`;
    case 'cardinalText': {
      const s = cardinal(n);
      return s.charAt(0).toUpperCase() + s.slice(1);
    }
    case 'ordinalText': {
      const s = cardinal(n);
      const o = s.endsWith('one')
        ? s.slice(0, -3) + 'first'
        : s.endsWith('two')
          ? s.slice(0, -3) + 'second'
          : s.endsWith('three')
            ? s.slice(0, -5) + 'third'
            : s.endsWith('ve')
              ? s.slice(0, -2) + 'fth'
              : s.endsWith('t')
                ? s + 'h'
                : s.endsWith('e')
                  ? s.slice(0, -1) + 'th'
                  : s.endsWith('y')
                    ? s.slice(0, -1) + 'ieth'
                    : s + 'th';
      return o.charAt(0).toUpperCase() + o.slice(1);
    }
    case 'decimalEnclosedCircle':
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n);
    case 'decimalEnclosedParen':
      return `(${n})`;
    case 'decimalEnclosedFullstop':
      return `${n}.`;
    case 'chicago':
      return ['*', '†', '‡', '§'][(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1);
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return String(n).replace(/\d/g, (d) => String.fromCodePoint(0xff10 + Number(d)));
    case 'ideographDigital':
    case 'japaneseDigitalTenThousand':
      return String(n).replace(/\d/g, (d) => '〇一二三四五六七八九'[Number(d)]);
    case 'ideographEnclosedCircle':
      return n >= 1 && n <= 10 ? String.fromCodePoint(0x3220 + n - 1) : counting(n, '〇一二三四五六七八九', '十百千', '零');
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'taiwaneseCounting':
    case 'taiwaneseCountingThousand':
      return counting(n, '〇一二三四五六七八九', '十百千', '零');
    case 'japaneseCounting':
      return counting(n, '〇一二三四五六七八九', '十百千', '');
    case 'ideographLegalTraditional':
      return counting(n, '零壹貳參肆伍陸柒捌玖', '拾佰仟', '零', true);
    case 'chineseLegalSimplified':
      return counting(n, '零壹贰叁肆伍陆柒捌玖', '拾佰仟', '零', true);
    case 'japaneseLegal':
      return counting(n, '〇壱弐参四伍六七八九', '拾百阡', '', true);
    case 'ideographTraditional':
      return n >= 1 ? '甲乙丙丁戊己庚辛壬癸'[(n - 1) % 10] : String(n);
    case 'ideographZodiac':
      return n >= 1 ? '子丑寅卯辰巳午未申酉戌亥'[(n - 1) % 12] : String(n);
    case 'none':
      return '';
    default:
      return String(n);
  }
}
