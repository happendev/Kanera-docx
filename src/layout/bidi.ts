/**
 * Right-to-left text: a simplified Unicode bidirectional algorithm (strong types, numbers and
 * neutrals; no explicit embeddings) and Arabic joining through the presentation forms most
 * Arabic fonts map.
 */

/** Bidi class of a character: strong left, strong right, Arabic letter, European or Arabic number, neutral. */
export type BidiClass = 'L' | 'R' | 'AL' | 'EN' | 'AN' | 'N';

export function bidiClass(cp: number): BidiClass {
  if (cp >= 0x30 && cp <= 0x39) return 'EN';
  if ((cp >= 0x660 && cp <= 0x669) || (cp >= 0x6f0 && cp <= 0x6f9)) return 'AN';
  if (cp >= 0x590 && cp <= 0x5ff) return 'R';
  if ((cp >= 0x600 && cp <= 0x8ff) || (cp >= 0xfb50 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return isMark(cp) ? 'N' : 'AL';
  if (cp >= 0xfb1d && cp <= 0xfb4f) return 'R';
  if ((cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || cp === 0xaa || cp === 0xb5 || cp === 0xba) return 'L';
  if (cp >= 0xc0 && cp !== 0xd7 && cp !== 0xf7 && cp < 0x2000) return 'L';
  if (cp >= 0x2e80 && !(cp >= 0xfe30 && cp <= 0xfe4f) && !(cp >= 0xff00 && cp <= 0xff20)) return 'L';
  return 'N';
}

/** Arabic combining marks (harakat), which take their base letter's joining. */
function isMark(cp: number): boolean {
  return (cp >= 0x610 && cp <= 0x61a) || (cp >= 0x64b && cp <= 0x65f) || cp === 0x670 || (cp >= 0x6d6 && cp <= 0x6ed);
}

/**
 * Embedding levels for a paragraph's characters: `rtl` marks characters of right-to-left runs
 * (w:rtl), whose neutrals then read right to left. Numbers run left to right inside
 * right-to-left text.
 */
export function bidiLevels(classes: BidiClass[], rtl: boolean[], base: 0 | 1): number[] {
  const n = classes.length;
  const types: BidiClass[] = classes.slice();
  // W2: European numbers after an Arabic letter read as Arabic numbers.
  let lastStrong: BidiClass = base ? 'R' : 'L';
  for (let i = 0; i < n; i++) {
    const t = types[i];
    if (t === 'L' || t === 'R' || t === 'AL') lastStrong = t;
    else if (t === 'EN' && lastStrong === 'AL') types[i] = 'AN';
    else if (t === 'N' && rtl[i]) types[i] = 'R';
  }
  for (let i = 0; i < n; i++) if (types[i] === 'AL') types[i] = 'R';
  // N1/N2: neutrals between strongs of one direction (numbers count as right to left) take it;
  // others take the paragraph direction.
  const dir = (t: BidiClass): 'L' | 'R' | null => (t === 'L' ? 'L' : t === 'R' || t === 'EN' || t === 'AN' ? 'R' : null);
  for (let i = 0; i < n; ) {
    if (types[i] !== 'N') {
      i++;
      continue;
    }
    let j = i;
    while (j < n && types[j] === 'N') j++;
    let before: 'L' | 'R' = base ? 'R' : 'L';
    for (let k = i - 1; k >= 0; k--) {
      const d = dir(types[k]);
      if (d) {
        before = types[k] === 'EN' && !base ? 'L' : d;
        break;
      }
    }
    let after: 'L' | 'R' = base ? 'R' : 'L';
    for (let k = j; k < n; k++) {
      const d = dir(types[k]);
      if (d) {
        after = types[k] === 'EN' && !base ? 'L' : d;
        break;
      }
    }
    const fill = before === after ? before : base ? 'R' : 'L';
    for (let k = i; k < j; k++) types[k] = fill;
    i = j;
  }
  // I1/I2
  return types.map((t) => {
    if (base === 0) return t === 'R' ? 1 : t === 'EN' || t === 'AN' ? 2 : 0;
    return t === 'R' ? 1 : 2;
  });
}

/** Visual order of a line's characters from their levels (rule L2): indices left to right. */
export function visualOrder(levels: number[]): number[] {
  const order = levels.map((_, i) => i);
  let max = 0;
  let minOdd = Infinity;
  for (const l of levels) {
    max = Math.max(max, l);
    if (l % 2) minOdd = Math.min(minOdd, l);
  }
  for (let level = max; level >= minOdd && level > 0; level--) {
    for (let i = 0; i < order.length; ) {
      if (levels[order[i]] < level) {
        i++;
        continue;
      }
      let j = i;
      while (j < order.length && levels[order[j]] >= level) j++;
      order.splice(i, j - i, ...order.slice(i, j).reverse());
      i = j;
    }
  }
  return order;
}

// Presentation forms (isolated, final, initial, medial) for U+0621..U+064A; two forms mean the
// letter joins only to the letter before it.
const FORMS: Record<number, number[]> = {
  0x621: [0xfe80],
  0x622: [0xfe81, 0xfe82],
  0x623: [0xfe83, 0xfe84],
  0x624: [0xfe85, 0xfe86],
  0x625: [0xfe87, 0xfe88],
  0x626: [0xfe89, 0xfe8a, 0xfe8b, 0xfe8c],
  0x627: [0xfe8d, 0xfe8e],
  0x628: [0xfe8f, 0xfe90, 0xfe91, 0xfe92],
  0x629: [0xfe93, 0xfe94],
  0x62a: [0xfe95, 0xfe96, 0xfe97, 0xfe98],
  0x62b: [0xfe99, 0xfe9a, 0xfe9b, 0xfe9c],
  0x62c: [0xfe9d, 0xfe9e, 0xfe9f, 0xfea0],
  0x62d: [0xfea1, 0xfea2, 0xfea3, 0xfea4],
  0x62e: [0xfea5, 0xfea6, 0xfea7, 0xfea8],
  0x62f: [0xfea9, 0xfeaa],
  0x630: [0xfeab, 0xfeac],
  0x631: [0xfead, 0xfeae],
  0x632: [0xfeaf, 0xfeb0],
  0x633: [0xfeb1, 0xfeb2, 0xfeb3, 0xfeb4],
  0x634: [0xfeb5, 0xfeb6, 0xfeb7, 0xfeb8],
  0x635: [0xfeb9, 0xfeba, 0xfebb, 0xfebc],
  0x636: [0xfebd, 0xfebe, 0xfebf, 0xfec0],
  0x637: [0xfec1, 0xfec2, 0xfec3, 0xfec4],
  0x638: [0xfec5, 0xfec6, 0xfec7, 0xfec8],
  0x639: [0xfec9, 0xfeca, 0xfecb, 0xfecc],
  0x63a: [0xfecd, 0xfece, 0xfecf, 0xfed0],
  0x641: [0xfed1, 0xfed2, 0xfed3, 0xfed4],
  0x642: [0xfed5, 0xfed6, 0xfed7, 0xfed8],
  0x643: [0xfed9, 0xfeda, 0xfedb, 0xfedc],
  0x644: [0xfedd, 0xfede, 0xfedf, 0xfee0],
  0x645: [0xfee1, 0xfee2, 0xfee3, 0xfee4],
  0x646: [0xfee5, 0xfee6, 0xfee7, 0xfee8],
  0x647: [0xfee9, 0xfeea, 0xfeeb, 0xfeec],
  0x648: [0xfeed, 0xfeee],
  0x649: [0xfeef, 0xfef0],
  0x64a: [0xfef1, 0xfef2, 0xfef3, 0xfef4],
};
// Lam followed by an alef: isolated and final ligatures.
const LAM_ALEF: Record<number, [number, number]> = { 0x622: [0xfef5, 0xfef6], 0x623: [0xfef7, 0xfef8], 0x625: [0xfef9, 0xfefa], 0x627: [0xfefb, 0xfefc] };
const TATWEEL = 0x640;

/**
 * Contextual forms for Arabic letters: for each character, the presentation form to draw (or the
 * character itself), and `0` for an alef merged into the lam before it.
 */
export function arabicForms(cps: number[]): number[] {
  const out = cps.slice();
  const joinsNext = (cp: number) => cp === TATWEEL || (FORMS[cp]?.length ?? 0) === 4;
  const joinsPrev = (cp: number) => cp === TATWEEL || (FORMS[cp]?.length ?? 0) >= 2;
  const neighbour = (i: number, step: number): number => {
    for (let k = i + step; k >= 0 && k < cps.length; k += step) if (!isMark(cps[k])) return cps[k];
    return -1;
  };
  for (let i = 0; i < cps.length; i++) {
    const cp = cps[i];
    const forms = FORMS[cp];
    if (!forms) continue;
    const prev = neighbour(i, -1);
    const next = neighbour(i, 1);
    const afterJoin = prev >= 0 && joinsNext(prev);
    if (cp === 0x644 && LAM_ALEF[next]) {
      out[i] = LAM_ALEF[next][afterJoin ? 1 : 0];
      for (let k = i + 1; k < cps.length; k++) {
        if (!isMark(cps[k])) {
          out[k] = 0;
          break;
        }
      }
      continue;
    }
    if (out[i] === 0) continue;
    const beforeJoin = forms.length === 4 && next >= 0 && joinsPrev(next);
    out[i] = forms[afterJoin && forms.length >= 2 ? (beforeJoin ? 3 : 1) : beforeJoin ? 2 : 0];
  }
  return out;
}
