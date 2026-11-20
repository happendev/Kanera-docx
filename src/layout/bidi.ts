/**
 * Right-to-left text: a simplified Unicode bidirectional algorithm (strong types, numbers and
 * neutrals; no explicit embeddings) and Arabic joining through the presentation forms most
 * Arabic fonts map.
 */

/**
 * Bidi class of a character: strong left, strong right, Arabic letter, European or Arabic number,
 * the separators and terminators that join numbers (ES, ET, CS) and other neutrals.
 */
export type BidiClass = 'L' | 'R' | 'AL' | 'EN' | 'AN' | 'ES' | 'ET' | 'CS' | 'N';

export function bidiClass(cp: number): BidiClass {
  if ((cp >= 0x30 && cp <= 0x39) || cp === 0xb2 || cp === 0xb3 || cp === 0xb9 || (cp >= 0x2070 && cp <= 0x2079 && cp !== 0x2071) || (cp >= 0x2080 && cp <= 0x2089) || (cp >= 0xff10 && cp <= 0xff19)) return 'EN';
  if ((cp >= 0x660 && cp <= 0x669) || cp === 0x66b || cp === 0x66c) return 'AN';
  if (cp >= 0x6f0 && cp <= 0x6f9) return 'EN';
  if (cp === 0x2b || cp === 0x2d || cp === 0x2212 || cp === 0x207a || cp === 0x207b || cp === 0x208a || cp === 0x208b || cp === 0xfe62 || cp === 0xfe63 || cp === 0xff0b || cp === 0xff0d) return 'ES';
  if ((cp >= 0x23 && cp <= 0x25) || (cp >= 0xa2 && cp <= 0xa5) || cp === 0xb0 || cp === 0xb1 || cp === 0x66a || (cp >= 0x2030 && cp <= 0x2034) || (cp >= 0x20a0 && cp <= 0x20cf) || cp === 0x2213 || cp === 0xff03 || cp === 0xff04 || cp === 0xff05) return 'ET';
  if (cp === 0x2c || cp === 0x2e || cp === 0x2f || cp === 0x3a || cp === 0xa0 || cp === 0x60c || cp === 0x202f || cp === 0x2044 || cp === 0xff0c || cp === 0xff0e || cp === 0xff0f || cp === 0xff1a) return 'CS';
  if (cp >= 0x590 && cp <= 0x5ff) return 'R';
  if ((cp >= 0x600 && cp <= 0x8ff) || (cp >= 0xfb50 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return isMark(cp) ? 'N' : 'AL';
  if (cp >= 0xfb1d && cp <= 0xfb4f) return 'R';
  if ((cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || cp === 0xaa || cp === 0xb5 || cp === 0xba) return 'L';
  if (cp >= 0xc0 && cp !== 0xd7 && cp !== 0xf7 && cp < 0x2000) return 'L';
  if (cp >= 0x2e80 && !(cp >= 0xfe30 && cp <= 0xfe4f) && !(cp >= 0xff00 && cp <= 0xff20)) return 'L';
  return 'N';
}

/** Characters drawn as their mirror image in right-to-left text (rule L4): brackets and the like. */
const MIRRORS: Record<number, number> = {};
for (const pair of ['()', '<>', '[]', '{}', '«»', '‹›', '⁅⁆', '⁽⁾', '₍₎', '≤≥', '⟨⟩', '〈〉', '《》', '「」', '『』', '【】', '〔〕', '（）', '［］', '｛｝', '＜＞']) {
  const [a, b] = [...pair].map((c) => c.codePointAt(0)!);
  MIRRORS[a] = b;
  MIRRORS[b] = a;
}

/** The character a right-to-left level draws in place of `cp`, or undefined when it has no mirror. */
export function mirrored(cp: number): number | undefined {
  return MIRRORS[cp];
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
  const sos: BidiClass = base ? 'R' : 'L';
  // W2: European numbers after an Arabic letter read as Arabic numbers.
  let lastStrong: BidiClass = sos;
  for (let i = 0; i < n; i++) {
    const t = types[i];
    if (t === 'L' || t === 'R' || t === 'AL') lastStrong = t;
    else if (t === 'EN' && lastStrong === 'AL') types[i] = 'AN';
  }
  for (let i = 0; i < n; i++) if (types[i] === 'AL') types[i] = 'R';
  // W4: a lone separator between two numbers joins them (1.5, 1,000, 2+3).
  for (let i = 1; i < n - 1; i++) {
    const prev = types[i - 1];
    if (prev !== types[i + 1]) continue;
    if (types[i] === 'ES' && prev === 'EN') types[i] = 'EN';
    else if (types[i] === 'CS' && (prev === 'EN' || prev === 'AN')) types[i] = prev;
  }
  // W5: currency and percent signs beside a European number belong to it.
  for (let i = 0; i < n; ) {
    if (types[i] !== 'ET') {
      i++;
      continue;
    }
    let j = i;
    while (j < n && types[j] === 'ET') j++;
    if (types[i - 1] === 'EN' || types[j] === 'EN') for (let k = i; k < j; k++) types[k] = 'EN';
    i = j;
  }
  // W6: other separators and terminators are neutral; neutrals of right-to-left runs (w:rtl) read
  // right to left.
  for (let i = 0; i < n; i++) {
    if (types[i] === 'ES' || types[i] === 'ET' || types[i] === 'CS') types[i] = 'N';
    if (types[i] === 'N' && rtl[i]) types[i] = 'R';
  }
  // W7: European numbers in left-to-right context are left to right.
  lastStrong = sos;
  for (let i = 0; i < n; i++) {
    if (types[i] === 'L' || types[i] === 'R') lastStrong = types[i];
    else if (types[i] === 'EN' && lastStrong === 'L') types[i] = 'L';
  }
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
        before = d;
        break;
      }
    }
    let after: 'L' | 'R' = base ? 'R' : 'L';
    for (let k = j; k < n; k++) {
      const d = dir(types[k]);
      if (d) {
        after = d;
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

// Presentation forms (isolated, final, initial, medial) of the Arabic letters; two forms mean the
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
  // Persian, Urdu and other letters, from Arabic Presentation Forms-A.
  0x671: [0xfb50, 0xfb51],
  0x679: [0xfb66, 0xfb67, 0xfb68, 0xfb69],
  0x67a: [0xfb5e, 0xfb5f, 0xfb60, 0xfb61],
  0x67b: [0xfb52, 0xfb53, 0xfb54, 0xfb55],
  0x67e: [0xfb56, 0xfb57, 0xfb58, 0xfb59],
  0x67f: [0xfb62, 0xfb63, 0xfb64, 0xfb65],
  0x680: [0xfb5a, 0xfb5b, 0xfb5c, 0xfb5d],
  0x683: [0xfb76, 0xfb77, 0xfb78, 0xfb79],
  0x684: [0xfb72, 0xfb73, 0xfb74, 0xfb75],
  0x686: [0xfb7a, 0xfb7b, 0xfb7c, 0xfb7d],
  0x687: [0xfb7e, 0xfb7f, 0xfb80, 0xfb81],
  0x688: [0xfb88, 0xfb89],
  0x68c: [0xfb84, 0xfb85],
  0x68d: [0xfb82, 0xfb83],
  0x68e: [0xfb86, 0xfb87],
  0x691: [0xfb8c, 0xfb8d],
  0x698: [0xfb8a, 0xfb8b],
  0x6a4: [0xfb6a, 0xfb6b, 0xfb6c, 0xfb6d],
  0x6a6: [0xfb6e, 0xfb6f, 0xfb70, 0xfb71],
  0x6a9: [0xfb8e, 0xfb8f, 0xfb90, 0xfb91],
  0x6ad: [0xfbd3, 0xfbd4, 0xfbd5, 0xfbd6],
  0x6af: [0xfb92, 0xfb93, 0xfb94, 0xfb95],
  0x6b1: [0xfb9a, 0xfb9b, 0xfb9c, 0xfb9d],
  0x6b3: [0xfb96, 0xfb97, 0xfb98, 0xfb99],
  0x6ba: [0xfb9e, 0xfb9f],
  0x6bb: [0xfba0, 0xfba1, 0xfba2, 0xfba3],
  0x6be: [0xfbaa, 0xfbab, 0xfbac, 0xfbad],
  0x6c0: [0xfba4, 0xfba5],
  0x6c1: [0xfba6, 0xfba7, 0xfba8, 0xfba9],
  0x6c5: [0xfbe0, 0xfbe1],
  0x6c6: [0xfbd9, 0xfbda],
  0x6c7: [0xfbd7, 0xfbd8],
  0x6c8: [0xfbdb, 0xfbdc],
  0x6c9: [0xfbe2, 0xfbe3],
  0x6cb: [0xfbde, 0xfbdf],
  0x6cc: [0xfbfc, 0xfbfd, 0xfbfe, 0xfbff],
  0x6d0: [0xfbe4, 0xfbe5, 0xfbe6, 0xfbe7],
  0x6d2: [0xfbae, 0xfbaf],
  0x6d3: [0xfbb0, 0xfbb1],
};

/** Whether `cp` is an Arabic letter that takes joining forms. */
export function joinsArabic(cp: number): boolean {
  return cp in FORMS;
}
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
