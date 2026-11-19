/**
 * Font substitution rules for fonts that are not available.
 *
 * Metric-compatible replacements come first: they have identical advance widths and vertical
 * metrics, so line breaks and pagination match Word exactly. The remaining entries are the
 * closest open fonts by design and width.
 */

export type GenericFamily = 'serif' | 'sans' | 'mono' | 'cjk' | 'symbol';

interface SubstituteRule {
  families: string[];
  metricCompatible: boolean;
  /**
   * Line metrics of the original font (ascent, descent, line gap in em) for substitutes whose
   * widths match but whose own vertical metrics differ, so line heights still match Word.
   */
  vmetrics?: [number, number, number];
}

const RULES: Record<string, SubstituteRule> = {
  calibri: { families: ['carlito'], metricCompatible: true },
  'calibri light': { families: ['carlito'], metricCompatible: false },
  cambria: { families: ['caladea'], metricCompatible: true, vmetrics: [1946 / 2048, 455 / 2048, 0] },
  'cambria math': { families: ['caladea'], metricCompatible: false },
  arial: { families: ['liberation sans', 'arimo'], metricCompatible: true },
  'arial mt': { families: ['liberation sans', 'arimo'], metricCompatible: true },
  helvetica: { families: ['liberation sans', 'arimo'], metricCompatible: true },
  'helvetica neue': { families: ['liberation sans', 'arimo'], metricCompatible: false },
  'arial unicode ms': { families: ['liberation sans', 'dejavu sans'], metricCompatible: false },
  'arial narrow': { families: ['liberation sans narrow', 'liberation sans'], metricCompatible: true },
  'times new roman': { families: ['liberation serif', 'tinos'], metricCompatible: true },
  times: { families: ['liberation serif', 'tinos'], metricCompatible: true },
  'tms rmn': { families: ['liberation serif'], metricCompatible: false },
  'courier new': { families: ['liberation mono', 'cousine'], metricCompatible: true },
  courier: { families: ['liberation mono', 'cousine'], metricCompatible: true },
  georgia: { families: ['gelasio'], metricCompatible: true, vmetrics: [1878 / 2048, 449 / 2048, 0] },
  'segoe ui': { families: ['selawik'], metricCompatible: true },
  'segoe ui light': { families: ['selawik light', 'selawik'], metricCompatible: true },
  'segoe ui semibold': { families: ['selawik semibold', 'selawik'], metricCompatible: true },
  'segoe ui semilight': { families: ['selawik semilight', 'selawik'], metricCompatible: true },
  roboto: { families: ['roboto'], metricCompatible: true },
  'roboto medium': { families: ['roboto medium', 'roboto'], metricCompatible: true },
  aptos: { families: ['selawik', 'liberation sans'], metricCompatible: false },
  'aptos display': { families: ['selawik', 'liberation sans'], metricCompatible: false },
  'aptos narrow': { families: ['liberation sans narrow', 'liberation sans'], metricCompatible: false },
  'aptos light': { families: ['selawik light', 'selawik'], metricCompatible: false },
  verdana: { families: ['dejavu sans'], metricCompatible: false },
  tahoma: { families: ['dejavu sans', 'liberation sans'], metricCompatible: false },
  'trebuchet ms': { families: ['selawik', 'liberation sans'], metricCompatible: false },
  'century gothic': { families: ['liberation sans'], metricCompatible: false },
  'gill sans mt': { families: ['selawik', 'liberation sans'], metricCompatible: false },
  'lucida sans unicode': { families: ['dejavu sans'], metricCompatible: false },
  'lucida grande': { families: ['dejavu sans'], metricCompatible: false },
  consolas: { families: ['liberation mono'], metricCompatible: false },
  'lucida console': { families: ['liberation mono'], metricCompatible: false },
  'book antiqua': { families: ['liberation serif'], metricCompatible: false },
  'bookman old style': { families: ['liberation serif'], metricCompatible: false },
  garamond: { families: ['liberation serif'], metricCompatible: false },
  'palatino linotype': { families: ['liberation serif'], metricCompatible: false },
  'franklin gothic book': { families: ['liberation sans'], metricCompatible: false },
  'open sans': { families: ['open sans', 'selawik'], metricCompatible: false },
  univers: { families: ['liberation sans'], metricCompatible: false },
  // Bullets in these fonts are drawn from Unicode fonts; the line still gets the original's
  // height (Word's PDFs report Symbol's ascent as 1.005 em, taller than Calibri's).
  symbol: { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [2059 / 2048, 442 / 2048, 0] },
  wingdings: { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [1841 / 2048, 432 / 2048, 0] },
  'ms sans serif': { families: ['liberation sans'], metricCompatible: false },
  'ms serif': { families: ['liberation serif'], metricCompatible: false },
  'small fonts': { families: ['liberation sans'], metricCompatible: false },
};

const CJK_FAMILIES = [
  'noto sans cjk sc',
  'noto sans cjk jp',
  'noto sans cjk tc',
  'noto sans cjk kr',
  'noto sans cjk',
  'source han sans',
  'source han sans sc',
  'noto serif cjk sc',
  'noto sans sc',
  'noto sans jp',
  'ms gothic',
  'ms mincho',
  'yu gothic',
  'simsun',
  'microsoft yahei',
  'malgun gothic',
  'droid sans fallback',
];

export const GENERIC_FALLBACKS: Record<GenericFamily, string[]> = {
  serif: ['liberation serif', 'tinos', 'dejavu serif', 'dejavu sans'],
  sans: ['liberation sans', 'arimo', 'dejavu sans'],
  mono: ['liberation mono', 'cousine', 'dejavu sans mono', 'dejavu sans'],
  cjk: CJK_FAMILIES,
  symbol: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'],
};

/** Families tried, in order, for characters the chosen font lacks. */
export const GLYPH_FALLBACKS = ['dejavu sans', 'liberation sans', 'noto sans symbols2', 'noto sans symbols', 'noto sans math', ...CJK_FAMILIES, 'symbola'];

export function normalizeFamily(name: string): string {
  return name
    .toLowerCase()
    .replace(/\s*\((w1|tt|ce|cyr|greek|tur|baltic|vietnamese|hebrew|arabic|body|headings)\)\s*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function substituteFamilies(family: string): { families: string[]; metricCompatible: boolean; vmetrics?: [number, number, number] } | null {
  const rule = RULES[normalizeFamily(family)];
  return rule ? { families: rule.families, metricCompatible: rule.metricCompatible, vmetrics: rule.vmetrics } : null;
}

export interface FontTableHints {
  /** w:family: roman | swiss | modern | script | decorative | auto */
  family?: string;
  /** w:pitch: fixed | variable | default */
  pitch?: string;
  /** w:panose1 hex string */
  panose?: string;
  /** w:charset hex */
  charset?: string;
  altName?: string;
}

/** Best guess at a font's generic class from Word's font table info and its name. */
export function genericFamily(name: string, hints?: FontTableHints): GenericFamily {
  const n = normalizeFamily(name);
  if (/^(symbol|wingdings|webdings)/.test(n)) return 'symbol';
  const charset = hints?.charset?.toLowerCase();
  if (charset && ['80', '81', '86', '88', '82'].includes(charset)) return 'cjk';
  if (/(gothic|mincho|ming|song|hei|kai|yahei|simsun|batang|gulim|dotum|malgun|meiryo|jhenghei|pmingliu|ヒラギノ|游|ＭＳ)/i.test(name) && !/century gothic|franklin gothic|copperplate gothic|news gothic|trade gothic|letter gothic/.test(n))
    return 'cjk';
  if (hints?.pitch === 'fixed' || hints?.family === 'modern' || /mono|courier|consol|typewriter|code/.test(n)) return 'mono';
  if (hints?.panose && hints.panose.length >= 4) {
    const familyKind = parseInt(hints.panose.slice(0, 2), 16);
    const serifStyle = parseInt(hints.panose.slice(2, 4), 16);
    if (familyKind === 2) {
      if (serifStyle >= 11 && serifStyle <= 13) return 'sans';
      if (serifStyle >= 2 && serifStyle <= 10) return 'serif';
    }
  }
  if (hints?.family === 'roman') return 'serif';
  if (hints?.family === 'swiss') return 'sans';
  if (/sans|gothic|grotesk|arial|helvet|verdana|tahoma|segoe|calibri|aptos|roboto|open|lato|montserrat|univers|futura/.test(n)) return 'sans';
  if (/serif|roman|times|georgia|garamond|book|cambria|palatino|minion|baskerville|century/.test(n)) return 'serif';
  return 'sans';
}
