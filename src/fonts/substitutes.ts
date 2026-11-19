/**
 * Font substitution rules for fonts that are not available.
 *
 * Metric-compatible replacements come first: their measured glyph widths match the original's, so
 * lines break at the same words wherever the substitute covers the text (coverage, kerning and line
 * metrics also affect layout; `vmetrics` restores the original's line metrics where they differ).
 * The remaining entries are the closest open fonts by design and width.
 */

export type GenericFamily = 'serif' | 'sans' | 'mono' | 'cjk' | 'symbol';

interface SubstituteRule {
  families: string[];
  /** Whether `families` match the original's advance widths. */
  metricCompatible: boolean;
  /** Class for names that the generic heuristics would misidentify. */
  generic?: GenericFamily;
  /** Families of a different width tried after `families`, which never count as metric-compatible. */
  close?: string[];
  /**
   * Line metrics of the original font (ascent, descent, line gap in em) for substitutes whose
   * widths match but whose own vertical metrics differ, so line heights still match Word.
   */
  vmetrics?: [number, number, number];
}

// Line metrics (`vmetrics`) of the Office and web fonts below are those Word reports for them in its
// PDFs; with them a substitute of another width still breaks pages where Word does. The closest
// widths come first: Tahoma is about as wide as Liberation Sans, not DejaVu Sans.

// Preserve Aptos's typographic line metrics even when a close substitute has different advances.
const APTOS_LINE_METRICS: [number, number, number] = [1923 / 2048, 577 / 2048, 0];

const RULES: Record<string, SubstituteRule> = {
  calibri: { families: ['carlito'], metricCompatible: true },
  'calibri light': { families: ['carlito'], metricCompatible: false },
  cambria: { families: ['caladea'], metricCompatible: true, vmetrics: [1946 / 2048, 455 / 2048, 0] },
  'cambria math': { families: ['caladea'], metricCompatible: false },
  arial: { families: ['liberation sans', 'arimo'], metricCompatible: true },
  'arial mt': { families: ['liberation sans', 'arimo'], metricCompatible: true },
  helvetica: { families: ['liberation sans', 'arimo'], metricCompatible: true },
  'helvetica neue': { families: ['liberation sans', 'arimo'], metricCompatible: false },
  'arial unicode ms': { families: ['liberation sans', 'dejavu sans'], metricCompatible: false, vmetrics: [1.069, 0.271, 0] },
  // Word's Arial Narrow has a taller line than the open font despite matching advances.
  'arial narrow': { families: ['liberation sans narrow'], close: ['liberation sans'], metricCompatible: true, vmetrics: [1916 / 2048, 434 / 2048, 0] },
  'times new roman': { families: ['liberation serif', 'tinos'], metricCompatible: true },
  times: { families: ['liberation serif', 'tinos'], metricCompatible: true },
  'tms rmn': { families: ['liberation serif'], metricCompatible: false },
  'courier new': { families: ['liberation mono', 'cousine'], metricCompatible: true },
  courier: { families: ['liberation mono', 'cousine'], metricCompatible: true },
  georgia: { families: ['gelasio'], metricCompatible: true, vmetrics: [1878 / 2048, 449 / 2048, 0] },
  'segoe ui': { families: ['selawik'], metricCompatible: true },
  'segoe ui light': { families: ['selawik light'], close: ['selawik'], metricCompatible: true },
  'segoe ui semibold': { families: ['selawik semibold'], close: ['selawik'], metricCompatible: true },
  'segoe ui semilight': { families: ['selawik semilight'], close: ['selawik'], metricCompatible: true },
  roboto: { families: ['roboto'], metricCompatible: true },
  'roboto medium': { families: ['roboto medium'], close: ['roboto'], metricCompatible: true },
  aptos: { families: ['morrow sans'], close: ['selawik', 'liberation sans'], metricCompatible: true, vmetrics: APTOS_LINE_METRICS },
  'aptos display': { families: ['morrow sans display'], close: ['selawik', 'liberation sans'], metricCompatible: true, vmetrics: APTOS_LINE_METRICS },
  'aptos narrow': { families: ['morrow sans narrow'], close: ['liberation sans narrow', 'liberation sans'], metricCompatible: true, vmetrics: APTOS_LINE_METRICS },
  'aptos light': { families: ['morrow sans light'], close: ['selawik light', 'selawik'], metricCompatible: true, vmetrics: APTOS_LINE_METRICS },
  'aptos mono': { families: ['morrow mono'], close: ['liberation mono'], metricCompatible: true },
  verdana: { families: ['dejavu sans'], metricCompatible: false, vmetrics: [2059 / 2048, 430 / 2048, 0] },
  tahoma: { families: ['liberation sans'], metricCompatible: false, vmetrics: [2049 / 2048, 423 / 2048, 0] },
  'trebuchet ms': { families: ['roboto', 'liberation sans'], metricCompatible: false, vmetrics: [1923 / 2048, 455 / 2048, 0] },
  // Century Gothic was drawn on ITC Avant Garde's widths, as is TeX Gyre Adventor (available via
  // fonts or systemFonts). Word's line metrics apply whichever face stands in.
  'century gothic': { families: ['tex gyre adventor'], close: ['liberation sans'], metricCompatible: true, vmetrics: [2060 / 2048, 451 / 2048, 0] },
  'gill sans mt': { families: ['selawik', 'liberation sans'], metricCompatible: false },
  'lucida sans unicode': { families: ['dejavu sans'], metricCompatible: false, vmetrics: [1.097, 0.44, 0] },
  'lucida grande': { families: ['dejavu sans'], metricCompatible: false },
  consolas: { families: ['liberation mono'], metricCompatible: false },
  'lucida console': { families: ['liberation mono'], metricCompatible: false },
  'book antiqua': { families: ['liberation serif'], metricCompatible: false, vmetrics: [0.923, 0.282, 0] },
  'bookman old style': { families: ['gelasio', 'liberation serif'], metricCompatible: false, vmetrics: [0.942, 0.232, 0] },
  garamond: { families: ['liberation serif'], metricCompatible: false, vmetrics: [0.862, 0.263, 0] },
  'palatino linotype': { families: ['gelasio', 'liberation serif'], metricCompatible: false, vmetrics: [1.05, 0.299, 0] },
  'franklin gothic book': { families: ['liberation sans'], metricCompatible: false },
  'open sans': { families: ['open sans', 'selawik'], metricCompatible: false, vmetrics: [1.069, 0.293, 0] },
  univers: { families: ['liberation sans'], metricCompatible: false },
  century: { families: ['gelasio', 'liberation serif'], metricCompatible: false, generic: 'serif', vmetrics: [0.986, 0.216, 0] },
  'century schoolbook': { families: ['tex gyre schola', 'c059', 'gelasio', 'liberation serif'], metricCompatible: false, vmetrics: [0.986, 0.216, 0] },
  'arial rounded mt bold': { families: ['liberation sans'], metricCompatible: false, vmetrics: [0.946, 0.211, 0] },
  lato: { families: ['liberation sans'], metricCompatible: false, vmetrics: [0.987, 0.213, 0] },
  poppins: { families: ['dejavu sans'], metricCompatible: false, vmetrics: [1.05, 0.35, 0] },
  'source sans pro': { families: ['taper', 'liberation sans'], metricCompatible: false, vmetrics: [0.984, 0.273, 0] },
  montserrat: { families: ['dejavu sans'], metricCompatible: false, vmetrics: [0.968, 0.251, 0] },
  ubuntu: { families: ['liberation sans'], metricCompatible: false, vmetrics: [0.932, 0.189, 0.028] },

  // Windows and Office families. Some compatible faces are available via fonts or systemFonts.
  'arial black': { families: ['archivo black', 'roboto'], metricCompatible: false, vmetrics: [1.101, 0.31, 0] },
  bahnschrift: { families: ['gleis'], close: ['roboto'], metricCompatible: true },
  'bahnschrift condensed': { families: ['gleis condensed'], close: ['liberation sans narrow'], metricCompatible: true },
  'bahnschrift semi condensed': { families: ['gleis semicondensed'], close: ['liberation sans narrow'], metricCompatible: true },
  candara: { families: ['taper'], close: ['carlito'], metricCompatible: true },
  'candara light': { families: ['taper light'], close: ['taper', 'carlito'], metricCompatible: true },
  'cascadia code': { families: ['cascadia mono', 'dejavu sans mono', 'liberation mono'], metricCompatible: false },
  'cascadia mono': { families: ['cascadia code', 'dejavu sans mono', 'liberation mono'], metricCompatible: false },
  'comic sans ms': { families: ['comic neue'], metricCompatible: false, vmetrics: [1.102, 0.292, 0] },
  constantia: { families: ['fidela'], close: ['gelasio', 'liberation serif'], metricCompatible: true, generic: 'serif' },
  corbel: { families: ['lintel'], close: ['carlito'], metricCompatible: true },
  'franklin gothic medium': { families: ['leyden medium'], close: ['libre franklin', 'roboto'], metricCompatible: true },
  gabriola: { families: ['seraphine'], close: ['cormorant', 'liberation serif'], metricCompatible: true, generic: 'serif' },
  impact: { families: ['anvil'], close: ['oswald', 'liberation sans narrow'], metricCompatible: true },
  'ink free': { families: ['scrawl'], close: ['patrick hand', 'comic neue'], metricCompatible: true },
  'microsoft sans serif': { families: ['ordinal'], close: ['liberation sans'], metricCompatible: true },
  'segoe print': { families: ['penmark'], close: ['kalam', 'comic neue'], metricCompatible: true },
  'segoe script': { families: ['flourish'], close: ['caveat', 'comic neue'], metricCompatible: true },
  'segoe ui variable': { families: ['selawik'], metricCompatible: false },
  'segoe ui variable display': { families: ['selawik'], metricCompatible: false },
  'segoe ui variable small': { families: ['selawik'], metricCompatible: false },
  'segoe ui variable text': { families: ['selawik'], metricCompatible: false },
  'segoe ui symbol': { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols'], metricCompatible: false },
  sitka: { families: ['baranof text', 'source serif 4', 'liberation serif'], metricCompatible: false, generic: 'serif' },
  'sitka text': { families: ['baranof text'], close: ['source serif 4', 'liberation serif'], metricCompatible: true, generic: 'serif' },
  'sitka heading': { families: ['baranof heading'], close: ['source serif 4', 'liberation serif'], metricCompatible: true, generic: 'serif' },
  sylfaen: { families: ['plinth'], close: ['noto serif', 'dejavu sans'], metricCompatible: true, generic: 'serif' },

  // macOS families. Menlo and DejaVu Sans Mono share Bitstream Vera ancestry; Monaco also
  // needs a fixed pitch substitute even when the document has no font table.
  menlo: { families: ['dejavu sans mono', 'liberation mono'], metricCompatible: false, generic: 'mono' },
  monaco: { families: ['dejavu sans mono', 'liberation mono'], metricCompatible: false, generic: 'mono' },
  'sf mono': { families: ['dejavu sans mono', 'liberation mono'], metricCompatible: false },
  'american typewriter': { families: ['bitter', 'gelasio', 'liberation serif'], metricCompatible: false, generic: 'serif' },
  avenir: { families: ['onward', 'outfit', 'roboto'], metricCompatible: false },
  'avenir next': { families: ['onward', 'outfit', 'roboto'], metricCompatible: false },
  'avenir next condensed': { families: ['roboto condensed', 'liberation sans narrow'], metricCompatible: false },
  'avenir next lt pro': { families: ['onward'], close: ['outfit', 'roboto'], metricCompatible: true },
  baskerville: { families: ['libre baskerville', 'liberation serif'], metricCompatible: false },
  'bodoni 72': { families: ['libre bodoni', 'bodoni moda', 'liberation serif'], metricCompatible: false, generic: 'serif' },
  chalkboard: { families: ['comic neue'], metricCompatible: false },
  'chalkboard se': { families: ['comic neue'], metricCompatible: false },
  didot: { families: ['gfs didot', 'libre bodoni', 'liberation serif'], metricCompatible: false, generic: 'serif' },
  'din alternate': { families: ['gleis', 'roboto'], metricCompatible: false },
  'din condensed': { families: ['gleis condensed', 'liberation sans narrow'], metricCompatible: false },
  futura: { families: ['jost', 'onward', 'roboto'], metricCompatible: false },
  geneva: { families: ['dejavu sans', 'liberation sans'], metricCompatible: false },
  'gill sans': { families: ['rill', 'cabin', 'selawik'], metricCompatible: false },
  'helvetica neue condensed': { families: ['liberation sans narrow', 'roboto condensed'], metricCompatible: false },
  'hoefler text': { families: ['eb garamond', 'liberation serif'], metricCompatible: false, generic: 'serif' },
  optima: { families: ['urw classico', 'selawik'], metricCompatible: false },
  palatino: { families: ['p052', 'crimson pro', 'liberation serif'], metricCompatible: false },
  'sf pro': { families: ['roboto', 'liberation sans'], metricCompatible: false },
  'sf pro display': { families: ['roboto', 'liberation sans'], metricCompatible: false },
  'sf pro text': { families: ['roboto', 'liberation sans'], metricCompatible: false },

  // Bullets in these fonts are drawn from Unicode fonts; the line still gets the original's
  // height (Word's PDFs report Symbol's ascent as 1.005 em, taller than Calibri's).
  symbol: { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [2059 / 2048, 442 / 2048, 0] },
  wingdings: { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [1841 / 2048, 432 / 2048, 0] },
  // Windows Hebrew: David is a narrow serif (Word's PDFs report 0.735/0.249 em line metrics).
  david: { families: ['liberation serif'], metricCompatible: false, generic: 'serif', vmetrics: [0.735, 0.249, 0] },
  'ms sans serif': { families: ['liberation sans'], metricCompatible: false },
  'ms serif': { families: ['liberation serif'], metricCompatible: false },
  'small fonts': { families: ['liberation sans'], metricCompatible: false },
};

type CjkRegion = 'sc' | 'tc' | 'hk' | 'jp' | 'kr';

/** Region order when neither the font nor the text's language says which glyph forms to use. */
const CJK_REGIONS: CjkRegion[] = ['sc', 'jp', 'tc', 'hk', 'kr'];

/** Installed East Asian fonts by region and style, tried after the open Noto and Source Han ones. */
const CJK_PLATFORM: Record<CjkRegion, { serif: string[]; sans: string[] }> = {
  sc: { serif: ['simsun', 'nsimsun', 'songti sc', 'stsong'], sans: ['microsoft yahei', 'dengxian', 'simhei', 'pingfang sc', 'hiragino sans gb'] },
  tc: { serif: ['pmingliu', 'mingliu', 'songti tc'], sans: ['microsoft jhenghei', 'pingfang tc'] },
  hk: { serif: ['mingliu_hkscs'], sans: ['pingfang hk'] },
  jp: { serif: ['yu mincho', 'ms mincho', 'ms pmincho', 'hiragino mincho pron'], sans: ['yu gothic', 'ms gothic', 'ms pgothic', 'meiryo', 'hiragino sans'] },
  kr: { serif: ['batang', 'nanummyeongjo', 'applemyungjo'], sans: ['malgun gothic', 'gulim', 'dotum', 'apple sd gothic neo', 'nanumgothic'] },
};

// Which glyph forms a font was made for: SimSun is Simplified Chinese, MS Mincho Japanese.
const CJK_FONT_REGIONS: Array<[RegExp, CjkRegion]> = [
  [/simsun|simhei|simkai|kaiti|fangsong|yahei|dengxian|hiragino sans gb|pingfang sc|songti sc|stsong|stheiti|stkaiti|宋体|黑体|楷体|仿宋|微软雅黑|等线|(noto|source han).* sc$/i, 'sc'],
  [/hkscs|pingfang hk|(noto|source han).* (hk|hc)$/i, 'hk'],
  [/mingliu|jhenghei|dfkai|pingfang tc|songti tc|新細明體|細明體|標楷體|微軟正黑體|(noto|source han).* tc$/i, 'tc'],
  [/mincho|ms p?gothic|ms ui gothic|yu gothic|meiryo|hiragino|osaka|游|ＭＳ|メイリオ|ヒラギノ|(noto|source han).* jp$/i, 'jp'],
  [/malgun|batang|gulim|dotum|gungsuh|myeongjo|myungjo|nanum|apple sd gothic|맑은|바탕|굴림|돋움|궁서|(noto|source han).* (kr|k)$/i, 'kr'],
];
const CJK_CHARSETS: Record<string, CjkRegion> = { '80': 'jp', '81': 'kr', '86': 'sc', '88': 'tc' };
const CJK_SERIF = /mincho|ming|song|sun|batang|gungsuh|myeongjo|myungjo|kai|fangsong|serif|明朝|明體|宋|楷|仿宋|바탕|궁서/i;

function languageRegion(lang: string | undefined): CjkRegion | undefined {
  const l = lang?.toLowerCase() ?? '';
  if (l.startsWith('ja')) return 'jp';
  if (l.startsWith('ko')) return 'kr';
  if (/^zh-(hk|hant-hk)/.test(l)) return 'hk';
  if (/^zh-(tw|mo|hant)/.test(l)) return 'tc';
  if (l.startsWith('zh')) return 'sc';
  return undefined;
}

/**
 * East Asian fallback families for a font, in order: those of the font's own region first (from its
 * name or font-table charset, then the text's East Asian language), each in the font's style (Mincho
 * and Song are serif, Gothic and Hei sans; a Latin font's own class otherwise), then other regions.
 * The Noto CJK fonts come per region (see notofonts/noto-cjk): "Noto Serif CJK SC" for SimSun.
 */
export function cjkFallbacks(family: string, lang?: string, hints?: FontTableHints): string[] {
  const generic = genericFamily(family, hints);
  const region = CJK_FONT_REGIONS.find(([re]) => re.test(family.trim()))?.[1]
    ?? (generic === 'cjk' ? CJK_CHARSETS[hints?.charset?.toLowerCase() ?? ''] : undefined)
    ?? languageRegion(lang);
  const serif = generic === 'cjk' ? CJK_SERIF.test(family) : generic === 'serif';
  const regions = region ? [region, ...(region === 'hk' ? ['tc' as const] : region === 'tc' ? ['hk' as const] : []), ...CJK_REGIONS] : CJK_REGIONS;
  const sourceHan: Record<CjkRegion, string> = { sc: 'sc', tc: 'tc', hk: 'hc', jp: 'jp', kr: 'k' };
  const out: string[] = [];
  for (const r of new Set(regions)) {
    for (const style of serif ? (['serif', 'sans'] as const) : (['sans', 'serif'] as const)) {
      out.push(`noto ${style} cjk ${r}`, `noto ${style} ${r}`, `source han ${style} ${sourceHan[r]}`, ...CJK_PLATFORM[r][style]);
    }
  }
  out.push('noto sans cjk', 'source han sans', 'droid sans fallback');
  return [...new Set(out)];
}

export const GENERIC_FALLBACKS: Record<GenericFamily, string[]> = {
  serif: ['liberation serif', 'tinos', 'dejavu serif', 'dejavu sans'],
  sans: ['liberation sans', 'arimo', 'dejavu sans'],
  mono: ['liberation mono', 'cousine', 'dejavu sans mono', 'dejavu sans'],
  cjk: cjkFallbacks(''),
  symbol: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'],
};

/** Families tried, in order, for characters the chosen font lacks; `cjk` is from cjkFallbacks(). */
export function glyphFallbacks(cjk: string[]): string[] {
  return ['dejavu sans', 'liberation sans', 'noto sans symbols2', 'noto sans symbols', 'noto sans math', ...cjk, 'symbola'];
}

export function normalizeFamily(name: string): string {
  return name
    .toLowerCase()
    .replace(/\s*\((w1|tt|ce|cyr|greek|tur|baltic|vietnamese|hebrew|arabic|body|headings)\)\s*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Strip a face's style suffix after trying the full name as a family. */
export function unstyledFamily(normalized: string): string {
  return normalized.replace(/\s+(?:(?:bold )?(?:italic|oblique)|bold|regular)$/, '');
}

function fontRule(normalized: string): SubstituteRule | undefined {
  if (Object.hasOwn(RULES, normalized)) return RULES[normalized];
  const unstyled = unstyledFamily(normalized);
  return Object.hasOwn(RULES, unstyled) ? RULES[unstyled] : undefined;
}

// Line metrics (ascent, descent, line gap in em) of the common East Asian system fonts. Their
// ideographs are one em wide like the Noto CJK fallbacks, but their lines are shorter than Noto's
// 1.45em, so without these single-spaced East Asian text would grow taller than in Word. Word gives
// SimSun and its kin about 1.3em, the extra split evenly above and below the em box.
// Fonts Word for Mac lacks (the _GB2312 ones, DFKai-SB) are left out: it draws them in YaHei.
const SONG_METRICS: [number, number, number] = [258 / 256, 74 / 256, 0];
const MING_METRICS: [number, number, number] = [205 / 256, 51 / 256, 0];
const DENG_METRICS: [number, number, number] = [1659 / 2048, 475 / 2048, 0];
const CJK_LINE_METRICS: Record<string, [number, number, number]> = {
  simsun: SONG_METRICS, nsimsun: SONG_METRICS, 宋体: SONG_METRICS, 新宋体: SONG_METRICS,
  simhei: SONG_METRICS, 黑体: SONG_METRICS, kaiti: SONG_METRICS, 楷体: SONG_METRICS,
  fangsong: SONG_METRICS, 仿宋: SONG_METRICS,
  'ms mincho': SONG_METRICS, 'ms pmincho': SONG_METRICS, 'ms gothic': SONG_METRICS, 'ms pgothic': SONG_METRICS,
  'ｍｓ 明朝': SONG_METRICS, 'ｍｓ ｐ明朝': SONG_METRICS, 'ｍｓ ゴシック': SONG_METRICS, 'ｍｓ ｐゴシック': SONG_METRICS,
  pmingliu: MING_METRICS, mingliu: MING_METRICS, mingliu_hkscs: MING_METRICS, 新細明體: MING_METRICS, 細明體: MING_METRICS,
  batang: [0.858, 0.142, 0], 바탕: [0.858, 0.142, 0],
  'microsoft yahei': [2167 / 2048, 536 / 2048, 0], 微软雅黑: [2167 / 2048, 536 / 2048, 0],
  dengxian: DENG_METRICS, 'dengxian light': DENG_METRICS, 等线: DENG_METRICS, '等线 light': DENG_METRICS,
  'microsoft jhenghei': [1.076, 0.254, 0], 微軟正黑體: [1.076, 0.254, 0],
  'malgun gothic': [1.088, 0.242, 0], '맑은 고딕': [1.088, 0.242, 0],
  meiryo: [1.06, 0.44, 0], メイリオ: [1.06, 0.44, 0],
};

/** The original line metrics of a missing East Asian system font, which its fallback keeps. */
export function cjkLineMetrics(family: string): [number, number, number] | undefined {
  const n = unstyledFamily(normalizeFamily(family));
  return Object.hasOwn(CJK_LINE_METRICS, n) ? CJK_LINE_METRICS[n] : undefined;
}

/** Families tried, in order, for a missing font; `compatible` are those with the same widths. */
export function substituteFamilies(family: string): { families: string[]; compatible: string[]; vmetrics?: [number, number, number] } | null {
  const rule = fontRule(normalizeFamily(family));
  return rule ? { families: [...rule.families, ...(rule.close ?? [])], compatible: rule.metricCompatible ? rule.families : [], vmetrics: rule.vmetrics } : null;
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
  if (/(gothic|mincho|ming|song|hei|kai|yahei|dengxian|等线|simsun|batang|gulim|dotum|malgun|meiryo|jhenghei|pmingliu|pingfang|hiragino|ヒラギノ|游|ＭＳ)/i.test(name) && !/century gothic|franklin gothic|copperplate gothic|news gothic|trade gothic|letter gothic/.test(n))
    return 'cjk';
  const known = fontRule(n)?.generic;
  if (known) return known;
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
