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
  // Word for Mac adds macOS Helvetica's 0.2 em line gap above the text: a single line is 1.2 em,
  // not Arial's 1.15. Helvetica Neue's lines are 1.2 em too.
  helvetica: { families: ['liberation sans', 'arimo'], metricCompatible: true, vmetrics: [0.77, 0.23, 0.2] },
  'helvetica neue': { families: ['liberation sans', 'arimo'], metricCompatible: false, vmetrics: [0.952, 0.213, 0.035] },
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
  'segoe ui black': { families: ['selawik'], metricCompatible: false, generic: 'sans', vmetrics: [2210 / 2048, 514 / 2048, 0] },
  // An Office cloud font for Arabic and Latin; its Arabic is drawn by a fallback at Dubai's width (see ARABIC_WIDTHS).
  dubai: { families: ['liberation sans', 'arimo'], metricCompatible: false, generic: 'sans', vmetrics: [1.129, 0.559, 0] },
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
  // Century Gothic was drawn on ITC Avant Garde's widths, as is TeX Gyre Adventor (in the full
  // pack). Word's line metrics apply whichever face stands in.
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
  // Word's Corbel lines are 1.208 em (Word for Mac's PDFs), Carlito's 1.221.
  corbel: { families: ['lintel'], close: ['carlito'], metricCompatible: true, vmetrics: [0.95, 0.258, 0] },
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
  geneva: { families: ['dejavu sans', 'liberation sans'], metricCompatible: false, vmetrics: [1, 0.25, 0.08] },
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
  'wingdings 2': { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [0.843, 0.211, 0] },
  'wingdings 3': { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [0.928, 0.211, 0] },
  webdings: { families: ['dejavu sans', 'noto sans symbols2', 'noto sans symbols', 'symbola'], metricCompatible: false, vmetrics: [0.8, 0.2, 0] },
  // Windows Hebrew: David is a narrow serif (Word's PDFs report 0.735/0.249 em line metrics).
  david: { families: ['liberation serif'], metricCompatible: false, generic: 'serif', vmetrics: [0.735, 0.249, 0] },
  'ms sans serif': { families: ['liberation sans'], metricCompatible: false },
  'ms serif': { families: ['liberation serif'], metricCompatible: false },
  'small fonts': { families: ['liberation sans'], metricCompatible: false },
  'britannic bold': { families: ['liberation sans'], metricCompatible: false, generic: 'sans', vmetrics: [0.833, 0.167, 0] },
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
// Kai (楷) is a brush script, neither Song nor Hei. It falls back to sans: Word for Mac draws 標楷體
// (DFKai-SB) and KaiTi_GB2312 in YaHei, and Noto Sans matched Word's Kai text as closely overall.
const CJK_SERIF = /mincho|ming|song|sun|batang|gungsuh|myeongjo|myungjo|fangsong|serif|明朝|明體|宋|仿宋|바탕|궁서/i;

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
 * East Asian fallback families for a font, in order: those in the font's style (Mincho and Song are
 * serif, Gothic and Hei sans; a Latin font's own class otherwise), the font's own region first (from
 * its name or font-table charset, then the text's East Asian language), then the other style.
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
  // Style before region: a Mincho or Song text reads closer in another region's serif (the forms
  // of a few characters differ) than in its own region's sans, the shape of every character.
  for (const style of serif ? (['serif', 'sans'] as const) : (['sans', 'serif'] as const)) {
    for (const r of new Set(regions)) {
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

/**
 * Fonts that carry their own Arabic and Hebrew on Windows and Mac. Word sets text in those scripts
 * in the font itself, so it sits on that font's lines even where a fallback has to draw it here.
 * Their stand-ins count too: a missing font falls back to them here as it does to Arial or Times New
 * Roman in Word.
 */
const OWN_ARABIC_HEBREW = new Set([
  'arial', 'times new roman', 'courier new', 'tahoma', 'segoe ui', 'microsoft sans serif', 'arial unicode ms',
  'liberation sans', 'liberation serif', 'liberation mono', 'arimo', 'tinos', 'cousine',
]);

/** Whether `cp` is Arabic or Hebrew that the font `family` really has (see OWN_ARABIC_HEBREW). */
export function ownArabicHebrew(family: string, cp: number): boolean {
  const script = (cp >= 0x590 && cp <= 0x6ff) || (cp >= 0x750 && cp <= 0x77f) || (cp >= 0x8a0 && cp <= 0x8ff) || (cp >= 0xfb1d && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfefe);
  return script && OWN_ARABIC_HEBREW.has(normalizeFamily(family));
}

/**
 * Arabic text in its joined forms (الرقم الاستدلالي للأثمان عند الاستهلاك في المملكة العربية
 * السعودية خلال شهر يونيو مقارنة بالشهر السابق, without spaces), to compare one font's Arabic
 * width with another's.
 */
export const ARABIC_SAMPLE = String.fromCodePoint(
  0xfe8d, 0xfedf, 0xfeae, 0xfed7, 0xfee2, 0xfe8d, 0xfefb, 0xfeb3, 0xfe98, 0xfeaa, 0xfefb, 0xfedf, 0xfef2, 0xfedf,
  0xfef8, 0xfe9b, 0xfee4, 0xfe8e, 0xfee5, 0xfecb, 0xfee8, 0xfeaa, 0xfe8d, 0xfefb, 0xfeb3, 0xfe98, 0xfeec, 0xfefc,
  0xfed9, 0xfed3, 0xfef2, 0xfe8d, 0xfedf, 0xfee4, 0xfee4, 0xfee0, 0xfedc, 0xfe94, 0xfe8d, 0xfedf, 0xfecc, 0xfeae,
  0xfe91, 0xfef4, 0xfe94, 0xfe8d, 0xfedf, 0xfeb4, 0xfecc, 0xfeee, 0xfea9, 0xfef3, 0xfe94, 0xfea7, 0xfefc, 0xfedd,
  0xfeb7, 0xfeec, 0xfeae, 0xfef3, 0xfeee, 0xfee7, 0xfef4, 0xfeee, 0xfee3, 0xfed8, 0xfe8e, 0xfead, 0xfee7, 0xfe94,
  0xfe91, 0xfe8e, 0xfedf, 0xfeb8, 0xfeec, 0xfeae, 0xfe8d, 0xfedf, 0xfeb4, 0xfe8e, 0xfe91, 0xfed6,
);

/**
 * Width (em) of ARABIC_SAMPLE in the Arabic of Word's own fonts, regular and bold, from Word's PDFs
 * set beside the fallbacks drawn here. A fallback drawing that Arabic is scaled to the same width,
 * so lines break where Word's do: DejaVu Sans's Arabic is about 1.3 times as wide as Arial's.
 * Stand-ins take their font's widths, since a missing font falls back to them as to Arial in Word.
 * (Courier New's Arabic is monospaced, which no scale imitates.)
 */
const ARABIC_WIDTHS: Record<string, [number, number]> = {
  arial: [30.3, 31.1],
  'times new roman': [29.8, 30.9],
  // An Office cloud font: Word for Mac downloads it, so a document asking for it gets its width.
  dubai: [35.05, 36.4],
};
const ARABIC_STAND_INS: Record<string, string> = {
  'liberation sans': 'arial', arimo: 'arial',
  'liberation serif': 'times new roman', tinos: 'times new roman',
};

/** Width of ARABIC_SAMPLE in the Arabic of `family` (see ARABIC_WIDTHS), when known. */
export function arabicWidth(family: string, bold: boolean): number | undefined {
  const norm = normalizeFamily(family);
  return ARABIC_WIDTHS[ARABIC_STAND_INS[norm] ?? norm]?.[bold ? 1 : 0];
}

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
// DengXian's own metrics are 1659 and 475 units; Word makes its lines 1.3 times that, like SimSun's.
const DENG_METRICS: [number, number, number] = [(1659 + 0.15 * 2134) / 2048, (475 + 0.15 * 2134) / 2048, 0];
// Meiryo's own metrics are 1.06 and 0.44em; Word makes its lines 1.3 times that too, so 10.5pt
// Meiryo takes two lines of an 18pt document grid.
const MEIRYO_METRICS: [number, number, number] = [1.06 + 0.15 * 1.5, 0.44 + 0.15 * 1.5, 0];
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
  meiryo: MEIRYO_METRICS, メイリオ: MEIRYO_METRICS,
};

// East Asian fonts of fixed pitch draw ASCII half an em wide (the proportional ones, PMingLiU or
// MS PGothic, and the newer UI fonts like YaHei or Meiryo, have Latin designs of their own).
const HALF_WIDTH_LATIN = new Set([
  'simsun', '宋体', 'nsimsun', '新宋体', 'simhei', '黑体', 'kaiti', '楷体', 'fangsong', '仿宋',
  'mingliu', '細明體', 'mingliu_hkscs', 'dfkai-sb', '標楷體',
  'ms mincho', 'ｍｓ 明朝', 'ms gothic', 'ｍｓ ゴシック',
  'batangche', '바탕체', 'gulimche', '굴림체', 'dotumche', '돋움체', 'gungsuhche', '궁서체',
]);

/** Whether `family` is an East Asian font whose ASCII characters are all half an em wide. */
export function halfWidthLatin(family: string): boolean {
  return HALF_WIDTH_LATIN.has(unstyledFamily(normalizeFamily(family)));
}

// Noto Sans CJK's line metrics (its win ascent and descent). An East Asian font of unknown metrics
// keeps them when Noto Serif CJK stands in, whose lines are 0.8% shorter, so the glyph style never
// moves the layout: on a document grid that small difference can halve the height of every line.
export const NOTO_SANS_CJK_LINE: [number, number, number] = [1.16, 0.288, 0];

/** The original line metrics of a missing East Asian system font, which its fallback keeps. */
export function cjkLineMetrics(family: string): [number, number, number] | undefined {
  const n = unstyledFamily(normalizeFamily(family));
  return Object.hasOwn(CJK_LINE_METRICS, n) ? CJK_LINE_METRICS[n] : undefined;
}

// Word on Windows draws an East Asian character that the requested East Asian font lacks from a
// fallback font of its own, and the line grows to about 1.71em, the height Word gives Microsoft
// YaHei and Malgun Gothic. Chinese and Japanese fonts have no Hangul, and a Japanese font (its
// Shift_JIS set, extended by JIS X 0213 with traditional forms like 檔) lacks the simplified ones.
const EAST_ASIAN_FALLBACK_LINE: [number, number, number] = [1.28, 0.43, 0];
let jisChars: Set<number> | null | undefined;
let gb2312Chars: Set<number> | null | undefined;

/** The characters of a double-byte legacy charset, or null where the runtime can't decode it. */
function legacyCharset(label: string, leads: [number, number], trails: [number, number]): Set<number> | null {
  try {
    const decoder = new TextDecoder(label);
    const out = new Set<number>();
    const pair = new Uint8Array(2);
    for (let a = leads[0]; a <= leads[1]; a++) {
      for (let b = trails[0]; b <= trails[1]; b++) {
        pair[0] = a;
        pair[1] = b;
        const s = decoder.decode(pair);
        if (s.length === 1 && s !== '\ufffd') out.add(s.charCodeAt(0));
      }
    }
    return out;
  } catch {
    return null;
  }
}

const isHangul = (cp: number) => (cp >= 0xac00 && cp <= 0xd7a3) || (cp >= 0x1100 && cp <= 0x11ff) || (cp >= 0x3130 && cp <= 0x318f);

/** Word's line metrics for an East Asian character that the East Asian system font `family` lacks. */
export function cjkFallbackLine(family: string, cp: number): [number, number, number] | undefined {
  const han = cp >= 0x4e00 && cp <= 0x9fff;
  if (!(han || isHangul(cp)) || !cjkLineMetrics(family)) return undefined;
  const region = CJK_FONT_REGIONS.find(([re]) => re.test(family.trim()))?.[1];
  if (!region || region === 'kr') return undefined;
  if (!han) return EAST_ASIAN_FALLBACK_LINE;
  if (region !== 'jp') return undefined;
  jisChars ??= legacyCharset('shift_jis', [0x81, 0xfc], [0x40, 0xfc]);
  gb2312Chars ??= legacyCharset('gbk', [0xb0, 0xf7], [0xa1, 0xfe]);
  return jisChars && gb2312Chars && !jisChars.has(cp) && gb2312Chars.has(cp) ? EAST_ASIAN_FALLBACK_LINE : undefined;
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
  if (/serif|roman|times|georgia|garamond|book|cambria|palatino|minion|baskerville|century|naskh/.test(n)) return 'serif';
  return 'sans';
}
