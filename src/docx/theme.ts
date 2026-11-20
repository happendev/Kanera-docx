import { child, children, descendant, type XmlElement } from '../util/xml.js';

export interface Theme {
  colors: Record<string, string>;
  majorLatin: string;
  minorLatin: string;
  majorEastAsia: string;
  minorEastAsia: string;
  majorCs: string;
  minorCs: string;
  /** The format scheme's fills (a:fillStyleLst), which shape styles reference by index (a:fillRef). */
  fillStyles: XmlElement[];
  hasTheme: boolean;
}

export const DEFAULT_THEME: Theme = {
  colors: {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '4472C4',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '5B9BD5',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
  },
  majorLatin: 'Calibri Light',
  minorLatin: 'Calibri',
  majorEastAsia: '',
  minorEastAsia: '',
  majorCs: '',
  minorCs: '',
  fillStyles: [],
  hasTheme: false,
};

/** The theme's script font for a language (w:themeFontLang), as DrawingML names scripts. */
function themeScript(lang: string | undefined): string | undefined {
  const l = lang?.toLowerCase() ?? '';
  if (!l) return undefined;
  if (l.startsWith('ja')) return 'Jpan';
  if (l.startsWith('ko')) return 'Hang';
  if (/^zh-(tw|hk|mo|hant)/.test(l)) return 'Hant';
  if (l.startsWith('zh')) return 'Hans';
  if (/^(ar|fa|ur|ps|sd|ug)\b/.test(l)) return 'Arab';
  if (l.startsWith('he') || l.startsWith('yi')) return 'Hebr';
  if (l.startsWith('th')) return 'Thai';
  return undefined;
}

/**
 * The theme's fonts and colours. `langs` is the document's w:themeFontLang: an empty East Asian or
 * complex script slot takes the theme's font for that language's script (游明朝 for Japanese).
 */
export function parseTheme(root: XmlElement | undefined, langs: { eastAsia?: string; bidi?: string } = {}): Theme {
  if (!root) return DEFAULT_THEME;
  const theme: Theme = { ...DEFAULT_THEME, colors: { ...DEFAULT_THEME.colors }, hasTheme: true };
  const scheme = descendant(root, 'a:clrScheme');
  if (scheme) {
    for (const c of children(scheme)) {
      const name = c.name.slice(2);
      const srgb = child(c, 'a:srgbClr');
      const sys = child(c, 'a:sysClr');
      const val = srgb?.attrs.val ?? sys?.attrs.lastClr ?? (sys?.attrs.val === 'window' ? 'FFFFFF' : sys?.attrs.val === 'windowText' ? '000000' : undefined);
      if (val) theme.colors[name] = val.toUpperCase();
    }
  }
  const fonts = descendant(root, 'a:fontScheme');
  if (fonts) {
    const major = child(fonts, 'a:majorFont');
    const minor = child(fonts, 'a:minorFont');
    theme.majorLatin = child(major, 'a:latin')?.attrs.typeface || theme.majorLatin;
    theme.minorLatin = child(minor, 'a:latin')?.attrs.typeface || theme.minorLatin;
    const script = (font: XmlElement | undefined, lang: string | undefined) => {
      const s = themeScript(lang);
      return (s && children(font, 'a:font').find((f) => f.attrs.script === s)?.attrs.typeface) || '';
    };
    theme.majorEastAsia = child(major, 'a:ea')?.attrs.typeface || script(major, langs.eastAsia);
    theme.minorEastAsia = child(minor, 'a:ea')?.attrs.typeface || script(minor, langs.eastAsia);
    theme.majorCs = child(major, 'a:cs')?.attrs.typeface || script(major, langs.bidi);
    theme.minorCs = child(minor, 'a:cs')?.attrs.typeface || script(minor, langs.bidi);
  }
  theme.fillStyles = children(descendant(root, 'a:fillStyleLst'));
  return theme;
}

/** Word's w:themeColor names map onto the DrawingML scheme slots. */
export function themeColorValue(theme: Theme, name: string, mapping?: Record<string, string>): string | undefined {
  const map: Record<string, string> = {
    dark1: 'dk1',
    light1: 'lt1',
    dark2: 'dk2',
    light2: 'lt2',
    text1: mapping?.t1 ?? 'dk1',
    background1: mapping?.bg1 ?? 'lt1',
    text2: mapping?.t2 ?? 'dk2',
    background2: mapping?.bg2 ?? 'lt2',
    tx1: mapping?.t1 ?? 'dk1',
    bg1: mapping?.bg1 ?? 'lt1',
    tx2: mapping?.t2 ?? 'dk2',
    bg2: mapping?.bg2 ?? 'lt2',
    hyperlink: 'hlink',
    followedHyperlink: 'folHlink',
    phClr: 'accent1',
  };
  const slot = map[name] ?? name;
  const norm = slot === 'dark1' ? 'dk1' : slot === 'light1' ? 'lt1' : slot === 'dark2' ? 'dk2' : slot === 'light2' ? 'lt2' : slot;
  return theme.colors[norm];
}

export function themeFont(theme: Theme, ref: string): string | undefined {
  switch (ref) {
    case 'majorAscii':
    case 'majorHAnsi':
      return theme.majorLatin;
    case 'minorAscii':
    case 'minorHAnsi':
      return theme.minorLatin;
    case 'majorEastAsia':
      return theme.majorEastAsia || undefined;
    case 'minorEastAsia':
      return theme.minorEastAsia || undefined;
    case 'majorBidi':
      return theme.majorCs || undefined;
    case 'minorBidi':
      return theme.minorCs || undefined;
    default:
      return undefined;
  }
}
