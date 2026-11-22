import { coversEmoji } from './emoji.js';
import { isEastAsian } from './script.js';

/** Lazy loaders use these family names to match the faces in each font's name table. */
export interface BundledFont {
  covers?: (cp: number) => boolean;
  file: string;
  family: string;
  bold: boolean;
  italic: boolean;
  weight: number;
}

const four = (family: string, prefix: string): BundledFont[] => [
  { file: `${prefix}-Regular.ttf`, family, bold: false, italic: false, weight: 400 },
  { file: `${prefix}-Bold.ttf`, family, bold: true, italic: false, weight: 700 },
  { file: `${prefix}-Italic.ttf`, family, bold: false, italic: true, weight: 400 },
  { file: `${prefix}-BoldItalic.ttf`, family, bold: true, italic: true, weight: 700 },
];

/** The opt-in browser pack keeps common Word substitutes and general-purpose glyph fallbacks. */
export const COMMON_FONTS: BundledFont[] = [
  ...four('Carlito', 'Carlito'),
  ...four('Caladea', 'Caladea'),
  ...four('Liberation Sans', 'LiberationSans'),
  ...four('Liberation Serif', 'LiberationSerif'),
  ...four('Liberation Mono', 'LiberationMono'),
  ...four('Gelasio', 'Gelasio'),
  ...four('Morrow Sans', 'MorrowSans'),
  ...four('Morrow Sans Display', 'MorrowSansDisplay'),
  ...four('Morrow Sans Narrow', 'MorrowSans-Narrow'),
  { file: 'MorrowSans-Light.ttf', family: 'Morrow Sans Light', bold: false, italic: false, weight: 300 },
  { file: 'MorrowSans-LightItalic.ttf', family: 'Morrow Sans Light', bold: false, italic: true, weight: 300 },
  { file: 'Selawik-Regular.ttf', family: 'Selawik', bold: false, italic: false, weight: 400 },
  { file: 'Selawik-Bold.ttf', family: 'Selawik', bold: true, italic: false, weight: 700 },
  { file: 'Selawik-Light.ttf', family: 'Selawik Light', bold: false, italic: false, weight: 300 },
  { file: 'Selawik-Semibold.ttf', family: 'Selawik Semibold', bold: false, italic: false, weight: 600 },
  { file: 'DejaVuSans.ttf', family: 'DejaVu Sans', bold: false, italic: false, weight: 400 },
  { file: 'DejaVuSans-Bold.ttf', family: 'DejaVu Sans', bold: true, italic: false, weight: 700 },
];

/** Font faces available by default in Node and the browser. */
export const FULL_FONTS: BundledFont[] = [
  ...COMMON_FONTS,
  { file: 'NotoSansCJKjp-Regular.ttf', family: 'Noto Sans CJK JP', bold: false, italic: false, weight: 400, covers: isEastAsian },
  { file: 'NotoSerifCJKsc-Regular.ttf', family: 'Noto Serif CJK SC', bold: false, italic: false, weight: 400, covers: isEastAsian },
  // Google's COLRv1 release of Noto Color Emoji, unmodified: colour glyphs over TrueType outlines.
  { file: 'Noto-COLRv1.ttf', family: 'Noto Color Emoji', bold: false, italic: false, weight: 400, covers: coversEmoji },
  ...four('Liberation Sans Narrow', 'LiberationSansNarrow'),
  ...four('Morrow Mono', 'MorrowMono'),
  ...four('Onward', 'Onward'),
  ...four('Taper', 'Taper'),
  { file: 'Taper-Light.ttf', family: 'Taper Light', bold: false, italic: false, weight: 400 },
  { file: 'Taper-LightItalic.ttf', family: 'Taper Light', bold: false, italic: true, weight: 400 },
  { file: 'Gleis-Regular.ttf', family: 'Gleis', bold: false, italic: false, weight: 400 },
  { file: 'Gleis-Bold.ttf', family: 'Gleis', bold: true, italic: false, weight: 700 },
  { file: 'Gleis-Condensed.ttf', family: 'Gleis Condensed', bold: false, italic: false, weight: 400 },
  { file: 'Gleis-BoldCondensed.ttf', family: 'Gleis Condensed', bold: true, italic: false, weight: 700 },
  { file: 'Gleis-SemiCondensed.ttf', family: 'Gleis SemiCondensed', bold: false, italic: false, weight: 400 },
  { file: 'Gleis-BoldSemiCondensed.ttf', family: 'Gleis SemiCondensed', bold: true, italic: false, weight: 700 },
  { file: 'Anvil.ttf', family: 'Anvil', bold: false, italic: false, weight: 400 },
  // The GUST release, unmodified (CFF outlines, so PDFs embed it whole).
  { file: 'texgyreadventor-regular.otf', family: 'TeX Gyre Adventor', bold: false, italic: false, weight: 400 },
  { file: 'texgyreadventor-bold.otf', family: 'TeX Gyre Adventor', bold: true, italic: false, weight: 700 },
  { file: 'texgyreadventor-italic.otf', family: 'TeX Gyre Adventor', bold: false, italic: true, weight: 400 },
  { file: 'texgyreadventor-bolditalic.otf', family: 'TeX Gyre Adventor', bold: true, italic: true, weight: 700 },
  { file: 'ArchivoBlack-Regular.ttf', family: 'Archivo Black', bold: false, italic: false, weight: 400 },
  ...four('Comic Neue', 'ComicNeue'),
  ...four('Roboto', 'Roboto'),
  { file: 'Roboto-Medium.ttf', family: 'Roboto Medium', bold: false, italic: false, weight: 500 },
  { file: 'Roboto-MediumItalic.ttf', family: 'Roboto Medium', bold: false, italic: true, weight: 500 },
  { file: 'Roboto-Light.ttf', family: 'Roboto Light', bold: false, italic: false, weight: 300 },
  { file: 'Roboto-LightItalic.ttf', family: 'Roboto Light', bold: false, italic: true, weight: 300 },
  { file: 'DejaVuSansMono.ttf', family: 'DejaVu Sans Mono', bold: false, italic: false, weight: 400 },
  { file: 'DejaVuSansMono-Bold.ttf', family: 'DejaVu Sans Mono', bold: true, italic: false, weight: 700 },
  { file: 'DejaVuSansMono-Oblique.ttf', family: 'DejaVu Sans Mono', bold: false, italic: true, weight: 400 },
  { file: 'DejaVuSansMono-BoldOblique.ttf', family: 'DejaVu Sans Mono', bold: true, italic: true, weight: 700 },
];
