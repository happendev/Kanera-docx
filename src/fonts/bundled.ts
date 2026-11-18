/**
 * Fonts shipped in the package's `fonts/` directory. Families are listed so a loader can fetch
 * only what a document needs; the family names must match each file's name table.
 */
export interface BundledFont {
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

export const BUNDLED_FONTS: BundledFont[] = [
  ...four('Carlito', 'Carlito'),
  ...four('Caladea', 'Caladea'),
  ...four('Liberation Sans', 'LiberationSans'),
  ...four('Liberation Serif', 'LiberationSerif'),
  ...four('Liberation Mono', 'LiberationMono'),
  ...four('Gelasio', 'Gelasio'),
  ...four('Roboto', 'Roboto'),
  { file: 'Roboto-Medium.ttf', family: 'Roboto Medium', bold: false, italic: false, weight: 500 },
  { file: 'Roboto-MediumItalic.ttf', family: 'Roboto Medium', bold: false, italic: true, weight: 500 },
  { file: 'Roboto-Light.ttf', family: 'Roboto Light', bold: false, italic: false, weight: 300 },
  { file: 'Roboto-LightItalic.ttf', family: 'Roboto Light', bold: false, italic: true, weight: 300 },
  { file: 'Selawik-Regular.ttf', family: 'Selawik', bold: false, italic: false, weight: 400 },
  { file: 'Selawik-Bold.ttf', family: 'Selawik', bold: true, italic: false, weight: 700 },
  { file: 'Selawik-Light.ttf', family: 'Selawik Light', bold: false, italic: false, weight: 300 },
  { file: 'Selawik-Semibold.ttf', family: 'Selawik Semibold', bold: false, italic: false, weight: 600 },
  { file: 'DejaVuSans.ttf', family: 'DejaVu Sans', bold: false, italic: false, weight: 400 },
  { file: 'DejaVuSans-Bold.ttf', family: 'DejaVu Sans', bold: true, italic: false, weight: 700 },
];
