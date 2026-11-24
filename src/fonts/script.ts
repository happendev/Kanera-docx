import type { RunProps } from '../ooxml/model.js';

export function isEastAsian(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x2e80 && cp <= 0x2fdf) ||
    (cp >= 0x2ff0 && cp <= 0x303f) ||
    (cp >= 0x3040 && cp <= 0x9fff) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xac00 && cp <= 0xd7ff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe30 && cp <= 0xfe4f) ||
    (cp >= 0xff00 && cp <= 0xffef) ||
    (cp >= 0x20000 && cp <= 0x2fa1f)
  );
}

function isComplexScript(cp: number): boolean {
  return (cp >= 0x0590 && cp <= 0x08ff) || (cp >= 0x0900 && cp <= 0x0dff) || (cp >= 0x0e00 && cp <= 0x0eff) || (cp >= 0xfb1d && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff);
}

/** Characters that take the East Asian font when the run hints eastAsia (quotes, some symbols). */
function eastAsianAmbiguous(cp: number): boolean {
  return (cp >= 0x2018 && cp <= 0x201f) || cp === 0x2026 || cp === 0x2025 || (cp >= 0x2010 && cp <= 0x2015) || cp === 0xb7 || (cp >= 0x2460 && cp <= 0x24ff) || (cp >= 0x25a0 && cp <= 0x26ff);
}

interface FontChoice {
  family: string;
  bold: boolean;
  italic: boolean;
  size: number;
}

export function fontFor(cp: number, p: RunProps): FontChoice {
  if (p.cs || p.rtl || isComplexScript(cp)) return { family: p.fontCs, bold: p.boldCs, italic: p.italicCs, size: p.sizeCs };
  if (isEastAsian(cp) || (p.hint === 'eastAsia' && eastAsianAmbiguous(cp))) return { family: p.fontEastAsia, bold: p.bold, italic: p.italic, size: p.size };
  if (cp < 0x80) return { family: p.fontAscii, bold: p.bold, italic: p.italic, size: p.size };
  return { family: p.fontHAnsi, bold: p.bold, italic: p.italic, size: p.size };
}

