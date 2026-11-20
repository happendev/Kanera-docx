import { attr, child, type XmlElement } from '../util/xml.js';
import type { Border, Color } from './model.js';

export const twip = (v: number) => v / 20;
export const emu = (v: number) => v / 12700;

/** OOXML boolean (ST_OnOff): absent element = undefined, element without w:val = true. */
export function onOff(el: XmlElement | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = el.attrs['w:val'];
  if (v === undefined) return true;
  return !(v === '0' || v === 'false' || v === 'off' || v === 'none');
}

export function num(v: string | undefined): number | undefined {
  if (v === undefined || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Parse a measurement that may carry a unit (Word 2010+ allows "12pt", "1in", "2.5cm" in some
 * attributes). Bare numbers use `baseToPt` (twips by default).
 */
export function measure(v: string | undefined, baseToPt: (n: number) => number = twip): number | undefined {
  if (v === undefined || v === '') return undefined;
  const m = /^(-?(?:\d+(?:\.\d*)?|\.\d+))\s*(pt|in|cm|mm|pc|pi|px|emu)?$/.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  let points: number;
  switch (m[2]) {
    case 'pt': points = n; break;
    case 'in': points = n * 72; break;
    case 'cm': points = n * (72 / 2.54); break;
    case 'mm': points = n * (72 / 25.4); break;
    case 'pc':
    case 'pi': points = n * 12; break;
    case 'px': points = n * 0.75; break;
    case 'emu': points = n / 12700; break;
    default: points = baseToPt(n);
  }
  return Number.isFinite(points) ? points : undefined;
}

export function twipAttr(el: XmlElement | undefined, name: string): number | undefined {
  return measure(attr(el, name));
}

export function valAttr(el: XmlElement | undefined): string | undefined {
  return el?.attrs['w:val'];
}

export function childVal(parent: XmlElement | undefined, name: string): string | undefined {
  return valAttr(child(parent, name));
}

export function hexColor(v: string | undefined): Color | undefined {
  if (v === undefined) return undefined;
  if (v === 'auto' || v === '') return null;
  const s = v.replace(/^#/, '');
  if (/^[0-9a-fA-F]{6}$/.test(s)) return s.toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(s)) return (s[0] + s[0] + s[1] + s[1] + s[2] + s[2]).toUpperCase();
  return Object.hasOwn(NAMED_COLORS, v.toLowerCase()) ? NAMED_COLORS[v.toLowerCase()] : undefined;
}

const NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  lime: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  aqua: '00FFFF',
  magenta: 'FF00FF',
  fuchsia: 'FF00FF',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  maroon: '800000',
  olive: '808000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  orange: 'FFA500',
};

export const HIGHLIGHT_COLORS: Record<string, string> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  green: '00FF00',
  magenta: 'FF00FF',
  red: 'FF0000',
  yellow: 'FFFF00',
  white: 'FFFFFF',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  darkGray: '808080',
  lightGray: 'C0C0C0',
};

// ---------------------------------------------------------------------------------------------
// Colour maths (DrawingML colour transforms and Word theme tints)

export function rgb(hex: string): [number, number, number] {
  return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
}

export function toHex(r: number, g: number, b: number): string {
  const c = (v: number) =>
    Math.max(0, Math.min(255, Math.round(v)))
      .toString(16)
      .padStart(2, '0');
  return (c(r) + c(g) + c(b)).toUpperCase();
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/** Apply DrawingML colour modifiers (values in 1/1000 percent as in the XML). */
const toLinear = (c: number) => {
  const v = c / 255;
  return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
};
const fromLinear = (v: number) => 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * Math.max(0, v) ** (1 / 2.4) - 0.055);

export function applyColorMods(hex: string, mods: Array<{ name: string; val: number }>): { color: string; alpha: number } {
  let [r, g, b] = rgb(hex);
  let alpha = 1;
  for (const { name, val } of mods) {
    const f = val / 100000;
    switch (name) {
      case 'lumMod':
      case 'lumOff': {
        const [h, s, l] = rgbToHsl(r, g, b);
        const nl = name === 'lumMod' ? l * f : l + f;
        [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)));
        break;
      }
      // Saturation is not capped at 100%: the colour moves on out and clips at the RGB cube, so a fully
      // saturated 00B050 under satMod 115% is 00BD4F in Word's PDFs.
      case 'satMod': {
        const [h, s, l] = rgbToHsl(r, g, b);
        [r, g, b] = hslToRgb(h, Math.max(0, s * f), l).map((c) => Math.max(0, Math.min(255, c))) as [number, number, number];
        break;
      }
      // Tint moves towards white by (1 - f) and shade towards black, both in linear light: accent2
      // C0504D tinted 77% is D19392 in Word's PDFs, and black tinted 75% is 898989.
      case 'tint':
        [r, g, b] = [r, g, b].map((c) => fromLinear(toLinear(c) * f + (1 - f)));
        break;
      case 'shade':
        [r, g, b] = [r, g, b].map((c) => fromLinear(toLinear(c) * f));
        break;
      case 'alpha':
        alpha = f;
        break;
      case 'alphaMod':
        alpha *= f;
        break;
      case 'gray': {
        const y = 0.3 * r + 0.59 * g + 0.11 * b;
        r = g = b = y;
        break;
      }
      case 'inv':
        r = 255 - r;
        g = 255 - g;
        b = 255 - b;
        break;
    }
  }
  return { color: toHex(r, g, b), alpha };
}

/** Word's w:themeTint / w:themeShade (hex byte, 0..FF) applied to a theme colour. */
export function applyWordTintShade(hex: string, tint?: string, shade?: string): string {
  let [r, g, b] = rgb(hex);
  if (tint && /^[0-9a-f]{2}$/i.test(tint)) {
    const t = parseInt(tint, 16) / 255;
    const [h, s, l] = rgbToHsl(r, g, b);
    [r, g, b] = hslToRgb(h, s, l * t + (1 - t));
  }
  if (shade && /^[0-9a-f]{2}$/i.test(shade)) {
    const sh = parseInt(shade, 16) / 255;
    const [h, s, l] = rgbToHsl(r, g, b);
    [r, g, b] = hslToRgb(h, s, l * sh);
  }
  return toHex(r, g, b);
}

/** Blend a pattern foreground over a background at `pct` (0..1). */
export function blend(fg: string, bg: string, pct: number): string {
  const a = rgb(fg);
  const b = rgb(bg);
  return toHex(a[0] * pct + b[0] * (1 - pct), a[1] * pct + b[1] * (1 - pct), a[2] * pct + b[2] * (1 - pct));
}

/** Border width in points from w:sz (eighths of a point), with Word's limits. */
export function parseBorder(el: XmlElement | undefined, themeColor: (name: string) => string | undefined): Border | null | undefined {
  if (!el) return undefined;
  const style = el.attrs['w:val'] ?? 'single';
  if (style === 'nil' || style === 'none') return null;
  const sz = num(el.attrs['w:sz']) ?? 4;
  let width = Math.max(0.25, Math.min(sz / 8, 12));
  // A double border is two lines of the given width with as wide a gap between them.
  if (style === 'double') width *= 3;
  else if (style.startsWith('triple')) width = Math.max(width, 0.75);
  const space = num(el.attrs['w:space']) ?? 0;
  let color = hexColor(el.attrs['w:color']);
  const tc = el.attrs['w:themeColor'];
  if ((color === undefined || color === null) && tc) {
    const base = themeColor(tc);
    if (base) color = applyWordTintShade(base, el.attrs['w:themeTint'], el.attrs['w:themeShade']);
  }
  return { style, width, space, color: color ?? null };
}
