// Emoji sequences, as Excel and Word draw them on Windows: a character asked for as an emoji (by
// U+FE0F, a skin tone or a joiner) prints from the emoji font, and so does one that is an emoji by
// default (Unicode Emoji_Presentation) when the text's font lacks it. Other symbols, such as ☀
// without U+FE0F, keep the text's font or the symbol fallback (Segoe UI Symbol in Office).

const PRESENTATION = /^\p{Emoji_Presentation}$/u;
const PICTOGRAPHIC = /^\p{Extended_Pictographic}$/u;
const EMOJI = /^\p{Emoji}$/u;

const ZWJ = 0x200d;
const TEXT_STYLE = 0xfe0e;
const EMOJI_STYLE = 0xfe0f;
const KEYCAP = 0x20e3;

/** Variation selectors, which choose a form of the character before them and draw nothing. */
export function isVariationSelector(cp: number): boolean {
  return (cp >= 0xfe00 && cp <= 0xfe0f) || (cp >= 0xe0100 && cp <= 0xe01ef);
}

/** Tag characters, which spell out a subdivision flag after 🏴 and draw nothing. */
export function isTag(cp: number): boolean {
  return cp >= 0xe0020 && cp <= 0xe007f;
}

function isSkinTone(cp: number): boolean {
  return cp >= 0x1f3fb && cp <= 0x1f3ff;
}

function isRegionalIndicator(cp: number): boolean {
  return cp >= 0x1f1e6 && cp <= 0x1f1ff;
}

const test = (re: RegExp, cp: number) => re.test(String.fromCodePoint(cp));

/** Whether `cp` can start an emoji sequence: a pictograph or emoji, not an ASCII digit, # or *. */
function isEmojiBase(cp: number): boolean {
  return cp >= 0xa9 && (test(PICTOGRAPHIC, cp) || test(EMOJI, cp));
}

/**
 * Emoji fonts, best first: Office's, then the bundled colour font, then Noto's outlines. Fonts
 * with only bitmap glyphs (CBDT, sbix) are passed over; there is nothing to draw them with.
 */
export const EMOJI_FAMILIES = ['segoe ui emoji', 'noto color emoji', 'noto emoji'];

/**
 * How an emoji font draws in place of Segoe UI Emoji: its glyphs' size relative to the text's, how
 * far (em) below the baseline they are moved (which leaves the line's height alone), and its
 * advance relative to its own at the text's size.
 */
export interface EmojiStandIn {
  scale: number;
  drop: number;
  advance: number;
}

/**
 * Stand-ins for Segoe UI Emoji, whose emoji advance 1.373 em. Word for Mac prints Segoe's colour
 * glyphs about 0.98 em across with their middle 0.35 em above the baseline; Noto Color Emoji's are
 * about 8% larger and lower, and advance 1275/1024 em. Noto Emoji's outlines print about 8% larger
 * and a little lower than its own size, as Segoe UI Emoji's do in Excel's PDFs.
 */
export const EMOJI_STAND_INS: Readonly<Record<string, EmojiStandIn>> = {
  'noto color emoji': { scale: 0.92, drop: -0.03, advance: 1.373 / (1275 / 1024) },
  'noto emoji': { scale: 1.08, drop: 0.03, advance: 1.08 },
};

/** Whether an emoji font can draw `cp`: the bundled one's coverage test. */
export function coversEmoji(cp: number): boolean {
  return cp === KEYCAP || isEmojiBase(cp);
}

/** Whether `next`, after an emoji base, asks for an emoji: U+FE0F, a skin tone or a joiner. */
export function asksForEmoji(next: number | undefined): boolean {
  return next === EMOJI_STYLE || next === ZWJ || (next !== undefined && isSkinTone(next));
}

/**
 * Whether `cp`, in text made of `codePoints`, may be part of an emoji sequence printed from an emoji
 * font: a keycap mark, an emoji by default that the text's font lacks (`missing`), or a symbol in
 * text with a character that asks for an emoji.
 */
export function mayPrintAsEmoji(cp: number, codePoints: ReadonlySet<number>, missing: boolean): boolean {
  if (cp === KEYCAP) return true;
  if (!isEmojiBase(cp)) return false;
  if (codePoints.has(EMOJI_STYLE) || codePoints.has(ZWJ)) return true;
  for (let tone = 0x1f3fb; tone <= 0x1f3ff; tone++) if (codePoints.has(tone)) return true;
  return missing && test(PRESENTATION, cp);
}

/**
 * The number of code points in the emoji sequence that starts at `cps[i]` and prints from an emoji
 * font, or 0 when none does: there is no sequence, or it asks for the text style (U+FE0E) or is a
 * symbol that is text by default (☀, ❄) with nothing asking for an emoji.
 */
export function emojiSequence(cps: readonly number[], i: number): number {
  const base = cps[i];
  if (isRegionalIndicator(base)) return isRegionalIndicator(cps[i + 1]) ? 2 : 1;
  if (base < 0x80) {
    // A keycap: a digit, # or * with U+20E3, maybe with U+FE0F between them.
    if (!/[0-9#*]/.test(String.fromCharCode(base))) return 0;
    const j = cps[i + 1] === EMOJI_STYLE ? i + 2 : i + 1;
    return cps[j] === KEYCAP ? j + 1 - i : 0;
  }
  if (!isEmojiBase(base)) return 0;
  let j = i + 1;
  if (cps[j] === TEXT_STYLE) return 0;
  let emoji = test(PRESENTATION, base);
  const modifiers = () => {
    while (j < cps.length) {
      if (cps[j] === EMOJI_STYLE || isSkinTone(cps[j])) emoji = true;
      else if (!isTag(cps[j]) && cps[j] !== KEYCAP) break;
      j++;
    }
  };
  modifiers();
  while (cps[j] === ZWJ && j + 1 < cps.length && isEmojiBase(cps[j + 1])) {
    emoji = true;
    j += 2;
    modifiers();
  }
  return emoji ? j - i : 0;
}
