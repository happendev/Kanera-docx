/**
 * The settings in which kanera-docx, kanera-xlsx and kanera-pptx differ. The shared core files
 * (`pnpm core:status` checks them) are identical in all three packages and read these; this file
 * is each package's own.
 */

/**
 * The Office application whose PDFs this package's output is measured against. Word's (measured on
 * Word for Mac's PDFs) and Excel's and PowerPoint's (on Windows) differ for some substitute fonts'
 * line metrics and widths: see `winVmetrics` and `winWidths` in src/fonts/substitutes.ts.
 */
export const APP: 'word' | 'excel' | 'powerpoint' = 'word';

/**
 * The bundled emoji font: Noto Color Emoji's colour glyphs ('color'; Word for Mac prints Segoe UI
 * Emoji in colour), or Noto Emoji's outlines ('outline').
 */
export const EMOJI_FONT: 'color' | 'outline' = 'color';

/**
 * The font the application draws text in when the computer lacks the font it asks for: a macOS
 * font such as Avenir Next on Windows, or one with no substitute rule. PowerPoint on Windows uses
 * Calibri ('calibri'). Undefined keeps the closest substitute of the missing font instead.
 */
export const MISSING_FONT: string | undefined = undefined;

/**
 * Whether a font named as PostScript and web fonts name it, its words run together with a weight
 * after them (AvenirNextLTPro-Demi), finds its family's substitute rule. PowerPoint's PDFs draw such
 * names in the family's font; Word's draw them in its fallback font, as it draws any font it lacks.
 */
export const FONT_NAME_VARIANTS: boolean = false;

/**
 * The width of the outline that makes synthetic bold text bolder, in em: Word's 0.03, PowerPoint's
 * 1/35 (its PDFs stroke a 28pt run 0.8pt wide).
 */
export const SYNTHETIC_BOLD_STROKE: number = 0.03;

/**
 * How a blurred shape shadow (a:outerShdw with a blur radius) is drawn: a picture at `dpi`, left out
 * under the shape itself when `clip`, and taking in the shape's outline as well as its fill when
 * `outline`. Word's PDFs (Word for Mac) draw it at 300 dpi, under the whole shape and its outline.
 */
export const SOFT_SHADOW: { dpi: number; clip: boolean; outline: boolean } = { dpi: 300, clip: false, outline: true };

/**
 * How text sits in a DrawingML shape. With `outlineInset` the text's insets count from inside the
 * shape's outline, so text keeps clear of half the outline's width on every side and a box that
 * fits its text grows by the outline's width. With `overflowFromTop` text too tall for a centred or
 * bottom-anchored box starts at the box's top and runs out of its bottom. Word's PDFs show both.
 */
export const SHAPE_TEXT: { outlineInset: boolean; overflowFromTop: boolean } = { outlineInset: true, overflowFromTop: true };

/**
 * Whether solid level and upright lines (borders, rules, underlines) are written to the PDF as filled
 * rectangles rather than strokes, as Word's PDFs draw them. Viewers widen a stroke thinner than a
 * pixel to a whole pixel but shade a thin fill by its coverage, so stroked 0.5pt borders look darker
 * and heavier on screen than Word's.
 */
export const LINE_FILLS: boolean = true;
