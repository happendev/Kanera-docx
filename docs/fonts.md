# Fonts

Office documents name fonts that are rarely free to redistribute. The package bundles open fonts
with the same glyph widths, so lines wrap at the same words. Fonts embedded in the `.docx` are used
first. Pass your own fonts in `fonts`, or set `systemFonts` in Node to use the machine's; see
[Options](options.md) for the font priority.

## Substitutes

The full pack includes every face below; the last column shows the smallest browser pack that has it.

| Document font                            | Substitute                       | Same glyph widths | Browser pack |
| ---------------------------------------- | -------------------------------- | ----------------- | ------------ |
| Calibri                                  | Carlito                          | yes               | Common |
| Cambria                                  | Caladea                          | yes               | Common |
| Arial, Helvetica                         | Liberation Sans                  | yes               | Common |
| Times New Roman                          | Liberation Serif                 | yes               | Common |
| Courier New                              | Liberation Mono                  | yes               | Common |
| Georgia                                  | Gelasio                          | yes               | Common |
| Segoe UI (Light, Semibold)                | Selawik                          | yes               | Common |
| Segoe UI Variable                        | Selawik                          | approximate       | Common |
| Aptos, Aptos Display                     | Morrow Sans, Morrow Sans Display | yes               | Common |
| Aptos Narrow                             | Morrow Sans Narrow               | yes               | Common |
| Aptos Light                              | Morrow Sans Light                | yes               | Common |
| Arial Narrow                             | Liberation Sans Narrow           | yes               | Full   |
| Aptos Mono                               | Morrow Mono                      | yes               | Full   |
| Bahnschrift (Regular, Bold and widths)    | Gleis                            | yes               | Full   |
| Candara (including Light)                | Taper                            | yes               | Full   |
| Impact                                   | Anvil                            | yes               | Full   |
| Century Gothic                           | TeX Gyre Adventor                | yes               | Full   |
| Avenir Next LT Pro                       | Onward                           | yes               | Full   |
| Avenir, Avenir Next                      | Onward                           | approximate       | Full   |
| Menlo, Monaco, SF Mono                    | DejaVu Sans Mono                 | approximate       | Full   |
| Arial Black                              | Archivo Black                    | approximate       | Full   |
| Comic Sans MS, Chalkboard, Chalkboard SE  | Comic Neue                       | approximate       | Full   |
| Roboto (Light, Medium)                    | Roboto                           | same font         | Full   |
| Segoe UI Emoji                           | Noto Color Emoji                 | no                | Full   |
| Symbol, Wingdings 1–3, Webdings            | mapped to Unicode                | n/a               | Both   |

"Same glyph widths" means the substitute matches the original's advance widths for shared
characters; kerning, coverage and line metrics can still differ. Caladea, Gelasio, Liberation Sans
Narrow, TeX Gyre Adventor and the Morrow Sans family also match the originals' line metrics. More
families, including optional ones such as Fidela for Constantia, are in the [substitution
rules](../src/fonts/substitutes.ts). An approximate substitute is reported in `warnings`; for the
original glyphs, supply the original font.

## Symbols, East Asian text, emoji and embedding

- **Symbol fonts.** Symbol, Wingdings 1–3 and Webdings characters map to Unicode, falling back to an
  older look-alike or a missing glyph (named in `warnings`). Pass the symbol font through `fonts`
  for the original shapes.
- **East Asian.** A missing font falls back by style, then region, then the text's language
  (SimSun to Noto Serif CJK SC, MS Mincho to Noto Serif CJK JP). Kai faces fall back to a sans, as
  in Word for Mac. The full pack includes Noto Sans CJK JP (20.7 MiB) and Noto Serif CJK SC
  (23.0 MiB). Node loads their layout tables first and reads the glyph outlines used in the PDF
  asynchronously. Documents with equations use the full fonts to measure glyph bounds. Add other
  regional Noto CJK fonts through `fonts` or `systemFonts`.
- **Emoji.** Emoji print in colour, as in Word for Mac, from Segoe UI Emoji's colour layers. This
  applies to a character that is an emoji by default and missing from the run's font, and to any
  symbol followed by U+FE0F, a skin tone or a joiner. Skin tones, families, flags and other
  sequences print as one glyph where the font has one. Segoe UI Emoji is used when supplied in
  `fonts` or found with `systemFonts`; otherwise the bundled Noto Color Emoji stands in, drawn at
  about the size of Word's Segoe glyphs and spaced to Segoe's widths. A line with emoji is as tall
  as its text, as in Word for Mac. The colour glyphs are drawn over invisible text, so emoji can
  still be copied and searched. Any font with COLR colour glyphs and TrueType outlines prints in
  colour; fonts with only bitmap emoji (CBDT, sbix) are passed over for emoji.
- **Embedding.** TrueType fonts are subset, with a `ToUnicode` map so text can be copied and
  searched. CFF fonts and fonts that forbid subsetting are embedded whole; fonts that forbid
  embedding are skipped with a warning. Missing bold and italic are synthesised as Word does.
- **Sizes.** The common browser pack is about 16 MiB and the full pack about 77 MiB.

## Caladea

The bundled Caladea is the 1.002 release, whose widths match Cambria (later releases changed
them); it lacks Ș ș Ț ț, some combining accents and superscript four, which come from a fallback.

## Licences and sources

**Licenses** are in [`fonts/licenses`](../fonts/licenses). Sources:
[Open Metric Fonts](https://github.com/open-metric-fonts/open-metric-fonts/tree/d5d565c5056491a20480674e9b1d95394188cfa8)
(Morrow Sans, Morrow Mono, Anvil, Gleis, Onward, Taper;
[font log](https://github.com/open-metric-fonts/open-metric-fonts/blob/d5d565c5056491a20480674e9b1d95394188cfa8/fonts/morrow-sans/FONTLOG.txt)) ·
[Archivo Black](https://github.com/google/fonts/tree/94a7d81318e438525a5285e07ab72c050fdfeb44/ofl/archivoblack) ·
[Comic Neue](https://github.com/google/fonts/tree/9b83905ab4e7a6034297ff95598d37644fb7d378/ofl/comicneue) ·
DejaVu Sans Mono 2.37 · TeX Gyre Adventor 2.501 (GUST Font License, unmodified) ·
Noto CJK static weight-400 instances of the variable
[Sans](https://github.com/notofonts/noto-cjk/blob/f8d157532fbfaeda587e826d4cd5b21a49186f7c/Sans/Variable/TTF/NotoSansCJKjp-VF.ttf) and
[Serif](https://github.com/notofonts/noto-cjk/blob/f8d157532fbfaeda587e826d4cd5b21a49186f7c/Serif/Variable/TTF/NotoSerifCJKsc-VF.ttf)
fonts, the Serif reduced to the glyphs and features the converter uses. Liberation Sans Narrow is
from Debian 1.07.6-4 under GPL 2.0 with the font embedding exception; its source is in
[`fonts/sources/liberation-sans-narrow`](../fonts/sources/liberation-sans-narrow). Noto Color Emoji
is the 2.057 COLRv1 build of
[googlefonts/noto-emoji](https://github.com/googlefonts/noto-emoji/tree/fc4ca365e7c20e78278ae702aa20434bfe704c8f)
as packaged by Fedora (`google-noto-color-emoji-fonts`), unmodified, as `Noto-COLRv1.ttf`.
