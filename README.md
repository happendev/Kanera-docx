# Kanera DOCX to PDF

`@happendev/docx2pdf` converts Word documents (`.docx`) to PDF, aiming to put every line, page break,
image and border where Word puts it. The TypeScript renderer uses Node’s built-in zlib for compression.
The browser and runtime-neutral core use pako, bundled with the package. Conversion runs without
LibreOffice, native add-ons or WebAssembly.

The library is used in [Kanera](https://github.com/happendev/Kanera).

```ts
import { docxToPdf } from '@happendev/docx2pdf';

const pdf = await docxToPdf('contract.docx'); // Uint8Array
```

## How it gets close to Word

Matching Word means reproducing its layout rules, not only its formatting:

- **Fonts that measure like Office fonts.** When a document uses Calibri, Cambria, Arial, Times New
  Roman, Courier New, Georgia, Segoe UI or Aptos and the font isn't available, the library
  uses a bundled open font whose glyph widths measure the same (see [Fonts](#fonts)), so lines
  wrap at the same words. Fonts embedded in the `.docx` are used first.
- **Word's line model.** Line heights use the Windows (GDI) font metrics Word uses; "multiple" line
  spacing adds its extra space below the text, and that space may hang into the bottom margin;
  spacing before a paragraph is dropped at the top of a page only after a natural page break.
  Superscripts and subscripts are sized in whole half-points as Word does and don't change the line
  height. Spaces and tabs never make a line taller (a line of nothing else takes its paragraph
  mark's height). A list bullet raises its line only by what its ascent exceeds the text's ascent
  and line gap, never deeper, and a missing Symbol or Wingdings font keeps its own line metrics.
  Automatic (HTML) paragraph spacing is dropped at the top of the document, a header, a cell or a
  text box, and at the bottom of a cell.
- **Word's line breaking.** Justified lines in Word 2013+ documents, a paragraph's last line
  included, shrink their spaces by up to a quarter to keep one more word, as Word does, but only
  when that changes the spaces less than half as much as stretching them without the word would
  (fitted to about 3,500 line breaks in Word's PDFs); older documents only stretch, and
  WordPerfect-style justification compresses the whole line slightly. Ordinary spaces hang past
  the right margin, while en, em and other fixed-width spaces wrap like text.
- **Word's pagination.** Widow/orphan control (on by default, as in Word), keep-with-next chains,
  keep-lines-together, page-break-before, table rows that can't split, repeated header rows,
  footnotes reserved at the page bottom, header and footer heights pushing the body down.
- **Word's table geometry.** Border widths add to row heights, Word 2013+ vs. older table
  indentation (a package without a settings part counts as older), vertically merged cells, exact
  and minimum row heights, cell spacing. A floating object laid out in its cell positions from the
  cell and makes the row tall enough to hold it. A nested table stays inside its cell's text area,
  shrinking to fit, and the empty paragraph that closes a cell after it takes no room.
- **Word's text wrap.** Text wraps around every floating object on the page, including one
  positioned against the page or margin that belongs to a paragraph further down, and a first line
  pushed aside by an object keeps its first-line indent from the object's edge.
- **Word's East Asian layout.** A document grid snaps lines to its pitch and centres the text in
  the grid lines it takes, and a character grid widens every character. With punctuation
  compression on, full-width punctuation gives up part of its blank half to keep a character on an
  overfull line. East Asian text sits a quarter em from Latin letters and digits (autospacing),
  and justification widens that space first. A missing SimSun, MS Mincho, PMingLiU, Batang or
  other common East Asian font keeps its own line height on the Noto fallback, and an empty theme
  East Asian font takes the theme's font for the document's language.

Accuracy is measured against real documents: against PDFs Word exported from them, word by word,
and against where Word last broke pages (`w:lastRenderedPageBreak`) and how many pages it rendered
(`docProps/app.xml`). `pnpm corpus` runs the second comparison for a folder of `.docx` files (see
[Development](#development)).

## Benchmark

The [benchmark](benchmarks/word/README.md) now runs 208 cases, including
[30 pinned Hugging Face cases](benchmarks/word/fixtures.json) and
[178 vendored cases](benchmarks/word/upstream/fixtures.json).

| Converter | Median cold time | Median warm time | Median cold peak RSS | Max cold peak RSS | Matched Word's page count | Accuracy (mean ink overlap) | Total PDF size |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `@happendev/docx2pdf` 0.1.0 | 79.0 ms | 9.1 ms | 74.1 MiB | 137.2 MiB | 29/30 | 65.7% | 1.93 MiB |
| EigenPal 2.39.0 | 554.6 ms | 77.6 ms | 236.0 MiB | 528.8 MiB | 28/30 | 74.9% | 1.99 MiB |
| LibreOffice 26.8 alpha | 247.8 ms | 29.6 ms | 184.0 MiB | 237.0 MiB | 26/30 | 66.8% | 2.42 MiB |
| Aspose.Words 26.11.0 | 962.4 ms | 50.4 ms | 101.0 MiB | 143.4 MiB | 26/30 | 55.2% | 1.62 MiB |

Cold timings include a fresh process launch and file I/O. Warm timings reuse one process per
engine after an unmeasured conversion, with file I/O still included. GNU time records peak
resident memory (RSS) for the cold runs. The JavaScript engines ran on Node 24 and licensed
Aspose.Words ran on .NET 10. EigenPal used best-effort mode. Kanera DOCX used the full font pack.

PDF sizes sum the cold-run files for each converter.

For fidelity, we rasterized each PDF at 96 DPI and compared corresponding pages. Ink overlap
combines how much of Word's ink the converter reproduced with how much of the converter's ink
matched Word, allowing a one-pixel shift. Unmatched pages score zero. We averaged page scores
within each document, then gave each document equal weight. The metric measures placement and
visible content, including missing images, and ignores colour differences. The [benchmark
instructions](benchmarks/word/README.md) explain how to reproduce the results and generate the
side-by-side screenshot gallery.

## Install

Install `@happendev/docx2pdf` from the GitHub repository with npm (installing from git with pnpm 11
can hang):

```bash
npm install github:happendev/kanera-docx
```

Node 20 or later. In the browser, any engine with ES2022, `fetch` and `TextDecoder`.

The package includes the full font pack. Node, the CLI and the browser load individual faces on
demand.

## Use

### Node

```ts
import { convertDocx, docxToPdf } from '@happendev/docx2pdf'; // or '@happendev/docx2pdf/node'

// A path, Uint8Array, Buffer or ArrayBuffer.
const pdf = await docxToPdf(await readFile('in.docx'));

// The full result:
const result = await convertDocx('in.docx', {
  systemFonts: true, // also use fonts installed on this machine (default false)
  fonts: [{ data: await readFile('BrandSans.ttf') }], // extra fonts, highest priority
  info: { title: 'Contract' }, // override PDF metadata
});
result.pdf; // Uint8Array
result.pageCount;
result.warnings; // e.g. 'Font "Proxima Nova" not available; using Liberation Sans'
```

`systemFonts` is off by default so the output doesn't depend on which machine runs the
conversion. Pass `true` to scan the usual font folders, or an array of folders to scan.

### Browser

```ts
import { docxToPdf } from '@happendev/docx2pdf'; // resolves to the browser build

const file = input.files[0]; // File, Blob, Uint8Array or ArrayBuffer
const pdf = await docxToPdf(file, { fontBaseUrl: '/assets/docx2pdf-fonts/' });
const url = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }));
```

The full pack is available by default. The browser fetches individual faces on demand and caches
them for the page's lifetime. With a bundler, copy the installed package's `fonts/` directory
into your static assets and point `fontBaseUrl` at it. When serving the package as-is, the default
`fontBaseUrl` resolves to that directory.

Pass `fetch` to add caching or auth. Set `fontBaseUrl: false` to use only embedded fonts and
faces you supply through `fonts`.

#### Common font pack in the browser

For a smaller static asset set, copy only the faces listed in `commonFontFiles`, exported from
`@happendev/docx2pdf/browser`, along with `fonts/licenses/`. Pass
`{ fontPack: 'common', fontBaseUrl: '/assets/docx2pdf-common-fonts/' }`.
You can supply additional faces through `fonts`. If no available face covers a character, the PDF
shows a missing glyph and `convertDocx` reports it in `warnings`.

### Core (any runtime)

`@happendev/docx2pdf/core` has no font loading of its own: give it every font through `fonts`.

```ts
import { convertDocx } from '@happendev/docx2pdf/core';

const { pdf } = await convertDocx(bytes, { fonts: [{ data: regular }, { data: bold }] });
```

### CLI

```bash
docx2pdf report.docx               # writes report.pdf next to it
docx2pdf report.docx out.pdf --font BrandSans.ttf --system-fonts
docx2pdf report.docx out.pdf --image-quality 70
```

## Options

| Option          | Where   | Default  | Meaning                                                              |
| --------------- | ------- | -------- | -------------------------------------------------------------------- |
| `fonts`         | all     | `[]`     | Fonts to use (`{ data }` or `{ family, bold, italic, load }`). Win over bundled and system fonts. |
| `archiveLimits` | all | 64 MiB per part, 256 MiB total | Maximum uncompressed bytes per ZIP part (`maxPartBytes`) and across distinct parts read (`maxTotalBytes`). Exceeding either limit stops conversion. |
| `imageLimits` | all | 16,777,216 pixels per image, 67,108,864 total | Set `maxImagePixels` and `maxTotalPixels` to budget source raster pixels. Repeated placements of one image count once. |
| `imageQuality` | all | automatic 80 | Encoder quality for eligible colour JPEGs. Set an integer from 1 to 100 or use `false` to disable lossy re-encoding. CLI: `--image-quality <1-100\|off>`. |
| `additionalLinkSchemes` | all | `[]` | PDF links support HTTP, HTTPS, mailto and internal bookmarks. Pass scheme names such as `['tel']` to enable additional URI actions. Text for other destinations is preserved without a link annotation. |
| `embeddedFonts` | all     | `true`   | Use fonts embedded in the `.docx` (including obfuscated `.odttf`).   |
| `info`          | all     | from doc | PDF title, author, subject, keywords, dates.                         |
| `compress`      | all     | `true`   | Deflate PDF content streams.                                         |
| `wordPlatform`  | all     | `'windows'` | Set to `'mac'` to match Word for Mac's PDFs, which draw text sizes, baselines and line starts on a 0.24pt grid (11pt text comes out at 11.04pt). Layout is the same. CLI: `--mac`. |
| `systemFonts`   | Node    | `false`  | `true` or a list of folders to scan for `.ttf`/`.otf`/`.ttc`.        |
| `fontPack`      | browser | `'full'` | Set to `'common'` to use the smaller browser pack.                      |
| `fontDir`       | Node    | package  | Folder holding the full pack.                                        |
| `fontBaseUrl`   | browser | module   | URL of the selected pack, or `false` to disable automatic loading.   |
| `fetch`         | browser | global   | Custom fetch for font downloads.                                     |

Font priority goes from `fonts` to embedded fonts, system fonts, substitutes and generic
fallbacks. A font passed in `fonts` under a family name replaces every other font of that name,
whatever its own name table says.

Automatic image quality uses a quality-80 JPEG only when it saves at least 10% and the decoded
pixel differences pass a fidelity check. An explicit number uses that quality when the resulting
stream is smaller. Quality is an encoder setting. Savings vary by image. PNG and supported TIFF
use lossless PDF compression of their decoded pixels. TIFF is decoded directly into a PDF image
stream. As with other conversions, existing decoder limits apply.

## Converting untrusted documents

Conversion stops when an image exceeds the pixel budget or either dimension exceeds 32,768 pixels.
TIFF edge-tile padding counts toward the budget. Each conversion also allows up to 256 font faces,
1,048,576 character mappings, 128 MiB of embedded font references, 256 columns per table and 64
columns per section. Set `embeddedFonts: false` to skip reading embedded font binaries while
keeping the font names and substitution hints.

For a hosted converter, run each job in a separate process with an OS-enforced memory limit.
Have the parent enforce a wall-clock deadline and terminate the child when it expires. Conversion
includes synchronous work, so a timer inside the converter can be delayed by the job itself.

**Cropping in Word hides pixels that can still be recovered from the PDF.** The renderer clips the
image on the page. JPEGs kept in their original encoding can include photo metadata. Before sharing a PDF,
permanently crop sensitive pixels and strip metadata from the source image, then replace the
picture in the document.

## Fonts

The full pack includes all common faces. The table shows which browser pack each face first
becomes available in.

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
| Avenir Next LT Pro                       | Onward                           | yes               | Full   |
| Avenir, Avenir Next                      | Onward                           | approximate       | Full   |
| Menlo, Monaco, SF Mono                    | DejaVu Sans Mono                 | approximate       | Full   |
| Arial Black                              | Archivo Black                    | approximate       | Full   |
| Comic Sans MS, Chalkboard, Chalkboard SE  | Comic Neue                       | approximate       | Full   |
| Roboto (Light, Medium)                    | Roboto                           | same font         | Full   |
| Symbol, Wingdings 1–3, Webdings            | mapped to Unicode                | n/a               | Both   |

"Same glyph widths" means the substitute is designed to match the original's advance widths for
shared characters. Coverage, kerning and line metrics can still affect layout. Caladea, Gelasio and
Liberation Sans Narrow use the original fonts' line metrics. A character the substitute lacks can
come from a font with different widths. Morrow Sans, Morrow Sans Display, Morrow Sans Narrow and
Morrow Sans Light match the line metrics of their Aptos counterparts. [Morrow Sans font log](https://github.com/open-metric-fonts/open-metric-fonts/blob/d5d565c5056491a20480674e9b1d95394188cfa8/fonts/morrow-sans/FONTLOG.txt)

The [substitution rules](src/fonts/substitutes.ts) list more families, including optional compatible
fonts such as Fidela for Constantia. Supply those through
`fonts` or `systemFonts`. When the converter uses an approximate substitute, it reports the change
in `warnings` because glyph shapes, line wrapping or spacing can differ. For the original glyphs,
use an embeddable copy of the original font.

Characters in Symbol, Wingdings, Wingdings 2 and 3 and Webdings are mapped to the Unicode
characters Unicode 7 encoded them as, and copy as those. When no available font has one, an older
look-alike stands in (U+27A2 for Wingdings' ➢ arrowhead bullet); otherwise, and for the few codes
with no Unicode equivalent, the symbol draws as a missing glyph and `warnings` names it. Pass the
symbol font itself through `fonts` to draw the original shapes.

East Asian text whose font is missing falls back by the font's region and style (SimSun to Noto
Serif CJK SC, MS Mincho to Noto Serif CJK JP, Malgun Gothic to Noto Sans CJK KR), then by the text's
East Asian language (`w:lang`). The full font pack provides Noto Sans CJK JP in all three runtimes.
Add regional Noto CJK fonts through `fonts` or `systemFonts` for other glyph forms.

Fonts are subset into the PDF (TrueType, keeping glyph IDs) with a `ToUnicode` map, so text can be
selected, searched and copied. CFF-based OpenType fonts and fonts that prohibit subsetting are embedded whole. Fonts that prohibit outline embedding are skipped with a warning. Missing bold or italic
faces are synthesised the way Word does.

The bundled Caladea is the original 1.002 release, whose widths match Cambria; later Caladea
releases changed the widths. It lacks the Romanian comma-below letters (Ș ș Ț ț), some combining
accents and superscript four, which come from a fallback font.

The common browser font files total about 16 MiB and include the Aptos Narrow and Aptos Light
substitutes. The full pack is about 49 MiB. A regular Arial document fetches only Liberation Sans
Regular. The browser can fetch Noto Sans CJK JP (20.7 MiB) from the full pack for
missing East Asian glyphs. It uses Japanese glyph forms. The pack supplies a static regular face
generated at weight 400 from the [upstream variable TrueType font](https://github.com/notofonts/noto-cjk/blob/f8d157532fbfaeda587e826d4cd5b21a49186f7c/Sans/Variable/TTF/NotoSansCJKjp-VF.ttf).

Font licenses are in [`fonts/licenses`](fonts/licenses). The full pack includes Liberation Sans
Narrow from Debian 1.07.6-4 under GPL 2.0 with the font embedding exception. Its corresponding source is in
[`fonts/sources/liberation-sans-narrow`](fonts/sources/liberation-sans-narrow).
Anvil, Gleis, Morrow Mono, Onward and Taper come from the same revision as Morrow Sans.
[Open Metric Fonts](https://github.com/open-metric-fonts/open-metric-fonts/tree/d5d565c5056491a20480674e9b1d95394188cfa8)

Archivo Black and Comic Neue are from Google Fonts. DejaVu Sans Mono is from DejaVu 2.37.
[Archivo Black](https://github.com/google/fonts/tree/94a7d81318e438525a5285e07ab72c050fdfeb44/ofl/archivoblack) · [Comic Neue](https://github.com/google/fonts/tree/9b83905ab4e7a6034297ff95598d37644fb7d378/ofl/comicneue)

## What is supported

- Paragraphs: alignment and justification, indents, hanging indents, tabs (left, centre, right,
  decimal, bar, leaders), spacing and line rules, borders and shading, contextual spacing,
  keep/widow rules, page and column breaks, soft hyphens.
- Runs: fonts by script, size, bold, italic, underline styles, strike, double strike, colour,
  highlight, shading, caps and small caps, superscript/subscript, position, character spacing and
  scaling, kerning (kern and GPOS pairs), hidden text, outline.
- Styles: document defaults, paragraph, character, table and linked styles, theme fonts and
  colours.
- Numbering: bullets and multi-level lists in every Word number format, restarts, legal numbering,
  picture bullets, and Symbol/Wingdings bullet characters.
- Tables: grid and preferred widths (fixed, auto, percent), merged cells both ways, cell and table
  borders, conditional table-style formatting (header row, banding, first column, ...),
  cell margins, vertical alignment, nested tables, rows split across pages, repeated header rows,
  row-level property exceptions, empty cells that don't count toward the row height (`w:hideMark`),
  cell spacing (`w:tblCellSpacing`), and floating tables (`w:tblpPr`) with text wrapping beside
  them.
- Sections: page size and orientation, margins, gutters, columns, different first/even headers and
  footers, header and footer inheritance, page number formats and restarts, page borders,
  continuous, even- and odd-page section breaks (an empty paragraph that only carries a section
  break before a continuous section takes no room, as in Word).
- Frames (`w:framePr`) in headers, footers and text boxes positioned against the margin and the
  paragraph that follows, such as Word's page number frames, with text wrapping around them.
- Fields: PAGE, NUMPAGES, SECTIONPAGES, external hyperlinks and internal links to bookmarks (as
  clickable PDF links), legacy form checkboxes and drop-downs; other fields (TOC, dates,
  references) show the result Word last saved.
- Footnotes and endnotes with separators, including footnotes continued across pages.
- Plain-text and HTML/XHTML imports (`w:altChunk`). HTML is rendered as plain text. The warnings identify flattened HTML and unsupported import types.
- Images: PNG (all colour types, transparency, interlacing), JPEG (including CMYK), GIF, BMP and
  TIFF (uncompressed, LZW, Deflate and PackBits). Basic SVG, EMF and WMF are drawn as vector paths.
  Images can be inline or floating, with cropping, rotation, flips and square/tight/top-and-bottom
  text wrap, also around pictures in headers and footers.
- Shapes and text boxes (DrawingML and VML): preset geometries, custom paths, fills, outlines, dash
  styles, transparency, groups, rotated and vertical text, WordArt text paths and watermarks.
- Tracked changes render as the final document (insertions shown, deletions hidden).
- Right-to-left paragraphs, headers and table text, with bidirectional reordering and Arabic
  letter joining.

## Limitations

- SVGs with features such as gradients, filters, stylesheets or text use Word's raster preview when
  available. Supported SVGs become PDF vector drawing commands, with metadata and unused
  definitions omitted. An unsupported SVG without a preview, WebP, or CCITT- or JPEG-compressed
  TIFF produces a warning. Metafiles drawn only in EMF+ records come out blank.
- Charts and SmartArt appear only through the fallback picture Word saves with them.
- Equations render as linear text.
- Bidirectional text follows a simplified Unicode algorithm without explicit embeddings, and Arabic
  joins through presentation forms; no Indic shaping, no OpenType ligatures, and right-to-left
  tables (`w:bidiVisual`) are not mirrored.
- A floating table taller than the space left on its page is laid out inline (Word splits it
  across pages and keeps it floating).
- Text in continuous-section columns is not balanced; no automatic hyphenation; frames in the body
  and drop caps flow as ordinary paragraphs; no line numbers; comments are not shown.
- Pagination can differ from Word when a document's fonts are missing and have no metric-compatible
  substitute.

## Development

```bash
pnpm install
pnpm test          # unit and end-to-end tests (uses Ghostscript, when installed, to read the PDFs back)
pnpm typecheck
pnpm build         # dist/

pnpm render in.docx --out out/ --png       # convert, and rasterise with Ghostscript
pnpm corpus list.txt --root ~/docs         # compare page counts and breaks with Word for many files
```

`list.txt` holds one `.docx` path per line (relative to `--root`). Set `BREAKS=1` to print where
each of Word's recorded page breaks landed. Keep private documents out of this repository.

Run the [full corpus accuracy check](benchmarks/word/README.md#full-corpus-accuracy) to measure ink
overlap and compare changes between runs across 2,925 Hugging Face cases and 178 vendored cases.

Code map: `src/docx` reads the package into a document model (styles, numbering, drawing),
`src/fonts` parses, substitutes and subsets fonts, `src/layout` breaks lines and paginates into
positioned drawing operations, and `src/pdf` writes those as PDF. `src/util` holds the ZIP reader,
compression adapters and XML parser.

## License

[Apache License 2.0](LICENSE). The bundled fonts keep their own licenses (see [Fonts](#fonts)).
