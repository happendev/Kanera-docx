# Kanera DOCX to PDF

`@happendev/docx2pdf` converts Word documents (`.docx`) to PDF, aiming to put every line, page break,
image and border where Word puts it. It is plain TypeScript with **no runtime dependencies**: no
LibreOffice, no native binaries, no WebAssembly. The same code runs in Node and in the browser.

```ts
import { docxToPdf } from '@happendev/docx2pdf';

const pdf = await docxToPdf('contract.docx'); // Uint8Array
```

## How it gets close to Word

Matching Word means reproducing its layout rules, not only its formatting:

- **Fonts that measure like Office fonts.** When a document uses Calibri, Cambria, Arial, Times New
  Roman, Courier New, Georgia, Segoe UI or Roboto and the font isn't available, the library uses
  a bundled open font with the same character widths and line metrics (see [Fonts](#fonts)), so
  lines wrap at the same words. Fonts embedded in the `.docx` are used first.
- **Word's line model.** Line heights use the Windows (GDI) font metrics Word uses; "multiple" line
  spacing adds its extra space below the text, and that space may hang into the bottom margin;
  spacing before a paragraph is dropped at the top of a page only after a natural page break.
  Superscripts and subscripts are sized in whole half-points as Word does and don't change the line
  height. A list bullet can make its line taller but never deeper, and a missing Symbol or
  Wingdings font keeps its own line metrics.
- **Word's line breaking.** Justified lines in Word 2013+ documents, a paragraph's last line
  included, shrink their spaces by up to a quarter to keep one more word, as Word does; older
  documents only stretch.
- **Word's pagination.** Widow/orphan control (on by default, as in Word), keep-with-next chains,
  keep-lines-together, page-break-before, table rows that can't split, repeated header rows,
  footnotes reserved at the page bottom, header and footer heights pushing the body down.
- **Word's table geometry.** Border widths add to row heights, Word 2013+ vs. older table
  indentation, vertically merged cells, exact and minimum row heights.

Accuracy is measured against real documents: against PDFs Word exported from them, word by word,
and against where Word last broke pages (`w:lastRenderedPageBreak`) and how many pages it rendered
(`docProps/app.xml`). `pnpm corpus` runs the second comparison for a folder of `.docx` files (see
[Development](#development)).

## Install

Install `@happendev/docx2pdf` from the GitHub repository with npm (installing from git with pnpm 11
can hang):

```bash
npm install github:happendev/kanera-docx
```

Node 20 or later. In the browser, any engine with ES2022, `fetch` and `TextDecoder`.

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
result.warnings; // e.g. 'Font "Aptos" not available; using Selawik'
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

The bundled fonts are fetched on demand, only for the families a document uses, and cached for the
page's lifetime. Copy `node_modules/@happendev/docx2pdf/fonts/` to your static assets and point
`fontBaseUrl` at that folder. Without a bundler, the default (`fonts/` next to the module) works
as-is. Pass `fetch` to add caching or auth, or `fontBaseUrl: false` to supply every font yourself
through `fonts`.

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
```

## Options

| Option          | Where   | Default  | Meaning                                                              |
| --------------- | ------- | -------- | -------------------------------------------------------------------- |
| `fonts`         | all     | `[]`     | Fonts to use (`{ data }` or `{ family, bold, italic, load }`). Win over bundled and system fonts. |
| `archiveLimits` | all | 64 MiB per part, 256 MiB total | Maximum uncompressed bytes per ZIP part (`maxPartBytes`) and across distinct parts read (`maxTotalBytes`). Exceeding either limit stops conversion. |
| `embeddedFonts` | all     | `true`   | Use fonts embedded in the `.docx` (including obfuscated `.odttf`).   |
| `info`          | all     | from doc | PDF title, author, subject, keywords, dates.                         |
| `compress`      | all     | `true`   | Deflate PDF content streams.                                         |
| `systemFonts`   | Node    | `false`  | `true` or a list of folders to scan for `.ttf`/`.otf`/`.ttc`.        |
| `fontDir`       | Node    | package  | Folder holding the bundled fonts.                                    |
| `fontBaseUrl`   | browser | module   | URL of the bundled fonts folder, or `false`.                         |
| `fetch`         | browser | global   | Custom fetch for font downloads.                                     |

Font priority: `fonts` → embedded → system → bundled substitutes → generic fallback (Liberation
Sans, DejaVu Sans and Noto Sans CJK JP according to glyph coverage).

## Fonts

| Document font                    | Bundled substitute   | Same metrics |
| -------------------------------- | -------------------- | ------------ |
| Calibri                          | Carlito              | yes          |
| Cambria                          | Caladea              | yes          |
| Arial, Helvetica                 | Liberation Sans      | yes          |
| Times New Roman                  | Liberation Serif     | yes          |
| Courier New                      | Liberation Mono      | yes          |
| Georgia                          | Gelasio              | yes          |
| Segoe UI (Light, Semibold)       | Selawik              | yes          |
| Roboto (Light, Medium)           | Roboto               | same font    |
| Symbol, Wingdings bullets        | mapped to Unicode    | n/a          |
| Anything else                    | closest generic      | no, warned   |

Fonts without a metric-compatible substitute (Calibri Light, Aptos, Verdana, Tahoma, Gill Sans,
Montserrat, ...) still render, but lines may wrap differently from Word, and the result says so in
`warnings`. Add the real font through `fonts` or `systemFonts` (where you are licensed to) to get
an exact match.

Fonts are subset into the PDF (TrueType, keeping glyph IDs) with a `ToUnicode` map, so text can be
selected, searched and copied. CFF-based OpenType fonts and fonts that prohibit subsetting are embedded whole. Fonts that prohibit outline embedding are skipped with a warning. Missing bold or italic
faces are synthesised the way Word does.

The bundled fonts are under the SIL Open Font License or the Liberation/DejaVu licenses; the
license texts are in [`fonts/licenses`](fonts/licenses). The bundled files total about 34 MiB. A regular Arial document fetches only
Liberation Sans Regular. When a document needs missing East Asian glyphs, the browser can also
fetch Noto Sans CJK JP (20.7 MiB), which uses Japanese glyph forms. The bundled copy is a
static regular face generated at weight 400 from the [upstream variable TrueType font](https://github.com/notofonts/noto-cjk/blob/f8d157532fbfaeda587e826d4cd5b21a49186f7c/Sans/Variable/TTF/NotoSansCJKjp-VF.ttf).

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
  and floating tables (`w:tblpPr`) with text wrapping beside them.
- Sections: page size and orientation, margins, gutters, columns, different first/even headers and
  footers, header and footer inheritance, page number formats and restarts, page borders,
  continuous, even- and odd-page section breaks.
- Fields: PAGE, NUMPAGES, SECTIONPAGES, external hyperlinks and internal links to bookmarks (as
  clickable PDF links), legacy form checkboxes; other fields (TOC, dates, references) show the
  result Word last saved.
- Footnotes and endnotes with separators, including footnotes continued across pages.
- Plain-text and HTML/XHTML imports (`w:altChunk`). HTML is rendered as plain text. The warnings identify flattened HTML and unsupported import types.
- Images: PNG (all colour types, transparency, interlacing), JPEG (including CMYK), GIF and BMP,
  inline or floating, with cropping, rotation, flips and square/tight/top-and-bottom text wrap.
- Shapes and text boxes (DrawingML and VML): preset geometries, custom paths, fills, outlines, dash
  styles, transparency, groups, rotated and vertical text, WordArt text paths and watermarks.
- Tracked changes render as the final document (insertions shown, deletions hidden).

## Limitations

- EMF, WMF, TIFF and SVG images are not drawn (a warning is added). Word stores a PNG fallback for
  SVG images, which is used when present.
- Charts and SmartArt appear only through the fallback picture Word saves with them.
- Equations render as linear text.
- No complex-script shaping (Arabic, Hebrew, Indic) or bidirectional reordering; no ligatures.
- A floating table taller than the space left on its page is laid out inline (Word splits it
  across pages and keeps it floating).
- Text in continuous-section columns is not balanced; no automatic hyphenation; frames
  (`w:framePr`) and drop caps flow as ordinary paragraphs; no line numbers; comments are not shown.
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

Code map: `src/docx` reads the package into a document model (styles, numbering, drawing),
`src/fonts` parses, substitutes and subsets fonts, `src/layout` breaks lines and paginates into
positioned drawing operations, and `src/pdf` writes those as PDF. `src/util` holds the ZIP reader,
inflate/deflate and XML parser.

## License

[Apache License 2.0](LICENSE). The bundled fonts keep their own licenses (see [Fonts](#fonts)).
