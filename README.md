# Kanera DOCX to PDF

[![License: Apache 2.0](https://img.shields.io/badge/license-Apache%202.0-blue)](LICENSE)

`@kanera/docx2pdf` converts Word documents (`.docx`) to PDF, aiming to put every line, page break,
image and border where Word puts it. The TypeScript renderer uses Node’s built-in zlib for compression.
The browser and runtime-neutral core use pako, bundled with the package. Conversion runs without
LibreOffice, native add-ons or WebAssembly.

Kanera uses this renderer in production. The library also extracts plain text and Markdown from
Word documents.

[Release notes](CHANGELOG.md) · [Report a conversion issue](https://github.com/happendev/Kanera-docx/issues/new?template=conversion-bug.yml)

<!-- markdown-summary:start -->
Kanera scored **98.5% Markdown accuracy** across 278 DOCX documents. [Markdown benchmark](#markdown-benchmark)
<!-- markdown-summary:end -->

<!-- benchmark-summary:start -->
### Benchmark at a glance

278 DOCX documents [Full comparison](#benchmark)

| Converter | Visual match | Page counts matched | Fresh process | Reused process | Median peak memory |
| --- | ---: | ---: | ---: | ---: | ---: |
| Kanera | **92.6%** | **273/278** | **94.4 ms** | **6.4 ms** | 75.6 MiB |
| LibreOffice | 80.3% | 256/278 | 807.2 ms | 58.0 ms | – |
| EigenPal | 70.9% | 257/278 | 775.1 ms | 257.8 ms | 278.6 MiB |
| Syncfusion DocIO (.NET) | 56.1% | 246/278 | 568.8 ms | 60.0 ms | 121.8 MiB |
| Aspose.Words (.NET) | 52.1% | 250/278 | 1018.2 ms | 36.2 ms | 105.9 MiB |
| office2pdf (Rust) | 49.1% | 240/278 | 321.4 ms | 107.5 ms | **42.5 MiB** |
| MiniPdf (.NET) | 43.8% | 247/278 | 344.4 ms | 162.9 ms | 137.2 MiB |
| rdocx (Rust) | 40.5% | 214/278 | 179.5 ms | 24.0 ms | 146.2 MiB |
| dxpdf (Rust) | 36.4% | 236/278 | 1024.3 ms | 949.4 ms | 258.3 MiB |

Visual match measures how closely text and graphics line up with the reference pages. Conversion times are medians and include file I/O.
<!-- benchmark-summary:end -->

## Quick start

### Install

```bash
npm install @kanera/docx2pdf
```

Use Node.js 20 or later. Browser conversion requires ES2022, `fetch` and `TextDecoder`.

### Node

```ts
import { writeFile } from 'node:fs/promises';
import { docxToPdf } from '@kanera/docx2pdf'; // or '@kanera/docx2pdf/node'

const pdf = await docxToPdf('in.docx'); // a path, Uint8Array, Buffer or ArrayBuffer
await writeFile('out.pdf', pdf);
```

### CLI

Run the CLI directly from npm. This writes `in.pdf` next to `in.docx`.

```bash
npx @kanera/docx2pdf in.docx
```

In a project with the package installed, use `npx --no-install docx2pdf`. Pass an output path or
options as needed.

```bash
npx --no-install docx2pdf in.docx out.pdf --font BrandSans.ttf --system-fonts
```

To get page images instead of a PDF, pass `--png` or `--jpeg` with a directory. This writes
`pages/in-1.png`, `pages/in-2.png` and so on, with page numbers zero-padded to the same width so
they sort in order. `--dpi` sets the resolution (default 150), `--pages` picks pages such as `1-3,5`,
and `--quality` sets the JPEG quality (default 90).

```bash
npx --no-install docx2pdf in.docx --png pages --dpi 200
```

To extract the text instead, pass `--text` or `--markdown` (`--md`). This writes `in.txt` or
`in.md`, or prints to standard output when the output is `-`. Add `--headers` to include headers
and footers, or `--no-notes` to leave out footnotes and endnotes.

```bash
npx --no-install docx2pdf in.docx --markdown -
```

### Page count and warnings

Use `convertDocx` to get conversion warnings along with the PDF and page count.

```ts
import { convertDocx } from '@kanera/docx2pdf';

const { pdf, pageCount, warnings } = await convertDocx('in.docx');
```

Tracked changes render as final text, with insertions shown and deletions hidden. The PDF omits
Word comments.

### Text and Markdown

`docxToText` and `docxToMarkdown` return the document's text without rendering it. They take the
same inputs as `docxToPdf`, load no fonts and decode no pictures.

```ts
import { docxToMarkdown, docxToText, extractText } from '@kanera/docx2pdf';

const text = await docxToText('in.docx');
const markdown = await docxToMarkdown('in.docx', { headersFooters: true });
const { text: md, warnings } = await extractText('in.docx', { format: 'markdown' });
```

The [Markdown benchmark](#markdown-benchmark) compares the Markdown with seven other converters.

### Page images

`docxToImages` draws pages as PNG or JPEG images from the same layout as the PDF, without
rendering a PDF or calling an outside rasteriser. `convertDocxToImages` also returns the page count
and warnings. Both take the layout and font options of `convertDocx`, and run in Node and the browser.

```ts
import { writeFile } from 'node:fs/promises';
import { docxToImages } from '@kanera/docx2pdf';

for (const { page, data } of await docxToImages('in.docx', { dpi: 150 })) {
  await writeFile(`page-${page}.png`, data);
}
const [cover] = await docxToImages('in.docx', { format: 'jpeg', quality: 85, pages: [1] });
```

Each image has its `page` number, `width`, `height` and encoded `data`. Scored against Word's
pages the way the [benchmark](#benchmark) scores PDFs, the images match as closely as the PDF does.

## Use

### Node options

`convertDocx` also accepts [options](#options) for fonts and PDF metadata.

```ts
import { readFile } from 'node:fs/promises';
import { convertDocx } from '@kanera/docx2pdf';

const result = await convertDocx('in.docx', {
  systemFonts: true, // also use installed fonts (off by default, for reproducible output)
  fonts: [{ data: await readFile('BrandSans.ttf') }], // extra fonts, highest priority
  info: { title: 'Contract' }, // override PDF metadata
});
result.pdf; // Uint8Array
result.pageCount;
result.warnings; // e.g. 'Font "Proxima Nova" not available; using Liberation Sans'
```

The package is ES modules only; CommonJS can `require()` it on Node 20.19+ and 22.12+, or use
`await import()`. If you bundle your server, copy the package's `fonts/` folder into the build
output and pass its path as `fontDir`.

#### In a web server

```ts
import { createServer } from 'node:http';
import { convertDocx } from '@kanera/docx2pdf';

const MAX_UPLOAD = 20 * 1024 * 1024;

createServer(async (req, res) => {
  if (req.method !== 'POST') return void res.writeHead(405).end();
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_UPLOAD) return void res.writeHead(413).end();
    chunks.push(chunk);
  }
  try {
    const { pdf } = await convertDocx(Buffer.concat(chunks));
    res.writeHead(200, { 'content-type': 'application/pdf' }).end(pdf);
  } catch (err) {
    res.writeHead(422, { 'content-type': 'text/plain' }).end(`Could not convert: ${(err as Error).message}`);
  }
}).listen(3000);
```

Conversion is synchronous, so a large document blocks other requests. For documents from users,
see [Converting untrusted documents](#converting-untrusted-documents).

### Browser

Conversion runs in the browser. To try the demo locally, run `pnpm demo:serve` from a checkout
with dependencies installed.
The demo previews and downloads PDFs, plain text (`.txt`), and Markdown (`.md`). In PDF mode, you can compare the included benchmark documents with Word's output. [Demo source](demo/)

For a bundled app that serves `public/` at the site root, copy the bundled fonts into that folder
after installing the package.

```bash
node -e "require('node:fs').cpSync('node_modules/@kanera/docx2pdf/fonts', 'public/docx2pdf-fonts', { recursive: true })"
```

Pass that directory's public URL as `fontBaseUrl` and supply a `File` from a file input or drop
handler. The browser fetches individual font faces on demand.

<!-- browser-conversion-example -->
```ts
import { convertDocx } from '@kanera/docx2pdf'; // resolves to the browser build

const { pdf } = await convertDocx(file, { fontBaseUrl: '/docx2pdf-fonts/' });
const pdfBlob = new Blob([pdf], { type: 'application/pdf' });
```

If you serve the package as-is, the default `fontBaseUrl` resolves to its bundled `fonts/` directory.

A plain Arial document downloads one face (about 400 KB); CJK text can fetch a Noto CJK face of
over 20 MB. Pass `fetch` to add caching or auth, or `fontBaseUrl: false` to use only embedded and
supplied fonts. For a smaller asset set, serve only the faces in `commonFontFiles` (exported from
`@kanera/docx2pdf/browser`) plus `fonts/licenses/`, and pass `fontPack: 'common'`.

Conversion runs on the calling thread. Use a Web Worker to keep the page responsive, and reuse the
worker across conversions so downloaded font faces stay cached.

```ts
// convert.worker.ts
import { convertDocx } from '@kanera/docx2pdf';

self.onmessage = async (e: MessageEvent<File>) => {
  try {
    self.postMessage(await convertDocx(e.data, { fontBaseUrl: '/docx2pdf-fonts/' }));
  } catch (err) {
    self.postMessage({ error: String(err) });
  }
};
```

```ts
// main.ts
const worker = new Worker(new URL('./convert.worker.ts', import.meta.url), { type: 'module' });
worker.onmessage = (e) => {
  if (e.data.error) return console.error(e.data.error);
  window.open(URL.createObjectURL(new Blob([e.data.pdf], { type: 'application/pdf' })));
};
worker.postMessage(file);
```

### Core (any runtime)

`@kanera/docx2pdf/core` loads no fonts itself; pass every font through `fonts`.

```ts
import { convertDocx } from '@kanera/docx2pdf/core';

const { pdf } = await convertDocx(bytes, { fonts: [{ data: regular }, { data: bold }] });
```

## Options

| Option          | Where   | Default  | Meaning                                                              |
| --------------- | ------- | -------- | -------------------------------------------------------------------- |
| `fonts`         | all     | `[]`     | Fonts to use (`{ data }` or `{ family, bold, italic, load }`). Win over bundled and system fonts. |
| `archiveLimits` | all | 64 MiB per part, 256 MiB total | Maximum uncompressed bytes per ZIP part (`maxPartBytes`) and across distinct parts read (`maxTotalBytes`). Exceeding either limit stops conversion. |
| `imageLimits` | all | 16,777,216 pixels per image, 67,108,864 total | Set `maxImagePixels` and `maxTotalPixels` to budget source raster pixels. Repeated placements of one image count once. JPEGs are only decoded for optional recompression, so one that exceeds the budget is embedded unchanged. |
| `imageQuality` | all | | Encoder quality for eligible colour JPEGs. Set an integer from 1 to 100 or use `false` to disable lossy re-encoding. CLI: `--image-quality <1-100\|off>`. |
| `additionalLinkSchemes` | all | `[]` | PDF links support HTTP, HTTPS, mailto and internal bookmarks. Pass scheme names such as `['tel']` to enable additional URI actions. Text for other destinations is preserved without a link annotation. |
| `embeddedFonts` | all     | `true`   | Use fonts embedded in the `.docx` (including obfuscated `.odttf`).   |
| `info`          | all     | from doc | PDF title, author, subject, keywords, dates.                         |
| `compress`      | all     | `true`   | Deflate PDF content streams.                                         |
| `wordPlatform`  | all     | `'windows'` | Set to `'mac'` to match Word for Mac's PDFs, which draw text sizes, baselines and line starts on a 0.24pt grid (11pt text comes out at 11.04pt). Layout is the same. CLI: `--mac`. |
| `format`        | images  | `'png'`  | `'png'` or `'jpeg'`.                                                 |
| `dpi`           | images  | `150`    | Image resolution in dots per inch. CLI: `--dpi`.                     |
| `quality`       | images  | `90`     | JPEG quality from 1 to 100. CLI: `--quality`.                        |
| `pages`         | images  | all      | Page numbers to draw, from 1. CLI: `--pages 1-3,5`.                  |
| `systemFonts`   | Node    | `false`  | `true` or a list of folders to scan for `.ttf`/`.otf`/`.ttc`.        |
| `fontPack`      | browser | `'full'` | Set to `'common'` to use the smaller browser pack.                      |
| `fontDir`       | Node    | package  | Folder holding the full pack.                                        |
| `fontBaseUrl`   | browser | module   | URL of the selected pack, or `false` to disable automatic loading.   |
| `fetch`         | browser | global   | Custom fetch for font downloads.                                     |

Font priority goes from `fonts` to embedded fonts, system fonts, substitutes and generic
fallbacks. A font passed in `fonts` under a family name replaces every other font of that name,
whatever its own name table says.

With `imageQuality` omitted, eligible colour JPEGs of up to 4,000,000 pixels are encoded at quality 80
when the 64 entries in their 8-bit quantization table 0 sum to 750 or less. Other JPEGs keep their
original stream. An explicit quality encodes each eligible JPEG within the configured image budget
once, even above 4 megapixels or when the encoded file grows. Recompressed JPEGs are baseline with
4:2:0 chroma; a JPEG of more than 32 scans keeps its original stream. Set `imageQuality: false` to
preserve all original JPEG streams.
PNG and supported TIFF
use lossless PDF compression of their decoded pixels. TIFF is decoded directly into a PDF image
stream, except that JPEG-compressed TIFFs are embedded as JPEGs (strips joined, not recompressed,
and following the JPEG quality rules above) and CCITT fax TIFFs of one strip keep their fax data.
As with other conversions, existing decoder limits apply.

### Text extraction

The text is what Word shows, in reading order. Tracked insertions are kept and deletions dropped,
hidden text and field codes are left out, and fields show the result Word last saved. Lists keep
Word's numbers and bullets, with Symbol and Wingdings characters mapped to Unicode. Text boxes,
shapes and SmartArt follow the paragraph that anchors them, and a dropped capital rejoins its
word. Charts give their title and the data cached in the document. Comments are left out.

Plain text puts each paragraph on its own line and each table row on a line with cells separated
by tabs. Equations are written in linear form, such as `x^2+a/(b+c)`. Note references show as
`[1]` and the notes follow the body.

Markdown follows GitHub's dialect:

- Heading styles, the Title style and outline levels become `#` headings.
- Lists are nested by level. Bullets become `-` and numbers such as `1.` and `2)` are kept.
  Labels that Markdown can't number, such as `a)`, `IV.` and `1.2.3`, are kept as text after a `-`.
- Bold, italic, strikethrough, superscript and subscript are preserved. Runs in a monospaced font
  become inline code. Paragraphs entirely in one become a code block. A document set entirely in
  Courier stays prose.
- External hyperlinks become links. Pictures become `![alt text](word/media/image1.png)`, where
  the path is the picture's part in the package.
- Tables become pipe tables. A single-cell table is unwrapped, and a table with nested tables is
  written as HTML.
- Footnotes and endnotes become Markdown footnotes. Equations become LaTeX in `$…$` or `$$` blocks.
- Quote styles become block quotes, and table-of-contents entries become a nested list.

| Option           | Default  | Meaning                                                                 |
| ---------------- | -------- | ----------------------------------------------------------------------- |
| `format`         | `'text'` | `'markdown'` for Markdown (`extractText` only).                          |
| `headersFooters` | `false`  | Include each distinct header before its section and each footer after it. |
| `notes`          | `true`   | Include footnotes and endnotes after the body.                           |
| `images`         | `true`   | Markdown: include picture references.                                    |
| `archiveLimits`  |          | As for conversion.                                                       |

## Converting untrusted documents

Conversion stops when an image exceeds the pixel budget or either dimension exceeds 32,768 pixels.
TIFF edge-tile padding counts toward the budget. Each conversion also allows up to 256 font faces,
1,048,576 character mappings, 128 MiB of embedded font references, 256 columns per table and 64
columns per section. XML parts support 128 levels of elements and layout allows 10,000 pages.
OpenType single-substitution and positioning features share a budget of 1,048,576 work units per conversion.
Set `embeddedFonts: false` to skip reading embedded font binaries while keeping the font names and substitution hints.

Charts allow 32,768 slots per data cache and 262,144 across a conversion. Empty positions in
sparse chart data count toward both limits.

For a hosted converter, run each job in a separate process with an OS-enforced memory limit.
Have the parent enforce a wall-clock deadline and terminate the child when it expires. Conversion
includes synchronous work, so a timer inside the converter can be delayed by the job itself.

**Cropping in Word hides pixels that can still be recovered from the PDF.** The renderer clips the
image on the page. JPEGs kept in their original encoding can include photo metadata. Before sharing a PDF,
permanently crop sensitive pixels and strip metadata from the source image, then replace the
picture in the document.

## Fonts

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
rules](src/fonts/substitutes.ts). An approximate substitute is reported in `warnings`; for the
original glyphs, supply the original font.

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

The bundled Caladea is the 1.002 release, whose widths match Cambria (later releases changed
them); it lacks Ș ș Ț ț, some combining accents and superscript four, which come from a fallback.

**Licenses** are in [`fonts/licenses`](fonts/licenses). Sources:
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
[`fonts/sources/liberation-sans-narrow`](fonts/sources/liberation-sans-narrow). Noto Color Emoji
is the 2.057 COLRv1 build of
[googlefonts/noto-emoji](https://github.com/googlefonts/noto-emoji/tree/fc4ca365e7c20e78278ae702aa20434bfe704c8f)
as packaged by Fedora (`google-noto-color-emoji-fonts`), unmodified, as `Noto-COLRv1.ttf`.

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
- Footnotes and endnotes with separators, including footnotes continued across pages and footnotes
  kept on the page of the table row that refers to them.
- Plain-text and HTML/XHTML imports (`w:altChunk`). HTML is rendered as plain text. The warnings identify flattened HTML and unsupported import types.
- Images: PNG (all colour types, transparency, interlacing), JPEG (including CMYK), GIF, BMP and
  TIFF (uncompressed, LZW, Deflate, PackBits, CCITT fax and JPEG). Basic SVG, EMF and WMF are drawn as vector paths.
  Images can be inline or floating, with cropping, rotation, flips and square/tight/top-and-bottom
  text wrap, also around pictures in headers and footers. Brightness, contrast and Washout apply to
  raster pictures, except CMYK JPEGs; SVG and metafile pictures are drawn unadjusted.
- Shapes and text boxes (DrawingML and VML): preset geometries, custom paths, fills (including
  linear and path gradients, blended as Word blends them), outlines, dash styles, transparency,
  groups, rotated and vertical text, text boxes that resize to fit their text, WordArt text paths
  and watermarks.
- Charts: column, bar, line, area, scatter, pie, doughnut and radar charts (clustered, stacked and
  100%, combinations and secondary axes) are drawn as vector graphics from the values cached in the
  document, with Excel's automatic axis scaling, number formats, titles, legends, data labels,
  trendlines, error bars, gradient fills and shadows. SmartArt is drawn from the shapes Word saves
  with the diagram, text included.
- Equations (OMML) laid out as Word's equation editor sets them: fractions, sub- and superscripts,
  radicals, integrals and other large operators with their limits, stretching brackets, matrices,
  equation arrays aligned at their marks, accents, bars and grouping characters, in Cambria Math's
  proportions.
- Vertical East Asian text: sections (`w:textDirection` `tbRl`) and text boxes set top to bottom
  with lines running right to left, ideographs and kana upright in the font's vertical forms;
  headers and footers stay horizontal, as in Word.
- Tracked changes render as the final document (insertions shown, deletions hidden).
- Right-to-left paragraphs, headers and tables (`w:bidiVisual`, first column on the right), with
  bidirectional reordering and Arabic, Persian and Urdu letter joining. Arabic drawn by a fallback
  font for Arial, Times New Roman or Dubai takes that font's width and line height, so lines break
  where Word's do.

## Limitations

- SVGs with features such as gradients, filters, stylesheets or text use Word's raster preview when
  available. Supported SVGs become PDF vector drawing commands, with metadata and unused
  definitions omitted. An unsupported SVG without a preview, WebP, or old-style JPEG TIFF
  without a complete JPEG stream produces a warning. Metafiles drawn only in EMF+ records come out blank.
- 3D charts are drawn as their flat versions. Surface charts are left blank, and newer chart types
  (waterfall, treemap, histogram and others) show the fallback Word saves with them. No document in
  the benchmark or test corpora has a radar chart or error bars, so those follow Excel's documented
  behaviour without a Word reference to check them against.
- Equations are drawn with Caladea standing in for Cambria Math, spaced to Cambria Math's widths.
- Bidirectional text follows a simplified Unicode algorithm without explicit embeddings, and Arabic
  joins through presentation forms; no Indic shaping and no OpenType ligatures. No Arabic font is
  bundled: pass one in `fonts` (such as Noto Naskh Arabic or Noto Sans Arabic) for a closer look
  than the DejaVu Sans fallback.
- No font is bundled for Thai, Myanmar, Khmer, Tibetan, Thaana, Ethiopic or the Indic scripts
  (Devanagari, Bengali, Tamil and others). Supply one in `fonts` or use `systemFonts`; otherwise
  that text prints as missing glyphs.
- Without Segoe UI Emoji, emoji print in Noto Color Emoji's designs, which differ from Segoe UI
  Emoji's, and a country flag prints as a flag where Segoe UI Emoji has two letters. Colour fonts
  print sweep gradients in one colour and repeating gradients as padded ones; colour glyphs with
  CFF outlines, SVG or bitmaps print without colour or not at all.
- Text in vertical sections does not wrap around floating objects.
- A floating table taller than the space left on its page is laid out inline (Word splits it
  across pages and keeps it floating).
- Text in continuous-section columns is not balanced; no automatic hyphenation; frames in the body
  and drop caps flow as ordinary paragraphs; no line numbers; comments are not shown.
- Pagination can differ from Word when a document's fonts are missing and have no metric-compatible
  substitute.

## How it gets close to Word

Matching Word means reproducing its layout rules, not only its formatting:

- **Metric-compatible fonts.** Missing Office fonts (Calibri, Cambria, Arial, Times New Roman,
  Aptos and others) are replaced by bundled open fonts with the same glyph widths, so lines wrap at
  the same words. Fonts embedded in the `.docx` are used first. See [Fonts](#fonts).
- **Line model.** Line heights use Word's Windows (GDI) font metrics, including its rules for
  "multiple" spacing, superscripts, bullets, and spacing at page, cell and text-box tops.
- **Line breaking.** Justified lines shrink or stretch their spaces the way Word 2013+ does, fitted
  to about 3,500 line breaks in Word's PDFs; older documents follow the older rules.
- **Pagination.** Widow/orphan control, keep-with-next and keep-lines-together, unsplittable rows,
  repeated header rows, footnotes and header/footer heights.
- **Tables and text wrap.** Word's border, indentation and row-height geometry, floating objects in
  cells, nested tables, and text wrapping around every floating object on the page.
- **East Asian layout.** Document and character grids, punctuation compression, autospacing between
  East Asian and Latin text, and metric-preserving fallbacks for missing CJK fonts, including the
  proportional kana and punctuation of fonts such as MS PGothic.
- **Word for Mac PDFs.** Set `wordPlatform: 'mac'` or pass `--mac` on the CLI to snap text sizes,
  baselines and line starts to Word for Mac’s 0.24 pt grid during PDF rendering.

Accuracy is checked against PDFs Word exported, and against the page breaks and page counts Word
recorded in each file (`pnpm corpus`, see [Development](#development)).

## Benchmark

The [benchmark](benchmarks/word/README.md) compares converters on 278 documents drawn from
[100 Hugging Face cases](benchmarks/word/fixtures.json) with Word for Mac PDFs and
[178 vendored cases](benchmarks/word/upstream/fixtures.json) with Microsoft 365 Word and LibreOffice references.

<!-- benchmark-results:start -->
| Converter | Mean ink overlap | Matching page count | Failed conversions | Median cold time | Median warm time | Warm p95 latency | Warm docs/sec | Median cold peak RSS | Max cold peak RSS | Total PDF size |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf 1.0.0 | **92.6%** | **273/278** | **0/278** | **94.4 ms** | **6.4 ms** | **62.9 ms** | **66.0** | 75.6 MiB | **160.8 MiB** | **17.43 MiB** |
| LibreOffice 26.8.3.1 | 80.3% | 256/278 | **0/278** | 807.2 ms | 58.0 ms | 244.1 ms | 10.5 | – | – | 44.46 MiB |
| EigenPal 2.39.0 | 70.9% | 257/278 | 1/278 | 775.1 ms | 257.8 ms | 661.0 ms | 3.6 | 278.6 MiB | 578.4 MiB | 21.98 MiB |
| Syncfusion DocIO (.NET) 28.2.4.0 | 56.1% | 246/278 | 1/278 | 568.8 ms | 60.0 ms | 2507.9 ms | 1.2 | 121.8 MiB | 1128.7 MiB | 18.91 MiB |
| Aspose.Words (.NET) 26.11.0.0 | 52.1% | 250/278 | **0/278** | 1018.2 ms | 36.2 ms | 356.7 ms | 11.3 | 105.9 MiB | 243.2 MiB | 19.84 MiB |
| office2pdf (Rust) 0.11.0 | 49.1% | 240/278 | 1/278 | 321.4 ms | 107.5 ms | 242.6 ms | 8.1 | **42.5 MiB** | 250.0 MiB | 30.30 MiB |
| MiniPdf (.NET) 0.43.2 | 43.8% | 247/278 | **0/278** | 344.4 ms | 162.9 ms | 193.8 ms | 6.3 | 137.2 MiB | 211.8 MiB | 2532.58 MiB |
| rdocx (Rust) 0.17.0 | 40.5% | 214/278 | 44/278 | 179.5 ms | 24.0 ms | 227.9 ms | 14.8 | 146.2 MiB | 466.3 MiB | 24.12 MiB |
| dxpdf (Rust) 0.8.1 | 36.4% | 236/278 | 9/278 | 1024.3 ms | 949.4 ms | 1271.7 ms | 1.0 | 258.3 MiB | 341.4 MiB | 15.80 MiB |

Peak RSS is the peak resident memory during a successful cold conversion. LibreOffice memory is unavailable because GNU time measured its launcher.

The smallest total PDF size is determined among converters that completed every document.
<!-- benchmark-results:end -->

Accuracy is measured as mean ink overlap. The benchmark compares pages at 96 DPI, allows a
one-pixel shift, and gives each document equal weight. Missing pages and failed conversions score
zero. See the [benchmark README](benchmarks/word/README.md) for the breakdown by reference source,
measurement details, and reproduction steps.

<!-- full-corpus-accuracy:start -->
**Full-corpus ink test.**
Kanera scored **90.49%** across 2,650 documents
(2,472 Hugging Face and 178 vendored, 10,376 reference pages).
Page counts matched on 2,554 documents. 1 conversion failed.
[Full corpus accuracy](benchmarks/word/README.md#full-corpus-accuracy).
<!-- full-corpus-accuracy:end -->

### Markdown benchmark

The [Markdown benchmark](benchmarks/markdown/README.md) converts the same 278 documents to Markdown
with seven other converters. It renders each converter's Markdown with one CommonMark renderer.
Words, reading order, bold and italic are scored against Word's own PDF. Headings, lists, tables,
links and pictures are scored against the document's XML.

<!-- markdown-benchmark:start -->
| Converter | Markdown accuracy | Text | Order | Headings | Lists | Tables | Bold | Italic | Links | Images | Leaked syntax | Failed |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf 1.0.0 | **98.5%** | **98.8%** | **97.6%** | **98.5%** | **99.0%** | **99.3%** | **94.1%** | **97.4%** | **100.0%** | **99.6%** | **1** | **0/278** |
| Pandoc 3.9 | 95.9% | 98.0% | 96.4% | 90.7% | 94.1% | **99.3%** | 90.2% | 90.7% | 99.1% | 84.8% | 182 | **0/278** |
| MarkItDown 0.1.9 | 94.8% | 98.1% | 96.2% | 92.5% | 89.3% | 98.5% | 88.7% | 91.8% | 84.3% | 98.7% | 126 | **0/278** |
| mammoth 1.13.0 + turndown 7.2.5 | 94.8% | 95.2% | 96.1% | 88.7% | 91.9% | 85.5% | 62.5% | 91.4% | 95.7% | 98.7% | 172 | **0/278** |
| Syncfusion DocIO (.NET) 28.2.4.0 | 91.5% | 90.0% | 95.6% | 89.9% | 73.8% | 98.4% | 75.1% | 87.6% | 90.4% | 87.4% | 221 | 1/278 |
| LibreOffice 26.8.3.1 | 90.3% | 90.7% | 95.1% | 80.5% | 78.9% | 90.0% | 66.7% | 76.4% | 96.5% | 51.2% | 1252 | 1/278 |
| Docling 2.151.1 | 89.9% | 97.0% | 95.0% | 98.3% | 90.8% | 91.1% | 79.6% | 86.2% | 92.2% | 0.0% | 280 | **0/278** |
| Aspose.Words (.NET) 26.11.0.0 | 85.3% | 97.3% | 96.2% | 79.5% | 89.8% | 97.0% | 80.9% | 96.9% | **100.0%** | 53.5% | 166 | **0/278** |

| Converter | Median cold time | Median warm time | Warm p95 latency | Warm docs/sec | Median cold peak RSS | Max cold peak RSS |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf 1.0.0 | **58.8 ms** | **3.6 ms** | **16.0 ms** | **193.2** | **70.5 MiB** | **105.5 MiB** |
| Pandoc 3.9 | 83.5 ms | 82.0 ms | 151.4 ms | 14.1 | 174.2 MiB | 180.4 MiB |
| MarkItDown 0.1.9 | 787.0 ms | 168.9 ms | 335.1 ms | 5.9 | 159.8 MiB | 187.7 MiB |
| mammoth 1.13.0 + turndown 7.2.5 | 153.9 ms | 26.1 ms | 52.1 ms | 40.0 | 97.3 MiB | 179.4 MiB |
| Syncfusion DocIO (.NET) 28.2.4.0 | 215.3 ms | 24.2 ms | 89.2 ms | 30.8 | 74.3 MiB | 175.1 MiB |
| LibreOffice 26.8.3.1 | 820.5 ms | 52.4 ms | 151.0 ms | 14.9 | – | – |
| Docling 2.151.1 | 2747.4 ms | 19.1 ms | 1378.3 ms | 3.2 | 382.1 MiB | 490.7 MiB |
| Aspose.Words (.NET) 26.11.0.0 | 711.0 ms | 27.9 ms | 223.3 ms | 14.9 | 91.7 MiB | 714.1 MiB |

LibreOffice memory is unavailable because GNU time measured its launcher.
<!-- markdown-benchmark:end -->

Markdown accuracy is the mean per document of the measures that apply to it. See the
[Markdown benchmark README](benchmarks/markdown/README.md) for what each measure compares and what
content is neutral.

## Development

```bash
pnpm install
pnpm test          # unit and end-to-end tests (uses Ghostscript, when installed, to read the PDFs back)
pnpm lint          # library, scripts, and browser demo
pnpm typecheck
pnpm build         # dist/

pnpm render in.docx --out out/ --png       # convert, and rasterise with Ghostscript
pnpm corpus list.txt --root ~/docs         # compare page counts and breaks with Word for many files
pnpm demo          # build the browser demo into out/demo (GitHub Pages uses main)
pnpm demo:serve    # serve it on http://localhost:8000 and rebuild on change
```

`list.txt` holds one `.docx` path per line (relative to `--root`). Set `BREAKS=1` to print where
each of Word's recorded page breaks landed. Keep private documents out of this repository.

To prepare an npm release, run `pnpm lint`, `pnpm typecheck`, `pnpm test`, and `npm pack --dry-run`.
The manifest sets public access for `@kanera/docx2pdf`, and npm runs the build when packing
or publishing. The browser demo uses GitHub Actions as its Pages source.

Code map: `src/docx` reads the package into a document model (styles, numbering, drawing),
`src/fonts` parses, substitutes and subsets fonts, `src/layout` breaks lines and paginates into
positioned drawing operations, and `src/pdf` writes those as PDF. `src/util` holds the ZIP reader,
compression adapters and XML parser.

## Related projects

- [Kanera XLSX to PDF](https://github.com/happendev/Kanera-xlsx) converts Excel workbooks in Node.js and the browser.
- [DocFoundry](https://github.com/happendev/DocFoundry) converts Word and Excel files on .NET.

## License

[Apache License 2.0](LICENSE). The bundled fonts keep their own licenses (see [Fonts](#fonts)).
