# Options

`convertDocx`, `docxToPdf`, `docxToImages` and `convertDocxToImages` take the options below. The
text functions take the [text extraction options](#text-extraction).

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

| Option          | Where   | Default  | Meaning                                                              |
| --------------- | ------- | -------- | -------------------------------------------------------------------- |
| `fonts`         | all     | `[]`     | Fonts to use (`{ data }` or `{ family, bold, italic, load }`). Win over bundled and system fonts. |
| `archiveLimits` | all | 64 MiB per part, 256 MiB total | Maximum uncompressed bytes per ZIP part (`maxPartBytes`) and across distinct parts read (`maxTotalBytes`). Exceeding either limit stops conversion. |
| `imageLimits` | all | 16,777,216 pixels per image, 67,108,864 total | Set `maxImagePixels` and `maxTotalPixels` to budget source raster pixels. Repeated placements of one image count once. JPEGs are only decoded for optional recompression, so one that exceeds the budget is embedded unchanged. |
| `imageQuality` | all | | Encoder quality for eligible colour JPEGs. Set an integer from 1 to 100 or use `false` to disable lossy re-encoding. CLI: `--image-quality <1-100\|off>`. |
| `additionalLinkSchemes` | all | `[]` | PDF links support HTTP, HTTPS, mailto and internal bookmarks. Pass scheme names such as `['tel']` to enable additional URI actions. Text for other destinations is preserved without a link annotation. |
| `useEmbeddedFonts` | all  | `true`   | Use fonts embedded in the `.docx` (including obfuscated `.odttf`).   |
| `info`          | all     | from doc | PDF title, author, subject, keywords, dates.                         |
| `compress`      | all     | `true`   | Deflate PDF content streams.                                         |
| `signal`        | all     |          | An `AbortSignal` that stops the conversion, rejecting with its reason. It is checked between stages, font loads and pages, so a long layout finishes first. |
| `wordPlatform`  | all     | `'windows'` | Set to `'mac'` to match Word for Mac's PDFs, which draw text sizes, baselines and line starts on a 0.24pt grid (11pt text comes out at 11.04pt). Layout is the same. CLI: `--mac`. |
| `format`        | images  | `'png'`  | `'png'` or `'jpeg'`.                                                 |
| `dpi`           | images  | `150`    | Image resolution in dots per inch. CLI: `--dpi`.                     |
| `width`, `height` | images | none  | Image size in pixels (1 to 65535), in place of `dpi`. With both, each page fits within them, keeping its shape unless `keepAspectRatio` is `false`; with one, the other follows the page's shape. |
| `keepAspectRatio` | images | `true` | `false` stretches each page to exactly `width` × `height`.        |
| `quality`       | images  | `90`     | JPEG quality from 1 to 100. CLI: `--quality`.                        |
| `pages`         | images  | all      | Page numbers to draw, from 1. CLI: `--pages 1-3,5`.                  |
| `onPage`        | images  |          | `(image, pageCount) => …` receives each image as soon as it is encoded, in page order, so only the pages in flight stay in memory; the result's images then have empty `data`. Drawing waits for a returned promise. |
| `systemFonts`   | Node    | `false`  | `true` or a list of folders to scan for `.ttf`/`.otf`/`.ttc`.        |
| `fontPack`      | browser | `'full'` | Set to `'common'` to use the smaller browser pack.                      |
| `fontDir`       | Node    | package  | Folder holding the full pack.                                        |
| `fontBaseUrl`   | browser | module   | URL of the selected pack, or `false` to disable automatic loading.   |
| `fetch`         | browser | global   | Custom fetch for font downloads.                                     |

Font priority goes from `fonts` to embedded fonts, system fonts, substitutes and generic
fallbacks. A font passed in `fonts` under a family name replaces every other font of that name,
whatever its own name table says.

## Image compression

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

## Page images

`docxToImages` draws pages as PNG or JPEG images from the same layout as the PDF, without
rendering a PDF or calling an outside rasteriser. `convertDocxToImages` also returns the page count
and warnings. Both take the layout and font options of `convertDocx`, and run in Node and the browser.

Each image has its `page` number, `width`, `height` and encoded `data`. Scored against Word's
pages the way the [benchmark](../benchmarks/word/README.md) scores PDFs, the images match as closely as the PDF does.

## Text extraction

`docxToText` and `docxToMarkdown` return the document's text without rendering it. They take the
same inputs as `docxToPdf`, load no fonts and decode no pictures. `extractText` returns the text
with its warnings.

```ts
import { docxToMarkdown, docxToText, extractText } from '@kanera/docx2pdf';

const text = await docxToText('in.docx');
const markdown = await docxToMarkdown('in.docx', { headersFooters: true });
const { text: md, warnings } = await extractText('in.docx', { format: 'markdown' });
```

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

| Option                  | Default  | Meaning                                                                   |
| ----------------------- | -------- | ------------------------------------------------------------------------- |
| `format`                | `'text'` | `'markdown'` for Markdown (`extractText` only).                           |
| `includeHeadersFooters` | `false`  | Include each distinct header before its section and each footer after it. |
| `includeNotes`          | `true`   | Include footnotes and endnotes after the body.                            |
| `includeImages`         | `true`   | Markdown: include picture references.                                     |
| `archiveLimits`         |          | As for conversion.                                                        |
| `signal`                |          | An `AbortSignal`; extraction does not start once it is aborted.          |

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
