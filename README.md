<a href="https://www.kanera.app"><picture>
  <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/happendev/Kanera/main/apps/web/public/assets/logo/Kanera%20dark%20long.svg">
  <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/happendev/Kanera/main/apps/web/public/assets/logo/Kanera%20light%20long.svg">
  <img alt="Kanera" src="https://raw.githubusercontent.com/happendev/Kanera/main/apps/web/public/assets/logo/Kanera%20light%20long.svg" width="200">
</picture></a>

# Kanera DOCX to PDF

**High-fidelity Word to PDF conversion for Node.js and the browser.**
Convert `.docx` files to PDF or page images, and extract them as plain text or Markdown, with a
TypeScript library and CLI. No Word, no LibreOffice, no native add-ons or WebAssembly.

[![npm](https://img.shields.io/npm/v/@kanera/docx2pdf?logo=npm&label=npm&color=cb3837)](https://www.npmjs.com/package/@kanera/docx2pdf)
[![Node.js 22+](https://img.shields.io/badge/Node.js-22%2B-5FA04E?logo=nodedotjs)](https://nodejs.org/)
[![Browser](https://img.shields.io/badge/browser-ES2022-4285F4)](https://github.com/happendev/kanera-docx/blob/main/docs/runtimes.md#browser)
[![TypeScript](https://img.shields.io/badge/types-included-3178C6?logo=typescript)](https://www.typescriptlang.org/)
[![License: Apache 2.0](https://img.shields.io/badge/license-Apache%202.0-blue)](https://github.com/happendev/kanera-docx/blob/main/LICENSE)

[Live demo](https://happendev.github.io/Kanera-docx/) · [Quick start](#quick-start) · [Usage](#usage) · [Benchmarks](#benchmarks) · [Documentation](#documentation) · [Release notes](https://github.com/happendev/kanera-docx/blob/main/CHANGELOG.md) · [Report a conversion issue](https://github.com/happendev/Kanera-docx/issues/new?template=conversion-bug.yml)

> **Stable and in production use.** `@kanera/docx2pdf` has a stable 1.x API and is used in
> production by [Kanera](https://github.com/happendev/Kanera) ([kanera.app](https://www.kanera.app)).
> Its output is scored against the pages Word printed for 278 benchmark documents, and for 2,650
> documents in a full-corpus check.

## What is it?

Kanera DOCX to PDF is a TypeScript library and command-line tool that lays out Word documents the
way Word does, then writes them as PDF or images. It also reads them as text or Markdown without
laying out pages. Every output works in Node.js and the browser.

| From `.docx` to | Functions | CLI |
| --- | --- | --- |
| PDF | `docxToPdf`, `convertDocx` | `docx2pdf in.docx` |
| PNG or JPEG page images | `docxToImages`, `convertDocxToImages` | `--png <dir>`, `--jpeg <dir>` |
| Plain text, Markdown | `docxToText`, `docxToMarkdown`, `extractText` | `--text`, `--markdown` |

## Why Kanera DOCX to PDF?

- **Closest to Word.** Against the pages Word prints itself, it scores higher than LibreOffice,
  Aspose, Syncfusion, EigenPal and the Rust converters, and its Markdown scores higher than Pandoc,
  MarkItDown, mammoth and Docling.
- **Runs in the browser and on a Node.js backend.** Convert on your server with Node.js 22 or
  later, or right in the user's browser so documents never have to be uploaded. It also runs in other
  JavaScript environments, such as edge functions.
- **Fast and efficient.** The fastest Word to PDF converter in the benchmark, from a fresh process
  or a reused one, using a fraction of LibreOffice's memory. Markdown extraction is the fastest and
  lightest of every converter tested.
- **Standalone, with zero dependencies.** Pure TypeScript that installs nothing else: no Word,
  LibreOffice, native add-ons or WebAssembly. Open fonts with Office's glyph widths are bundled.
- **Deterministic.** System fonts are off by default, so a file converts to the same PDF on every
  machine.
- **Safe for untrusted files.** Limits on archive size, pixels, fonts, nesting and pages stop
  hostile documents.

<!-- benchmark-glance:start -->
| Conversion | Accuracy | Best other converter | Time per file | Files per second | Peak memory |
| --- | ---: | ---: | ---: | ---: | ---: |
| Word to PDF | **92.7%** | 80.3% LibreOffice | **3.9 ms** vs 59.6 ms | **77.8** vs 10.3 | **80.9 MiB** vs 252.7 MiB |
| Word to page images | **92.8%** | 80.9% LibreOffice + pdftoppm | **11.5 ms** vs 171.4 ms † | **30.3** vs 2.5 † | 101.7 MiB vs – |
| Word to Markdown | **98.5%** | 95.9% Pandoc | **1.9 ms** vs 85.6 ms | **241.3** vs 13.8 | **70.3 MiB** vs 167.8 MiB |
| Word to text | **98.9%** | 98.4% Aspose.Words (.NET) | **1.6 ms** vs 26.6 ms | **288.7** vs 31.7 | **70.3 MiB** vs 91.0 MiB |
<!-- benchmark-glance:end -->

Kanera against the most accurate other converter on 278 documents. Times are warm medians; memory
is the median peak of a fresh process. † marks a converter timed on a 50-document sample, and
LibreOffice's memory couldn't be measured for page images. Full results:
[Word to PDF](https://github.com/happendev/kanera-docx/blob/main/benchmarks/word/README.md#results) ·
[Page images](https://github.com/happendev/kanera-docx/blob/main/benchmarks/image/README.md#results) ·
[Markdown](https://github.com/happendev/kanera-docx/blob/main/benchmarks/markdown/README.md#results) ·
[Text](https://github.com/happendev/kanera-docx/blob/main/benchmarks/text/README.md#results)

## Quick start

Install the package in a project that uses Node.js 22 or later:

```bash
npm install @kanera/docx2pdf
```

```ts
import { writeFile } from 'node:fs/promises';
import { docxToPdf } from '@kanera/docx2pdf';

await writeFile('report.pdf', await docxToPdf('report.docx'));
```

Or run the command-line tool without installing it:

```bash
npx @kanera/docx2pdf report.docx         # writes report.pdf
```

The package includes the fonts it needs. Nothing else is needed.

## Usage

### Bytes and warnings

In Node, every function takes a path, `Uint8Array`, `Buffer` or `ArrayBuffer`. `convertDocx` also
returns the page count and warnings, such as fonts that were substituted.

```ts
import { convertDocx } from '@kanera/docx2pdf';

const { pdf, pageCount, warnings } = await convertDocx(bytes, { wordPlatform: 'mac' });
for (const warning of warnings) console.warn(warning);
```

### Browser

Serve the package's `fonts/` folder and pass its URL; the browser fetches only the font faces a
document uses.

<!-- browser-conversion-example -->
```ts
import { convertDocx } from '@kanera/docx2pdf'; // resolves to the browser build

const { pdf } = await convertDocx(file, { fontBaseUrl: '/docx2pdf-fonts/' });
const pdfBlob = new Blob([pdf], { type: 'application/pdf' });
```

Run conversions in a Web Worker to keep the page responsive. See
[Node, browser and other runtimes](https://github.com/happendev/kanera-docx/blob/main/docs/runtimes.md)
for serving the fonts, a worker example, and `@kanera/docx2pdf/core` for other runtimes.

### Web server

Conversion is synchronous, so a large document blocks other requests while it converts.

```ts
const { pdf } = await convertDocx(Buffer.concat(chunks));
res.writeHead(200, { 'content-type': 'application/pdf' }).end(pdf);
```

The [full example](https://github.com/happendev/kanera-docx/blob/main/docs/runtimes.md#in-a-web-server)
limits the upload size and reports documents that can't be converted.

### Page images

`docxToImages` draws pages as PNG or JPEG from the same layout as the PDF, without a PDF renderer.

```ts
for (const { page, data } of await docxToImages('in.docx', { dpi: 150 })) {
  await writeFile(`page-${page}.png`, data);
}
const [cover] = await docxToImages('in.docx', { format: 'jpeg', quality: 85, pages: [1] });
```

### Text and Markdown

`docxToText` and `docxToMarkdown` read what Word shows, in reading order, without laying out pages.
Markdown keeps headings, lists, tables, emphasis, links, footnotes and equations.

```ts
const text = await docxToText('in.docx');
const markdown = await docxToMarkdown('in.docx', { includeHeadersFooters: true });
```

### Command line

```bash
npx @kanera/docx2pdf in.docx out.pdf --mac           # match Word for Mac's PDFs
npx @kanera/docx2pdf in.docx --font BrandSans.ttf    # add a font
npx @kanera/docx2pdf in.docx --png pages --dpi 200   # pages/in-1.png, ...
npx @kanera/docx2pdf in.docx --markdown -            # Markdown to standard output
```

In a project with the package installed, use `npx --no-install docx2pdf`. See the
[CLI reference](https://github.com/happendev/kanera-docx/blob/main/docs/cli.md).

### Fonts

Word documents name fonts that are rarely free to redistribute. The package bundles open fonts
with the same glyph widths (Carlito for Calibri, Liberation Sans for Arial, and so on), so lines
wrap at the same words. Fonts embedded in a document are used first. Pass your own fonts in
`fonts`, or set `systemFonts` in Node to use the machine's. See
[Fonts](https://github.com/happendev/kanera-docx/blob/main/docs/fonts.md).

### Untrusted documents

Conversion stops with an error when a file exceeds limits on size, pixels, fonts, nesting or pages.
For a hosted converter, also run each job in its own process with a memory limit and a deadline.
See [Converting untrusted documents](https://github.com/happendev/kanera-docx/blob/main/docs/options.md#converting-untrusted-documents).

## Documentation

- [Options](https://github.com/happendev/kanera-docx/blob/main/docs/options.md): every conversion,
  page image and text option, image compression, text extraction and limits.
- [Node, browser and other runtimes](https://github.com/happendev/kanera-docx/blob/main/docs/runtimes.md):
  ES modules and CommonJS, bundling, web servers, serving fonts, Web Workers and the core build.
- [Command-line tool](https://github.com/happendev/kanera-docx/blob/main/docs/cli.md): every flag.
- [Fonts](https://github.com/happendev/kanera-docx/blob/main/docs/fonts.md): bundled substitutes,
  symbol, East Asian and emoji fonts, embedding and licences.
- [Supported features and limitations](https://github.com/happendev/kanera-docx/blob/main/docs/features.md):
  what is supported and how the converter matches Word's layout.
- [Development](https://github.com/happendev/kanera-docx/blob/main/docs/development.md): scripts,
  releases and the code map.

## Benchmarks

Each converter's PDFs are drawn at 96 DPI and compared with the pages Word printed for the same
file. Visual match is the mean ink overlap with one pixel of tolerance. Missing pages and failed
conversions score zero, and each document has equal weight. Times are medians and include file
I/O. Cold timings start a fresh process per file. Warm timings reuse one long-running process.
Timed conversions run one at a time. † marks a converter timed on a fixed sample of 50 files.

### Word to PDF

278 documents with Word for Mac, Microsoft 365 Word and LibreOffice references.
[Corpus, method and full results](https://github.com/happendev/kanera-docx/blob/main/benchmarks/word/README.md#results)

<!-- benchmark-summary:start -->
| Converter | Visual match | Page counts matched | Fresh process | Reused process | Warm docs/sec | Median peak memory |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Kanera | **92.7%** | **273/278** | **85.6 ms** | **3.9 ms** | **77.8** | 80.9 MiB |
| LibreOffice | 80.3% | 256/278 | 791.6 ms | 59.6 ms | 10.3 | 252.7 MiB |
| EigenPal | 70.9% | 257/278 | 766.1 ms | 255.3 ms | 3.7 | 276.6 MiB |
| Syncfusion DocIO (.NET) | 56.8% | 246/278 | 589.2 ms | 60.7 ms | 1.1 | 121.4 MiB |
| Aspose.Words (.NET) | 51.9% | 250/278 | 1040.4 ms | 36.6 ms | 11.2 | 105.3 MiB |
| office2pdf (Rust) | 49.1% | 240/278 | 321.6 ms | 114.6 ms | 7.5 | **42.8 MiB** |
| MiniPdf (.NET) | 43.8% | 247/278 | 351.9 ms | 163.6 ms | 6.2 | 135.4 MiB |
| rdocx (Rust) | 40.5% | 214/278 | 208.8 ms | 26.0 ms | 13.7 | 146.1 MiB |
| dxpdf (Rust) | 36.4% | 236/278 | 986.7 ms | 1011.6 ms | 1.0 | 257.5 MiB |
<!-- benchmark-summary:end -->

<!-- full-corpus-accuracy:start -->
Kanera scores **90.74%** on the full corpus of 2,650 documents (10,376 reference pages), with matching page counts on 2,561 and 1 failed conversion.
<!-- full-corpus-accuracy:end -->

### Word to page images

The same 278 documents drawn as PNG pages at 150 DPI and compared with Word's pages in the same way.
[Method and full results](https://github.com/happendev/kanera-docx/blob/main/benchmarks/image/README.md#results)

<!-- image-benchmark-summary:start -->
| Converter | Visual match | Page counts matched | Fresh process | Reused process | Warm docs/sec | Median peak memory |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Kanera | **92.8%** | **273/278** | **131.9 ms** | **11.5 ms** | **30.3** | **101.7 MiB** |
| LibreOffice + pdftoppm | 80.9% | 256/278 | 927.8 ms † | 171.4 ms † | 2.5 † | – |
| Syncfusion DocIO (.NET) | 58.9% | 246/278 | 578.2 ms | 72.1 ms | 1.8 | 113.2 MiB |
| Aspose.Words (.NET) | 53.2% | 250/278 | 1038.5 ms † | 82.7 ms † | 3.6 † | 117.9 MiB † |
| rdocx (Rust) | 44.0% | 235/278 | 251.4 ms | 39.6 ms | 6.9 | 146.2 MiB |
<!-- image-benchmark-summary:end -->

### Word to Markdown

The same 278 documents, rendered with one CommonMark renderer. Words, reading order, bold and
italic are scored against Word's own PDF; headings, lists, tables, links and pictures against the
document. [Method and full results](https://github.com/happendev/kanera-docx/blob/main/benchmarks/markdown/README.md)

<!-- markdown-benchmark-summary:start -->
| Converter | Markdown accuracy | Fresh process | Reused process | Warm docs/sec | Median peak memory |
| --- | ---: | ---: | ---: | ---: | ---: |
| Kanera | **98.5%** | **63.6 ms** | **1.9 ms** | **241.3** | **70.3 MiB** |
| Pandoc | 95.9% | 92.6 ms | 85.6 ms | 13.8 | 167.8 MiB |
| MarkItDown | 94.8% | 685.8 ms | 155.2 ms | 6.5 | 157.0 MiB |
| mammoth + turndown | 94.8% | 159.1 ms | 26.0 ms | 40.5 | 98.7 MiB |
| Syncfusion DocIO (.NET) | 91.5% | 222.8 ms | 26.6 ms | 30.4 | 73.8 MiB |
| LibreOffice | 90.3% | 782.8 ms | 47.3 ms | 17.0 | 249.7 MiB |
| Docling | 89.9% | 2432.5 ms † | 19.6 ms † | 6.6 † | 376.7 MiB † |
| Aspose.Words (.NET) | 85.3% | 741.3 ms | 26.0 ms | 14.4 | 90.9 MiB |
<!-- markdown-benchmark-summary:end -->

### Word to text

The same 278 documents as plain text. Words and reading order are scored against Word's own PDF, and
paragraphs against the document.
[Method and full results](https://github.com/happendev/kanera-docx/blob/main/benchmarks/text/README.md#results)

<!-- text-benchmark-summary:start -->
| Converter | Text accuracy | Fresh process | Reused process | Warm docs/sec | Median peak memory |
| --- | ---: | ---: | ---: | ---: | ---: |
| Kanera | **98.9%** | **60.9 ms** | **1.6 ms** | **288.7** | **70.3 MiB** |
| Aspose.Words (.NET) | 98.4% | 714.8 ms | 26.6 ms | 31.7 | 91.0 MiB |
| Syncfusion DocIO (.NET) | 97.9% | 215.6 ms | 29.1 ms | 35.2 | 73.8 MiB |
| mammoth | 97.9% | 139.0 ms | 25.1 ms | 45.6 | 96.3 MiB |
| Docling | 97.0% | 2534.9 ms | 17.1 ms | 16.1 | 377.2 MiB |
| Pandoc | 90.4% | 86.6 ms | 84.7 ms | 14.0 | 167.6 MiB |
<!-- text-benchmark-summary:end -->

## Development

```bash
pnpm install
pnpm build
pnpm test
```

The tests build their Word documents in code; the benchmarks use the vendored and downloaded
documents in [`benchmarks`](https://github.com/happendev/kanera-docx/tree/main/benchmarks). See
[Development](https://github.com/happendev/kanera-docx/blob/main/docs/development.md) for the other
scripts and the code layout. Keep private documents out of this repository.

## Related projects

- [DocFoundry](https://github.com/happendev/DocFoundry) converts Word, Excel and PowerPoint files
  on .NET.
- [Kanera XLSX to PDF](https://github.com/happendev/Kanera-xlsx) and
  [Kanera PPTX to PDF](https://github.com/happendev/kanera-pptx) convert Excel workbooks and
  PowerPoint presentations in Node.js and the browser.

## License

Apache License 2.0. See [LICENSE](https://github.com/happendev/kanera-docx/blob/main/LICENSE) and
[NOTICE](https://github.com/happendev/kanera-docx/blob/main/NOTICE). The bundled fonts keep their own
licences, listed in [`fonts/licenses`](https://github.com/happendev/kanera-docx/tree/main/fonts/licenses).
