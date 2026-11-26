# Changelog

## 1.0.1

- Improved Word-to-PDF visual match from 92.6% to 92.7% across 278 benchmark documents and from 90.53% to 90.74%.
- Matched Word more closely when balancing columns, handling page breaks, and positioning rotated inline pictures and text inside shapes.
- Improved font substitution on Windows and the rendering of character borders, table border joins, and thin lines.

## 1.0.0

Initial release.

- Convert Word documents (`.docx`) to PDF in Node.js, the browser or any JavaScript runtime, with `@kanera/docx2pdf` or the `docx2pdf` CLI.
- Draw pages as PNG or JPEG images with `docxToImages` or `--png` and `--jpeg`, from the same layout as the PDF and without an outside rasteriser, at a resolution or fitted to a size (`width`, `height`). `onPage` receives each image as it is drawn.
- Stop a conversion with an `AbortSignal` (`signal`).
- Extract plain text or Markdown with `docxToText`, `docxToMarkdown` or `--text` and `--markdown`. Markdown keeps headings, list numbering, tables, emphasis, links and footnotes, and writes equations as LaTeX.
- Lay out pages as Word does: tables, footnotes and endnotes, charts, equations, shapes and shadows, vertical East Asian text and right-to-left text. Use `--mac` to match Word for Mac's text placement.
- Use fonts embedded in documents and bundled fonts with Office's glyph widths, and print colour emoji that can still be searched and copied.
- Read PNG, JPEG, GIF, BMP, TIFF, SVG, EMF and WMF pictures. Each picture is embedded once, and JPEG and fax data are kept as stored unless recompressed.
- Stop hostile documents with limits on size (`maxInputBytes`), pixels, fonts, nesting and pages.
