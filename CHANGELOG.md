# Changelog

## 1.0.0 (unreleased)

Initial release.

- Convert embedded RGBA PNGs to PDF faster by reusing their lossless colour filters.
- Convert Word documents (`.docx`) to PDF in Node.js, the browser or any JavaScript runtime, with `@kanera/docx2pdf` or the `docx2pdf` CLI.
- Draw pages as PNG or JPEG images with `docxToImages` or `--png` and `--jpeg`, from the same layout as the PDF and without an outside rasteriser.
- Extract plain text or Markdown with `docxToText`, `docxToMarkdown` or `--text` and `--markdown`. Markdown keeps headings, list numbering, tables, emphasis, links and footnotes, and writes equations as LaTeX.
- Start Node conversions faster and reuse common styles, numbering and themes across documents.
- Lay out pages as Word does: tables, footnotes and endnotes, charts, equations, shapes and shadows, vertical East Asian text and right-to-left text. Use `--mac` to match Word for Mac's text placement.
- Use fonts embedded in documents and bundled fonts with Office's glyph widths, and print colour emoji that can still be searched and copied.
- Read PNG, JPEG, GIF, BMP, TIFF, SVG, EMF and WMF pictures. Each picture is embedded once, and JPEG and fax data are kept as stored unless recompressed.
- Stop hostile documents with limits on size, pixels, fonts, nesting and pages.
