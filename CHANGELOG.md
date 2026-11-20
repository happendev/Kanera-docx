# Changelog

## 1.0.0 (unreleased)

### Added

- Convert DOCX files to PDF with `@kanera/docx2pdf` in Node.js 20+, in the browser, or through the `docx2pdf` CLI. The runtime-neutral `/core` entry point accepts document bytes and caller-supplied fonts.
- Use `docxToPdf` to get PDF bytes, or `convertDocx` to also get the page count and conversion warnings.
- Render Word-style line breaks and pagination, including tables with repeated headers, footnotes, and text wrapping around floating objects. Render images, supported charts and diagrams, and preserve hyperlinks and bookmark links in the PDF.
- Use embedded fonts, supply custom fonts, or fall back to bundled fonts with matching glyph widths for many common Office fonts. The browser can load a smaller font pack on demand.
- Compare browser conversion results with reference PDFs exported by Word in the demo. The [benchmark suite](benchmarks/word/README.md) documents how page counts and visual accuracy are measured.
