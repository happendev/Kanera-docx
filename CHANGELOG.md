# Changelog

## 1.0.0 (unreleased)

Initial release.

- Convert DOCX to PDF with `@kanera/docx2pdf` in Node.js or the browser.
- Render pages as PNG or JPEG images with `docxToImages` or the CLI's `--png` and `--jpeg`, drawn from the same layout without an outside rasteriser.
- Extract plain text or Markdown through the API or CLI, or preview and download either format in the browser demo. Markdown keeps headings, list numbering, tables, emphasis and links, and writes equations as LaTeX. The [Markdown benchmark](benchmarks/markdown/README.md) compares it with seven other converters.
- Render Word's page layout, tables, charts and equations, including vertical East Asian text and right-to-left tables. Use `--mac` to match Word for Mac's text placement.
- Keep long footnotes visible across pages and honour the document's page and endnote numbering.
- Use embedded fonts and bundled Office-compatible fonts, with searchable colour emoji.
- Draw bubble charts, trendline equations and R² values, and rounded chart frames.
- Embed a picture once when a document stores it under several names, and keep palette TIFFs indexed: smaller PDFs.
- Show CCITT fax and JPEG-compressed TIFFs, embedding their fax or JPEG data as it is.
- Recompress JPEGs with the built-in codec, about twice as fast and with smaller 4:2:0 output, and read JPEGs with unusual chroma sampling.
- Keep large photos in page images: a JPEG past the pixel limits is decoded smaller, and the limits apply to each page.
- Draw shapes' and pictures' shadows, and blur shape and chart shadows as Word does.
