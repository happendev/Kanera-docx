# Changelog

## 1.0.0 (unreleased)

Initial release.

- Convert DOCX to PDF with `@kanera/docx2pdf` in Node.js or the browser. The `docx2pdf` CLI also exports one PNG per page.
- Extract plain text or Markdown through the API or CLI, or preview and download either format in the browser demo. Markdown keeps headings, list numbering, tables, emphasis and links, and writes equations as LaTeX. The [Markdown benchmark](benchmarks/markdown/README.md) compares it with seven other converters.
- Render Word's page layout, tables, charts and equations, including vertical East Asian text and right-to-left tables. Use `--mac` to match Word for Mac's text placement.
- Keep long footnotes visible across pages and honour the document's page and endnote numbering.
- Use embedded fonts and bundled Office-compatible fonts, with searchable colour emoji.
