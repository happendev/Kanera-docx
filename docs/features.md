# Supported features and limitations

What the converter draws from a Word document, what it leaves out, and how it matches Word's layout.
Text and Markdown extraction is described in [Options](options.md#text-extraction).

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
  the same words. Fonts embedded in the `.docx` are used first. See [Fonts](fonts.md).
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
recorded in each file (`pnpm corpus`, see [Development](development.md)).
