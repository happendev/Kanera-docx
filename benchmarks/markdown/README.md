# DOCX-to-Markdown benchmark

The benchmark converts the [Word benchmark's](../word/README.md) 278 documents to Markdown with each
engine. It scores what a reader of the Markdown sees, then times each engine and measures its memory
the same way as the PDF benchmark.

| Engine | How it runs |
| --- | --- |
| @kanera/docx2pdf | `docx2pdf in.docx --markdown out.md` |
| Aspose.Words for .NET | `Document.Save` with `SaveFormat.Markdown` |
| Syncfusion DocIO for .NET | `WordDocument.Save` with `FormatType.Markdown` |
| Pandoc | `pandoc -f docx -t gfm` |
| Docling | `DocumentConverter().convert(...).document.export_to_markdown()` |
| MarkItDown | `MarkItDown().convert(...).text_content` |
| mammoth + turndown | mammoth's HTML, converted by turndown with its GitHub plugin, `#` headings, fenced code and `-` bullets |
| LibreOffice | `soffice --headless --convert-to md` |

Every other engine runs with its default options.

## Results

<!-- markdown-accuracy:start -->
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
<!-- markdown-accuracy:end -->

<!-- markdown-speed:start -->
| Converter | Median cold time | Median warm time | Warm p95 latency | Warm docs/sec | Median cold peak RSS | Max cold peak RSS |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf 1.0.0 | **64.0 ms** | **4.0 ms** | **16.0 ms** | **179.3** | **70.4 MiB** | **105.7 MiB** |
| Pandoc 3.9 | 83.5 ms | 82.0 ms | 151.4 ms | 14.1 | 174.2 MiB | 180.4 MiB |
| MarkItDown 0.1.9 | 787.0 ms | 168.9 ms | 335.1 ms | 5.9 | 159.8 MiB | 187.7 MiB |
| mammoth 1.13.0 + turndown 7.2.5 | 153.9 ms | 26.1 ms | 52.1 ms | 40.0 | 97.3 MiB | 179.4 MiB |
| Syncfusion DocIO (.NET) 28.2.4.0 | 215.3 ms | 24.2 ms | 89.2 ms | 30.8 | 74.3 MiB | 175.1 MiB |
| LibreOffice 26.8.3.1 | 820.5 ms | 52.4 ms | 151.0 ms | 14.9 | – | – |
| Docling 2.151.1 | 2747.4 ms | 19.1 ms | 1378.3 ms | 3.2 | 382.1 MiB | 490.7 MiB |
| Aspose.Words (.NET) 26.11.0.0 | 711.0 ms | 27.9 ms | 223.3 ms | 14.9 | 91.7 MiB | 714.1 MiB |

LibreOffice memory is unavailable because GNU time measured its launcher.
<!-- markdown-speed:end -->

Bold marks the best value in each column, ties included.

## What is measured

Every engine's Markdown is rendered with one CommonMark renderer, [markdown-it-py](https://github.com/executablebooks/markdown-it-py).
It has GitHub's tables, strikethrough and footnotes, and `$` math with GitHub's rules. Scores come
from the rendered HTML, so syntax the renderer doesn't understand shows as text and counts against
the engine. A picture embedded as a `data:` URI counts as a picture whatever its format, as a viewer
shows it or a broken-image icon rather than its base64 text.

| Measure | Compared with the reference |
| --- | --- |
| Markdown accuracy | Mean per document of the measures below that apply to it. A failed conversion scores zero. |
| Text | F1 of the words, over all documents. |
| Order | Similarity of the word sequences (normalised Indel distance), averaged over documents. |
| Headings | F1 of headings, counted when both the text and the level match. All of a document's levels may shift by the same amount, so headings that start at `##` still match. |
| Lists | F1 of list items by their text. |
| Tables | F1 of the words in table cells. |
| Bold, Italic | F1 of the words shown in bold or italic, on the 120 documents with a Word PDF. |
| Links | Share of external link targets present. |
| Images | Pictures found, against pictures in the document. |
| Leaked syntax | Markdown markers shown as text, such as `**`, `[x](` or a stray `#`, beyond those in the document's text. |

### References

No engine under test produces a reference.

- **Words, order, bold and italic** come from Word's own PDF of the document, read with PyMuPDF.
  These are the 100 Hugging Face documents and the 20 vendored documents with a Microsoft Word PDF.
  - A word counts as bold or italic when Word drew most of its characters in a bold or italic face.
  - Text drawn both filled and stroked counts as bold, and slanted glyphs count as italic: Word fakes
    bold and italic this way for fonts without those faces.
  - For the other 158 documents, the words and their order come from the DOCX text.
- **Structure** comes from a small reader of the DOCX XML in `score.py`. It covers:
  - headings by outline level;
  - numbered and bulleted paragraphs;
  - tables of two or more columns;
  - external hyperlinks and pictures.

### Neutral content

Some content has no single right answer in Markdown. It doesn't count for or against an engine:

- Headers and footers.
  - Lines in the page margins of Word's PDF are left out.
  - So are lines at the top or bottom of a page made only of header and footer words, when they repeat at
    the same height on another page. These come from a header taller than its margin.
  - An engine that writes headers and footers isn't penalised for their words.
- Text boxes, and tables inside them. Their words count when Word's PDF shows them, but extra ones are
  not penalised.
- Single-column tables. They are boxes around paragraphs as much as tables.
- Table-of-contents entries and the Title and Subtitle styles. Each may be written as a heading, a list
  or a paragraph.
- Chart and equation text. Word draws both. Engines write a chart's title and data, a picture or
  nothing, and an equation as LaTeX, MathML, a picture or text, which a word comparison can't match.
- Emphasis inside headings and titles. Word draws headings bold, and the heading already says so.

### Limitations

- Word's PDF has no structure. Bold and italic are compared as word counts per document, not per
  position.
- Heading words are removed from the PDF's bold and italic by count. A heading that Word doesn't draw
  bold can take a matching word out of the body's bold.
- Header lines are recognised by their words. A body line made only of header words, at the top of a
  page and repeated on another page, is also left out.
- Speed and memory are measured on Linux. LibreOffice runs through Flatpak here, so its memory is not
  available.

## Run from the repository root

Run the [Word benchmark's setup](../word/README.md#run-from-the-repository-root) first. It downloads
and verifies the documents and their Word PDFs. Then use Linux with GNU time at `/usr/bin/time`, Node 24,
Python 3.11–3.13, and the .NET 10 SDK for Aspose and Syncfusion. Put headless LibreOffice (`soffice`) on PATH.

```sh
pnpm build
npm ci --prefix benchmarks/markdown --ignore-scripts
dotnet build benchmarks/markdown/dotnet -c Release -p:RestoreLockedMode=true

python3 -m venv out/benchmark-markdown/venv
out/benchmark-markdown/venv/bin/python -m pip install -r benchmarks/markdown/requirements.txt \
  --extra-index-url https://download.pytorch.org/whl/cpu

ASPOSE_WORDS_LICENSE="/absolute/path/to/Aspose.Words.lic" \
SYNCFUSION_LICENSE_KEY="your-v28-key" \
  out/benchmark-markdown/venv/bin/python benchmarks/markdown/convert.py
out/benchmark-markdown/venv/bin/python benchmarks/markdown/score.py
```

- Aspose.Words and Syncfusion DocIO are skipped when their license variable is unset. Their licenses
  work as they do in the [Word benchmark](../word/README.md#include-asposewords-for-net).
- `convert.py --engines kanera,pandoc` reruns only the named engines and keeps the other results.
  `--no-warm` skips the warm runs.
- Converting with every engine takes about an hour, most of it Docling and LibreOffice.

### Outputs

The Markdown, timings and scores go to `out/benchmark-markdown/`:

- `md/<engine>/<case>.md`, the Markdown of each engine's cold run;
- `conversion.json`, with versions and per-document timings;
- `results.csv`, with every measure per document and engine;
- `summary.json` and `report.md`.

Scoring the full comparison updates the tables above and in the main README. Use `--no-readme`
to leave them unchanged.

### Timing

- **Cold**: one process per document, timed from start to exit. GNU time records the peak resident memory.
- **Warm**: each engine converts the documents in one resident process, after one untimed conversion.
  - Pandoc has no resident mode, so its warm time is a process per document.
  - LibreOffice converts the documents as one batch, timed by when each output appears.

Each engine times out after 300 seconds per document.
