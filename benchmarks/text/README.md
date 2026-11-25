# DOCX-to-text benchmark

The benchmark converts the [Word benchmark's](../word/README.md) 278 documents to plain text with
each engine that writes it. It scores the words, their reading order and the document's paragraphs,
then times each engine and measures its memory as the [Markdown benchmark](../markdown/README.md)
does. It shares that benchmark's harness, references and neutral content.

| Engine | How it runs |
| --- | --- |
| @kanera/docx2pdf | `docx2pdf in.docx --text out.txt` |
| Aspose.Words for .NET | `Document.Save` with `SaveFormat.Text` |
| Syncfusion DocIO for .NET | `WordDocument.Save` with `FormatType.Txt` |
| Pandoc | `pandoc -f docx -t plain` |
| Docling | `DocumentConverter().convert(...).document.export_to_text()` |
| mammoth | `mammoth.extractRawText` |

Every engine runs with its default options. MarkItDown writes only Markdown, so it is not in this
benchmark. LibreOffice isn't either: its plain-text export (26.8.3.1, with `txt`, `Text` or
`Text (encoded)`) wrote without end on 28 of the first 185 documents, among them simple tables that
it converts to Markdown in under a second.

## Results

<!-- text-accuracy:start -->
| Converter | Text accuracy | Words | Order | Paragraphs | Leaked markup | Failed |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf 1.0.0 | **98.9%** | **98.8%** | **97.7%** | 99.1% | **0** | **0/278** |
| Aspose.Words (.NET) 26.11.0.0 | 98.4% | **98.8%** | 97.0% | 97.3% | 2 | **0/278** |
| Syncfusion DocIO (.NET) 28.2.4.0 | 97.9% | 98.6% | 96.2% | 99.3% | 2 | 1/278 |
| mammoth 1.13.0 | 97.9% | 98.3% | 96.1% | **99.4%** | **0** | **0/278** |
| Docling 2.151.1 | 97.0% | 97.2% | 95.1% | 99.0% | 7 | **0/278** |
| Pandoc 3.9 | 90.4% | 98.0% | 94.5% | 69.8% | **0** | **0/278** |
<!-- text-accuracy:end -->

<!-- text-speed:start -->
| Converter | Median cold time | Median warm time | Warm p95 latency | Warm docs/sec | Median cold peak RSS | Max cold peak RSS |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf 1.0.0 | **61.1 ms** | **1.6 ms** | **13.7 ms** | **279.2** | **70.5 MiB** | **97.6 MiB** |
| Aspose.Words (.NET) 26.11.0.0 | 714.8 ms | 26.6 ms | 62.0 ms | 31.7 | 91.0 MiB | 107.9 MiB |
| Syncfusion DocIO (.NET) 28.2.4.0 | 215.6 ms | 29.1 ms | 50.0 ms | 35.2 | 73.8 MiB | 108.8 MiB |
| mammoth 1.13.0 | 139.0 ms | 25.1 ms | 42.2 ms | 45.6 | 96.3 MiB | 148.8 MiB |
| Docling 2.151.1 | 2534.9 ms | 17.1 ms | 327.1 ms | 16.1 | 377.2 MiB | 485.1 MiB |
| Pandoc 3.9 | 86.6 ms | 84.7 ms | 158.8 ms | 14.0 | 167.6 MiB | 171.7 MiB |
<!-- text-speed:end -->

Bold marks the best value in each column, ties included.

## What is measured

| Measure | Compared with the reference |
| --- | --- |
| Text accuracy | Mean per document of the measures below that apply to it. A failed conversion scores zero. |
| Words | F1 of the words, over all documents. |
| Order | Similarity of the word sequences (normalised Indel distance), averaged over documents. |
| Paragraphs | Share of the document's body paragraphs found whole on a line of their own, over all documents. |
| Leaked markup | Markdown emphasis or link syntax and HTML tags in the text, beyond those in the document's text. |

Words and order come from Word's own PDF of the document where there is one (120 documents), and
from the DOCX text otherwise, exactly as in the [Markdown benchmark](../markdown/README.md#references).
The same content is [neutral](../markdown/README.md#neutral-content): headers, footers, text boxes,
charts and equations. An engine that writes headers and footers isn't penalised for their words.

Paragraphs come from the document's XML. A paragraph counts when some line of the text holds all of
it and nothing else, ignoring a list number or bullet in front of either. This rewards text that
keeps one paragraph per line, rather than wrapping a paragraph across lines or running paragraphs
together. Paragraphs in tables and text boxes, and paragraphs with a line break or a note mark, are
left out, because a correct engine may write them on more than one line.

## Run from the repository root

Set up the [Markdown benchmark](../markdown/README.md#run-from-the-repository-root) first; this
benchmark uses its Node packages, .NET runner and virtual environment. Then:

```sh
pnpm build
ASPOSE_WORDS_LICENSE="/absolute/path/to/Aspose.Words.lic" \
SYNCFUSION_LICENSE_KEY="your-v28-key" \
  out/benchmark-markdown/venv/bin/python benchmarks/text/convert.py
out/benchmark-markdown/venv/bin/python benchmarks/text/score.py
```

- `convert.py` takes the Markdown benchmark's options: `--engines kanera,pandoc` reruns only the
  named engines and keeps the other results, and `--no-warm` skips the warm runs. `--limit 5`
  converts only the first five documents for a quick check; score them with `score.py --limit 5`,
  which leaves the READMEs unchanged.
- Converting with every engine takes about half an hour, most of it Docling.
- A Markdown run and a text run share the lock on `out/benchmark-markdown/`, so they never overlap
  and never time each other.

### Outputs

The text, timings and scores go to `out/benchmark-text/`, laid out as the Markdown benchmark's are:

- `txt/<engine>/<case>.txt`, the text of each engine's cold run;
- `conversion.json`, with versions and per-document timings;
- `results.csv`, with every measure per document and engine;
- `summary.json` and `report.md`.

Scoring the full comparison updates the tables above and the plain-text summary in the main
README. Use `--no-readme` to leave them unchanged.

Timing works as in the [Markdown benchmark](../markdown/README.md#timing).
