# DOCX-to-page-image benchmark

The benchmark renders every page of the [Word benchmark's](../word/README.md) 278 cases to PNG and
compares the pages with Word's own. It uses the same SHA-256-verified inputs and references, the
same ink score and the same timing protocol as the Word-to-PDF benchmark, so the two sets of results
can be read side by side.

## Results

The [main README](../../README.md#word-to-page-images) shows a summary of these results.

<!-- image-benchmark-results:start -->
| Converter | Mean ink overlap | Matching page count | Failed conversions | Median cold time | Median warm time | Warm p95 latency | Warm docs/sec | Median cold peak RSS | Max cold peak RSS | Total PNG size |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| @kanera/docx2pdf | **92.8%** | **273/278** | **0/278** | **131.9 ms** | **11.5 ms** | **149.9 ms** | **30.3** | **101.7 MiB** | **254.5 MiB** | 156.90 MiB |
| LibreOffice 26.8.3.1 + pdftoppm 26.01.0 | 80.9% | 256/278 | **0/278** | 927.8 ms † | 171.4 ms † | 1318.4 ms † | 2.5 † | – | – | 128.88 MiB |
| Syncfusion DocIO (.NET) 28.2.4.0 | 58.9% | 246/278 | 1/278 | 578.2 ms | 72.1 ms | 1997.8 ms | 1.8 | 113.2 MiB | 660.6 MiB | 73.01 MiB ‡ |
| Aspose.Words (.NET) 26.11.0.0 | 53.2% | 250/278 | **0/278** | 1038.5 ms † | 82.7 ms † | 1038.7 ms † | 3.6 † | 117.9 MiB † | 397.1 MiB † | **118.10 MiB** |
| rdocx (Rust) 0.17.0 | 44.0% | 235/278 | 3/278 | 251.4 ms | 39.6 ms | 592.7 ms | 6.9 | 146.2 MiB | 466.1 MiB | 98.78 MiB |

Peak RSS is the peak resident memory during a successful cold conversion. LibreOffice memory is unavailable because GNU time measured its launcher.

The smallest total PNG size is determined among converters that rendered every document at 150 dpi.

‡ Rendered at a fixed resolution with no setting for 150 dpi (Syncfusion DocIO (.NET) at 96 dpi), so its PNGs are smaller and are not compared for size.

† Timed on a fixed 50-document sample only, because timing every document would take more than 5 minutes; the other documents were converted untimed and count for accuracy and failures.
<!-- image-benchmark-results:end -->

### By suite

Ink overlap by suite and head-to-head counts against each converter.

<!-- image-benchmark-breakdown:start -->
278 documents.

| Cases | Reference | @kanera/docx2pdf | LibreOffice + pdftoppm | Aspose.Words (.NET) | Syncfusion DocIO (.NET) | rdocx (Rust) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 150 classic | Microsoft 365 Word page images | **99.5%** | 92.1% | 52.6% | 62.7% | 46.7% |
| 100 Hugging Face | Word for Mac PDFs | **86.1%** | 70.5% | 55.2% | 55.0% | 43.3% |
| 20 issues | Microsoft 365 Word PDFs | **90.1%** | 54.0% | 50.0% | 54.7% | 33.4% |
| 8 issues | LibreOffice PDFs | 58.6% | **67.4%** | 49.7% | 49.0% | 29.6% |

Per-document ink overlap differences under half a percentage point count as ties.

| Compared with | Kanera higher | Kanera lower | Ties |
| --- | ---: | ---: | ---: |
| LibreOffice + pdftoppm | 159 | 21 | 98 |
| Aspose.Words (.NET) | 256 | 13 | 9 |
| Syncfusion DocIO (.NET) | 260 | 11 | 7 |
| rdocx (Rust) | 242 | 3 | 33 |
<!-- image-benchmark-breakdown:end -->

## Converters

- **@kanera/docx2pdf** runs this repository's CLI, `docx2pdf <input.docx> --png <dir> --dpi 150`,
  for each cold conversion. The warm worker, [`kanera.mjs`](kanera.mjs), calls
  `convertDocxToImages` with the same options.
- **LibreOffice + pdftoppm.** `soffice --convert-to png` writes only the first page, so the
  benchmark measures the usual all-pages route: LibreOffice exports a PDF and Poppler's `pdftoppm`
  rasterizes it at 150 dpi. Both steps are timed together.
- **Aspose.Words for .NET** saves each page with `ImageSaveOptions(SaveFormat.Png)` at a
  resolution of 150 dpi ([runner](aspose/Program.cs)).
- **Syncfusion DocIO** renders the pages with `WordDocument.RenderAsImages()` and
  `DocIORenderer` ([runner](syncfusion/Program.cs)). DocIO 28 has no resolution setting and always
  renders at 96 dpi, so its pages are scored without resampling and its PNG size is not compared.
- **rdocx** renders through its Python bindings with `render_all_pages(150, system_fonts=True)`
  ([runner](rdocx_runner.py)), so its cold times include Python startup.

The other engines in the Word benchmark write only PDFs: EigenPal (`exportPdf`), MiniPdf
(`ConvertToPdf`), dxpdf (`convert` and `convert_file`) and office2pdf (PDF output only). They are
left out rather than measured through a PDF rasterizer.

## Measurement

- **Accuracy.** Each PNG is resampled to 96 dpi with a Lanczos filter, as the classic suite's
  150 dpi Word page images are, and transparent areas count as white paper. The Word benchmark's
  ink score then compares it with the reference page: the harmonic mean of how much ink each page
  reproduces from the other within one pixel, ignoring colour. Pages are paired by index, and
  missing or extra pages score zero. Page scores are averaged within each document, and each
  document has equal weight.
- **Page counts.** The page count is the number of PNGs a converter writes.
- **Cold speed and memory.** Each document is rendered in a fresh process, which includes runtime
  startup, reading the DOCX and writing every PNG. GNU time records the peak resident set size. For
  LibreOffice + pdftoppm, one shell runs both steps, and GNU time reports the larger of their peaks.
- **Warm speed.** Each engine renders an initial document once, then times every document in the
  same process. LibreOffice converts the documents in one CLI batch, timed between successive PDF
  completions as in the Word benchmark, and `pdftoppm` then rasterizes each PDF in its own process.
  A document's warm time is the sum of the two. A warm batch may run for 30 minutes.
  The 26 November 2026 Kanera warm figures are medians of four full-corpus runs, each in a fresh
  Node process.
- **Sampling, timeouts and failures** follow the Word benchmark. Each engine first renders the
  same 50-document sample. If that projects past five minutes for the corpus, the engine is timed
  on the sample only (†) and renders the other documents untimed for scoring. A conversion that
  fails, runs longer than five minutes, writes no pages or writes a PNG the scorer cannot read
  scores zero and counts as a failed conversion.
- **PNG size** is the total size of the cold-run PNGs across the corpus.

LibreOffice and Poppler ran as host programs launched from a container in our runs, so their
timings include that launch. When GNU time measures only the launcher, the README shows no
LibreOffice memory figures. Record the versions and fonts of each run, because a different
environment can change the scores.

## Run from the repository root

Set up the [Word benchmark](../word/README.md#run-from-the-repository-root) first: build the
package, create the Python environment and run `fetch.py`. Put `soffice` and `pdftoppm` on PATH.

Build the .NET runners with the .NET 10 SDK. The licences are the ones the Word benchmark uses:
`ASPOSE_WORDS_LICENSE` is the path of an Aspose.Words for .NET licence, and
`SYNCFUSION_LICENSE_KEY` is a key for Syncfusion version 28. The Syncfusion runner refuses a key
that does not validate, because an unlicensed run draws a trial banner on every page. For rdocx,
install the Word benchmark's Rust requirements.

```sh
dotnet build benchmarks/image/aspose/AsposeImageBenchmark.csproj -c Release \
  -p:RestoreLockedMode=true --artifacts-path out/benchmark-image/dotnet \
  -o out/benchmark-image/aspose-bin
dotnet build benchmarks/image/syncfusion/SyncfusionImageBenchmark.csproj -c Release \
  -p:RestoreLockedMode=true --artifacts-path out/benchmark-image/dotnet \
  -o out/benchmark-image/syncfusion-bin
out/benchmark-word/venv/bin/python -m pip install -r benchmarks/word/rust/requirements.txt

ASPOSE_WORDS_LICENSE="/absolute/path/to/Aspose.Words.lic" SYNCFUSION_LICENSE_KEY="your-v28-key" \
  out/benchmark-word/venv/bin/python benchmarks/image/convert_pages.py --all
out/benchmark-word/venv/bin/python benchmarks/image/score_pages.py --all
```

Use `--aspose`, `--syncfusion` and `--rdocx` in place of `--all` to choose engines; Kanera and
LibreOffice always run. Pass the same engine flags to both commands. Scoring a run of every
document updates the [results](#results) above and the summary in the main README. A run of every
engine rewrites both tables. A run of some engines replaces their rows and keeps the others.

`convert_pages.py --kanera-only` runs Kanera alone into `out/benchmark-image/kanera-only/`, and
`score_pages.py --kanera-only` scores it and refreshes the Kanera rows.

The results go in `out/benchmark-image/`. `report.md` has the full table, `results.csv` the
per-document and per-page scores, `conversion.json` the versions, timings, peak memory and errors,
and `gallery.html` side-by-side images of every page (`--no-gallery` skips them).

Caching, resuming and work directories work as in the Word benchmark. Each comparison engine's
results and pages are saved in `out/benchmark-image/cache/<engine>-<key>/` and reused while the
engine and runtime versions, runner files, inputs, references, CPU, resolution and timing settings
are unchanged. Kanera is always measured. `--no-cache` measures every engine again. A run that
stops resumes from `conversion-cold.json`, and `--fresh` starts again. Scores are kept in
`out/benchmark-image/score-cache/`, keyed by the PNGs' SHA-256 digests, the references, the tools
and the scoring code. Scoring uses up to eight processes (`--jobs`).

For a quick check, `--limit N` uses the first N documents of the timing sample order and `--cases`
takes a comma-separated list of case IDs. Use them with `--output out/<directory>` on both
commands. A run of fewer documents leaves the READMEs unchanged.
