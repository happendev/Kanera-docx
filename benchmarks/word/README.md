# DOCX-to-PDF benchmark

The benchmark runs 278 cases by default. It downloads the [100 pinned Hugging Face pairs](fixtures.json)
with Microsoft Word for Mac references and uses all [178 vendored cases](upstream/fixtures.json).
Both sets are verified against SHA-256 hashes. See the [Hugging Face corpus](https://huggingface.co/datasets/arthrod/neurotic_docx_bench-fixtures)
for attribution and notices about individual files.

The Hugging Face sample combines 39 documents from `sources_500`, 26 from `en_pairs_500`, and
35 focused fixtures. The source documents cover text, tables, images, and combinations of those
elements. We selected by Word reference page count and DOCX features without using converter scores.

The sample leaves out two kinds of document whose Word reference shows something docx2pdf
deliberately doesn't draw:

- **Tracked changes and comments.** Word exported these in markup view, with struck-through
  deletions, coloured insertions, and comment balloons beside a shrunken page. docx2pdf renders the
  final text without comments, so only pairs in the corpus's `clean` state are used.
- **Scripts without a bundled font.** Thai, Devanagari, Khmer, and Thaana text needs a font the user
  supplies.

Each excluded document was replaced by the clean pair with the lowest corpus ID from the same
`sets` group and page-count range below, or from `sources_500` when that group had none left.

| Word reference pages | Documents |
| --- | ---: |
| 1 | 37 |
| 2–5 | 39 |
| 6–15 | 17 |
| 16–29 | 7 |

The sample includes mixed page orientations, columns, charts, diagrams, equations, notes, and
Arabic, Cyrillic, Greek, Latin, and East Asian text.

## Run from the repository root

Use Linux with GNU time at `/usr/bin/time`, Node 24 with npm and pnpm, and Python 3.11–3.13 with
venv. Put headless LibreOffice (`soffice`) and Poppler (`pdftoppm`) on PATH. Use a licensed
installation of EigenPal where required by its terms. The npm and Python dependencies below are
only for the benchmark.

```sh
pnpm install --frozen-lockfile
pnpm build
npm ci --prefix benchmarks/word --ignore-scripts

python3 -m venv out/benchmark-word/venv
out/benchmark-word/venv/bin/python -m pip install -r benchmarks/word/requirements.txt
out/benchmark-word/venv/bin/python benchmarks/word/fetch.py
out/benchmark-word/venv/bin/python benchmarks/word/convert.py
out/benchmark-word/venv/bin/python benchmarks/word/score.py
```

Scoring a completed full comparison updates the benchmark summary near the top of the main README,
its comparison table, and the breakdown below. The breakdown shows the recorded run date in UTC.
Scoring uses up to eight processes by default. Use `score.py --jobs 1` to score one case at a
time, or set a different worker count with `--jobs`. Cold and warm conversion timings are serial.

If `convert.py` exits with an error during the warm pass, run it again with the same engine flags
and build. It reuses the cold timings and PDFs from `conversion-cold.json` and reruns the warm
pass. Use `--fresh` to start from cold conversions. Run `score.py` after conversion finishes
successfully.

### Fast Kanera-only run

Build the current checkout with `pnpm build` and prepare the fixtures and Python environment with
the steps above. The Kanera-only run uses all 278 fixtures and needs Node, GNU time, and Poppler.

```sh
out/benchmark-word/venv/bin/python benchmarks/word/convert.py --kanera-only
out/benchmark-word/venv/bin/python benchmarks/word/score.py --kanera-only
```

The results go in `out/benchmark-word/kanera-only/`. Add `--gallery` to the scoring command to
generate the side-by-side screenshots. The full comparison generates a gallery by default and
accepts `--no-gallery` for a report and CSV without screenshots.
Scoring a Kanera-only run refreshes the Kanera rows in the main README. The breakdown here
reflects the last full comparison.

### Compare bundled CJK font loading

After building the package and fetching the fixtures, compare the default Node loader (`sparse`)
with full reads of the same fonts using an explicit `fontDir` (`full`). For the packaged Noto CJK
faces, the default loader caches layout tables and reads the required outline ranges asynchronously
at PDF embedding time. Documents with equations use full fonts so layout can measure glyph bounds.

```sh
fixture=out/benchmark-word/cases/upstream-nthu_article/input.docx
mkdir -p out/cjk-font-comparison
for mode in full sparse; do
  /usr/bin/time -f '%M KiB peak RSS, %e s elapsed' \
    node benchmarks/word/cjk-fonts.mjs "$mode" "$fixture" "out/cjk-font-comparison/$mode.pdf"
done
cmp out/cjk-font-comparison/full.pdf out/cjk-font-comparison/sparse.pdf
```

The worker also accepts `--batch jobs.json` in place of the input and output paths. Each job is
an `[inputPath, outputPath]` pair. It emits times in milliseconds in input order.

On 2026-11-20 with Node 24.18.0, all 278 cold PDFs and all 278 PDFs in each of three warm batches
matched the pre-change Node build byte for byte. With the automatic JPEG cap active in both builds,
the maximum cold peak RSS fell from 230.4 to 200.1 MiB. The median cold peak rounded to 84.1 MiB
in both builds.

### Include Aspose.Words for .NET

Install the .NET 10 SDK. Obtain a [temporary Aspose.Words for .NET license](https://purchase.aspose.com/pricing/words/net)
or use an existing license for the pinned version. Keep the license outside the checkout. On
Linux, the benchmark includes the SkiaSharp native assets needed to render graphics.

After installing dependencies and downloading the fixtures, build the runner and use `--aspose`
on both commands to include it in the run.

```sh
dotnet build benchmarks/word/aspose/Benchmark.csproj -c Release \
  -p:RestoreLockedMode=true --artifacts-path out/benchmark-word/dotnet \
  -o out/benchmark-word/aspose-bin

ASPOSE_WORDS_LICENSE="/absolute/path/to/Aspose.Words.lic" \
  out/benchmark-word/venv/bin/python benchmarks/word/convert.py --aspose

out/benchmark-word/venv/bin/python benchmarks/word/score.py --aspose
```

### Include Syncfusion DocIO

Syncfusion license keys only work with their major version, so the runner pins DocIO 28.2.4. Use a
key for version 28; an unlicensed run adds a trial watermark, which the scorer rejects. Build the
runner, then add `--syncfusion` to both commands (it combines with `--aspose`).

```sh
dotnet build benchmarks/word/syncfusion/SyncfusionBenchmark.csproj -c Release \
  -p:RestoreLockedMode=true --artifacts-path out/benchmark-word/dotnet \
  -o out/benchmark-word/syncfusion-bin

SYNCFUSION_LICENSE_KEY="your-v28-key" \
  out/benchmark-word/venv/bin/python benchmarks/word/convert.py --syncfusion

out/benchmark-word/venv/bin/python benchmarks/word/score.py --syncfusion
```

### Include MiniPdf (.NET)

The runner uses [MiniPdf 0.43.2](https://www.nuget.org/packages/MiniPdf/0.43.2) on .NET 10,
with host fonts and PDF content-stream compression enabled.

```sh
dotnet build benchmarks/word/minipdf/MiniPdfBenchmark.csproj -c Release \
  -p:RestoreLockedMode=true --artifacts-path out/benchmark-word/dotnet \
  -o out/benchmark-word/minipdf-bin

out/benchmark-word/venv/bin/python benchmarks/word/convert.py --minipdf
out/benchmark-word/venv/bin/python benchmarks/word/score.py --minipdf
```

Use `--minipdf-only` on both commands to run it on its own and refresh its README results.
Standalone results are saved under `out/benchmark-word/minipdf-only/`.

### Include the Rust engines

Three open-source Rust converters can join the run:

- [dxpdf](https://github.com/nerdy-pro/dxpdf) renders with Skia.
- [rdocx](https://github.com/tensorbee/rdocx) is pure Rust.
- [office2pdf](https://github.com/developer0hye/office2pdf) is pure Rust and lays out with Typst.

The benchmark uses their prebuilt releases, so no Rust toolchain is needed. dxpdf and rdocx are
installed as Python wheels that bundle their native libraries.
[`rust/runner.py`](rust/runner.py) calls them through those bindings, so their cold times include
Python startup. office2pdf runs as its statically linked CLI. The installer downloads the pinned
release and verifies its SHA-256 hash.

```sh
out/benchmark-word/venv/bin/python -m pip install -r benchmarks/word/rust/requirements.txt
out/benchmark-word/venv/bin/python benchmarks/word/rust/install.py

out/benchmark-word/venv/bin/python benchmarks/word/convert.py --rust
out/benchmark-word/venv/bin/python benchmarks/word/score.py --rust
```

`--rust`, `--aspose`, `--syncfusion`, and `--minipdf` can be combined.

For every engine, a failed conversion, or one that runs longer than five minutes, does not stop
the run. The failure is recorded in `conversion.json`, and that document scores zero. A PDF that
Poppler cannot rasterize, or whose text extraction crashes or hangs, is scored the same way: the
scorer skips it, counts it as a failed conversion, and names the reason in the `unreadable_pdf`
column of `results.csv`. The report
counts failed conversions for every engine. Speed and memory figures cover successful conversions
only.

## Latest results

Ink overlap by suite for the [benchmark comparison](../../README.md#benchmark).

<!-- benchmark-breakdown:start -->
**Last full benchmark run** 2026-11-20 (UTC), 278 documents.

| Cases | Reference | @kanera/docx2pdf | EigenPal | LibreOffice | Aspose.Words (.NET) | Syncfusion DocIO (.NET) | dxpdf (Rust) | rdocx (Rust) | office2pdf (Rust) | MiniPdf (.NET) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 150 classic | Microsoft 365 Word page images | **99.1%** | 74.0% | 91.2% | 51.0% | 58.4% | 36.9% | 44.8% | 57.1% | 55.9% |
| 100 Hugging Face | Word for Mac PDFs | **83.2%** | 66.9% | 70.3% | 54.4% | 54.2% | 37.7% | 39.4% | 40.9% | 28.8% |
| 20 issues | Microsoft 365 Word PDFs | **90.5%** | 71.1% | 53.7% | 49.8% | 51.9% | 31.2% | 18.7% | 32.3% | 31.4% |
| 8 issues | LibreOffice PDFs | 58.6% | 62.9% | **68.1%** | 49.4% | 48.6% | 24.7% | 26.8% | 42.4% | 36.7% |

Per-document ink overlap differences under half a percentage point count as ties.

| Compared with | Kanera higher | Kanera lower | Ties |
| --- | ---: | ---: | ---: |
| EigenPal | 216 | 26 | 36 |
| LibreOffice | 163 | 25 | 90 |
| Aspose.Words (.NET) | 252 | 15 | 11 |
| Syncfusion DocIO (.NET) | 252 | 13 | 13 |
| dxpdf (Rust) | 269 | 6 | 3 |
| rdocx (Rust) | 242 | 3 | 33 |
| office2pdf (Rust) | 270 | 5 | 3 |
| MiniPdf (.NET) | 270 | 5 | 3 |

MiniPdf results come from a standalone run.
<!-- benchmark-breakdown:end -->

On 2026-11-20, LibreOffice ran as a Flatpak on the host, launched from a container. The timings
include that launch, and GNU time measured the launcher’s memory. LibreOffice used the host’s
Microsoft fonts, including Arial, Aptos and Times New Roman. Aspose and Syncfusion fell back to
Liberation and DejaVu.

## Measurement

- **Cold speed.** The benchmark launches a fresh process for each converter and document.
  Wall time includes runtime and library startup, reading the DOCX, and writing the PDF. LibreOffice
  uses a separate profile for each invocation. The host's filesystem caches may already be warm.
- **Warm speed.** Each engine converts an initial fixture once before timing the full set in the
  same process. Timings still include reading the DOCX and writing the PDF. LibreOffice gets the
  documents in one CLI batch. Its per-file timing is the interval between successive PDF write
  completions, observed with Linux inotify. office2pdf is timed the same way. If such a batch
  exits partway, the document it stopped on gets no warm time and a new batch resumes after it.
  The report includes median and p95 per-document
  latency. Warm docs/sec is the number of successful timed conversions divided by the sum of
  their durations, excluding the warm-up conversion.
- **Memory.** GNU time records peak resident set size (RSS) for each cold invocation, including runtime
  overhead. The report shows the median and maximum across files. When a command spawns child
  processes, GNU time reports the largest process peak. Times and memory vary by host and fonts.
- **PDF file size.** The report and gallery sum the cold-run PDF sizes across the corpus for each
  converter. `results.csv` records each converted PDF's size in bytes.
- **Layout.** Poppler rasterizes the cold-run output at 96 DPI. Pages are paired by their index. The ink score is the
  harmonic mean of how much ink each page reproduces from the other within one pixel. It ignores
  colour differences. Missing or extra pages score zero. Page scores are averaged within each
  document, and each document has equal weight in the summary.
- **Text and page counts.** The CSV includes page counts and the reference application and format.
  For PDF references it also records the percentage of reference words found on the same page
  and the median position offset for tokens that occur once on that page in both outputs.
  PDF extraction can tokenize or order text differently across engines.
- **Settings.** Our renderer uses the full font pack. EigenPal uses its
  packaged fonts with the optional CJK package, proposed revision mode, comments off, and
  best-effort fidelity. LibreOffice uses local fonts. Aspose uses its default
  PDF export settings and font lookup. Syncfusion DocIO converts with its default `DocIORenderer`
  settings and local fonts. dxpdf and office2pdf use their default settings and local fonts.
  rdocx uses its default settings with system fonts enabled, because by default it only uses its
  bundled fallback fonts. Record your LibreOffice and Poppler versions and font
  environment when comparing runs. A different environment can change visual scores.

Open `out/benchmark-word/gallery.html` to scan the screenshots. Each PNG shows the reference followed by
the converters in one row, using a common scale for that page. The generated `report.md` links to
every page and lists every case. `results.csv` contains the per-file scores and `conversion.json`
records versions, timing, peak memory, and diagnostics. The downloaded benchmark fixtures,
converted PDFs, images, results, and build outputs are Git-ignored.

## Vendored fixtures

- **Issue suite.** All 28 cases in the [published issue coverage](https://api.github.com/repositories/1374445024/contents/assets/benchmarks/dotnet/issue/docx/benchmark_coverage.json?ref=96c015dffad995f4e03ab07b25c073d424ddb1c5)
  are vendored as DOCX/PDF pairs from the [pinned source revision](https://api.github.com/repositories/1163851617/contents/tests/Issue_Files?ref=7a775fb6469ce7bb1322e6fe9d0c37edfaf67afa).
  The checked-in references include 20 Word exports and 8 LibreOffice exports.
- **Classic suite.** The pinned [fixture generator](https://api.github.com/repositories/1163851617/git/blobs/ef3dda01bdd13298363c5602bb9539bf2f11a8e2)
  produces 150 cases using python-docx 1.2.0 and Pillow 13.0.0. We vendor those inputs with the
  corresponding published 150 DPI Microsoft 365 Word page images and resample them to 96 DPI for
  scoring. The [published coverage](https://api.github.com/repositories/1374445024/contents/assets/benchmarks/dotnet/classic/docx/benchmark_coverage.json?ref=96c015dffad995f4e03ab07b25c073d424ddb1c5)
  lists another 30 image-heavy cases whose source DOCX files are unavailable in the public repository.

## Full corpus accuracy

Run this accuracy check from the repository root on Linux. It builds the current checkout and
scores 2,650 DOCX files from two sources.

- **Hugging Face.** 2,472 of the 2,925 pairs in `word/documents.csv` at the pinned revision of the
  public corpus, with Microsoft Word for Mac references. The runner downloads these into a local
  cache. It leaves out the 447 pairs with tracked changes or comments and six pairs in scripts
  without a bundled font, for the reasons the [sample](#docx-to-pdf-benchmark) gives.
- **Vendored fixtures.** The same 178 [issue and classic cases](#vendored-fixtures) as the benchmark.

The runner uses the same 96 DPI ink score as the benchmark, averaging pages within each document
and giving each document equal weight. Missing and extra pages score zero. A failed conversion
scores zero, is listed in the report, and makes the command exit with status 1.
A completed run of both corpora refreshes the full-corpus ink score and run date in the main README.
Each worker reuses a Node process across documents and replaces it after a conversion error or
timeout. Output pages are rasterized as temporary raw RGB images to avoid PNG compression.
Rendered reference pages are cached as lossless grayscale PNGs under `out/word-reference-rasters/`.
The cache key includes the reference digest, raster settings, and rendering environment. Use
`--reference-cache-dir` to store the rendered pages elsewhere.

Use Python 3.11–3.13, pnpm and Node, and put Poppler (`pdftoppm`) on PATH. Set up the Python
environment once, then run the accuracy check before and after a renderer change.

```sh
pnpm install --frozen-lockfile
python3 -m venv out/word-accuracy-venv
out/word-accuracy-venv/bin/python -m pip install -r benchmarks/word/requirements.txt

out/word-accuracy-venv/bin/python benchmarks/word/accuracy.py --out out/accuracy-before.json
out/word-accuracy-venv/bin/python benchmarks/word/accuracy.py \
  --baseline out/accuracy-before.json --out out/accuracy-after.json
```

Use `--corpus upstream` to run only the vendored cases, or `--corpus huggingface` to run only
the Hugging Face pairs and compare with an earlier Hugging Face-only report.

The first Hugging Face run downloads missing pairs into the system temporary directory and
waits for any rate-limit windows to reset. Later runs verify SHA-256 hashes and reuse the files.
Use `--cache-dir` to point to the `word/` directory of an existing Hugging Face snapshot or to
choose where to cache the download. The JSON report includes overall and per-corpus accuracy,
per-document scores and reference applications, page counts, and tool versions. Use the same
corpus selection, tool versions and font environment when comparing runs.
