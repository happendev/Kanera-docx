# Command-line tool

The package installs a `docx2pdf` command.

Run the CLI directly from npm. This writes `in.pdf` next to `in.docx`.

```bash
npx @kanera/docx2pdf in.docx
```

In a project with the package installed, use `npx --no-install docx2pdf`. Pass an output path or
options as needed.

```bash
npx --no-install docx2pdf in.docx out.pdf --font BrandSans.ttf --system-fonts
```

## Page images

To get page images instead of a PDF, pass `--png` or `--jpeg` with a directory. This writes
`pages/in-1.png`, `pages/in-2.png` and so on, with page numbers zero-padded to the same width so
they sort in order. `--dpi` sets the resolution (default 150), `--pages` picks pages such as `1-3,5`,
and `--quality` sets the JPEG quality (default 90).

```bash
npx --no-install docx2pdf in.docx --png pages --dpi 200
```

## Text and Markdown

To extract the text instead, pass `--text` or `--markdown` (`--md`). This writes `in.txt` or
`in.md`, or prints to standard output when the output is `-`. Add `--headers` to include headers
and footers, or `--no-notes` to leave out footnotes and endnotes.

```bash
npx --no-install docx2pdf in.docx --markdown -
```

## Every flag

This is the output of `docx2pdf --help`:

```text
Usage: docx2pdf <input.docx> [output.pdf] [options]
       docx2pdf <input.docx> --png|--jpeg <dir> [--dpi <n>] [--pages <list>] [options]
       docx2pdf <input.docx> --text|--markdown [output|-] [--headers] [--no-notes]

Options:
  --font <file>      Use this TTF/OTF/TTC font (repeatable)
  --system-fonts     Also use fonts installed on this machine
  --mac              Match Word for Mac's PDFs (text on a 0.24pt grid)
  --image-quality <1-100|off>  JPEG quality (default: automatic selection up to 4 MP at quality 80)
  --png <dir>        Write each page as <dir>/<input>-<n>.png instead of a PDF
  --jpeg <dir>       Write each page as <dir>/<input>-<n>.jpg instead of a PDF
  --dpi <n>          Resolution of page images (default: 150)
  --quality <1-100>  Quality of --jpeg images (default: 90)
  --pages <list>     Pages to write as images, such as 1-3,5 (default: all)
  --text             Write the document's plain text (<input>.txt, or - for stdout)
  --markdown, --md   Write the document as Markdown (<input>.md, or - for stdout)
  --headers          With --text or --markdown, include headers and footers
  --no-notes         With --text or --markdown, leave out footnotes and endnotes
  --quiet            Do not print warnings
  -h, --help         Show this help
```

The [options](options.md) describe what each setting does.
