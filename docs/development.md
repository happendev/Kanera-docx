# Development

```bash
pnpm install
pnpm test          # unit and end-to-end tests (uses Ghostscript, when installed, to read the PDFs back)
pnpm lint          # library, scripts, and browser demo
pnpm typecheck
pnpm build         # dist/

pnpm render in.docx --out out/ --png       # convert, and rasterise with Ghostscript
pnpm corpus list.txt --root ~/docs         # compare page counts and breaks with Word for many files
pnpm demo          # build the browser demo into out/demo (GitHub Pages uses main)
pnpm demo:serve    # serve it on http://localhost:8000 and rebuild on change
```

`list.txt` holds one `.docx` path per line (relative to `--root`). Set `BREAKS=1` to print where
each of Word's recorded page breaks landed. Keep private documents out of this repository.

## Releases

To prepare an npm release, run `pnpm lint`, `pnpm typecheck`, `pnpm test`, and `npm pack --dry-run`.
The manifest sets public access for `@kanera/docx2pdf`, and npm runs the build when packing
or publishing. The browser demo uses GitHub Actions as its Pages source.

## Code map

`src/docx` reads the package into a document model (styles, numbering, drawing),
`src/fonts` parses, substitutes and subsets fonts, `src/layout` breaks lines and paginates into
positioned drawing operations, and `src/pdf` writes those as PDF. `src/util` holds the ZIP reader,
compression adapters and XML parser.

## Shared core

kanera-docx, kanera-xlsx and kanera-pptx share a core that must stay identical in all three repos.
[AGENTS.md](../AGENTS.md) lists the core files and how to sync them with `pnpm core:status` and
`pnpm core:push`.

## Benchmarks

The [Word benchmark](../benchmarks/word/README.md) and the
[Markdown benchmark](../benchmarks/markdown/README.md) describe their setup and how to run them.
Scoring a run refreshes the tables in those READMEs and the summary tables in the main README.
