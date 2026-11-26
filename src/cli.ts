#!/usr/bin/env node
import { readInputFile, readInputPath } from './util/node-input.js';
import { mkdir, open, readFile, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { basename, dirname, extname, join } from 'node:path';
import type { FontSource } from './core.js';
import { extractTextWithCompression } from './text/index.js';
import { nodeCompression } from './util/node-compression.js';

const USAGE = `Usage: docx2pdf <input.docx> [output.pdf] [options]
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
  -h, --help         Show this help`;

async function main(argv: string[]): Promise<number> {
  const files: string[] = [];
  const fonts: FontSource[] = [];
  let systemFonts = false;
  let quiet = false;
  let mac = false;
  let imageQuality: number | false | undefined;
  let imageDir: string | undefined;
  let imageFormat: 'png' | 'jpeg' = 'png';
  let dpi: number | undefined;
  let quality: number | undefined;
  let pageList: string | undefined;
  let textFormat: 'text' | 'markdown' | undefined;
  let includeHeadersFooters = false;
  let notes = true;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '-h' || a === '--help') {
      console.log(USAGE);
      return 0;
    }
    if (a === '--font') {
      const file = argv[++i];
      if (!file || file.startsWith('-')) {
        console.error('--font requires a filename');
        return 2;
      }
      fonts.push({ data: await readFile(file) });
    } else if (a === '--system-fonts') systemFonts = true;
    else if (a === '--image-quality') {
      const value = argv[++i];
      if (value === 'off') imageQuality = false;
      else if (value && /^[0-9]+$/.test(value) && Number(value) >= 1 && Number(value) <= 100) imageQuality = Number(value);
      else {
        console.error('--image-quality requires an integer from 1 to 100 or off');
        return 2;
      }
    } else if (a === '--png' || a === '--jpeg') {
      if (imageDir !== undefined) {
        console.error('Pass only one of --png and --jpeg');
        return 2;
      }
      imageDir = argv[++i];
      imageFormat = a === '--png' ? 'png' : 'jpeg';
      if (!imageDir) {
        console.error(`${a} requires a directory`);
        return 2;
      }
    } else if (a === '--quality') {
      const value = argv[++i];
      if (value && /^[0-9]+$/.test(value) && Number(value) >= 1 && Number(value) <= 100) quality = Number(value);
      else {
        console.error('--quality requires an integer from 1 to 100');
        return 2;
      }
    } else if (a === '--pages') {
      pageList = argv[++i];
      if (!pageList || !/^\d+(-\d+)?(,\d+(-\d+)?)*$/.test(pageList)) {
        console.error('--pages requires page numbers or ranges, such as 1-3,5');
        return 2;
      }
    } else if (a === '--dpi') {
      const value = argv[++i];
      if (value && /^[0-9]+$/.test(value) && Number(value) >= 1 && Number(value) <= 2400) dpi = Number(value);
      else {
        console.error('--dpi requires an integer from 1 to 2400');
        return 2;
      }
    } else if (a === '--text') textFormat = 'text';
    else if (a === '--markdown' || a === '--md') textFormat = 'markdown';
    else if (a === '--headers') includeHeadersFooters = true;
    else if (a === '--no-notes') notes = false;
    else if (a === '--quiet') quiet = true;
    else if (a === '--mac') mac = true;
    else if (a === '-' && files.length === 1) files.push(a);
    else if (a.startsWith('-')) {
      console.error(`Unknown option ${a}\n\n${USAGE}`);
      return 2;
    } else files.push(a);
  }
  if (!files.length || files.length > 2) {
    console.error(USAGE);
    return 2;
  }
  if ((dpi !== undefined || pageList !== undefined) && imageDir === undefined) {
    console.error('--dpi and --pages only apply with --png or --jpeg');
    return 2;
  }
  if (quality !== undefined && imageFormat !== 'jpeg') {
    console.error('--quality only applies with --jpeg');
    return 2;
  }
  const input = files[0];
  if (textFormat) {
    if (imageDir !== undefined || fonts.length || systemFonts || mac || imageQuality !== undefined) {
      console.error('--text and --markdown take only --headers, --no-notes and --quiet');
      return 2;
    }
    const result = extractTextWithCompression(await readInputPath(input), { format: textFormat, includeHeadersFooters, includeNotes: notes }, nodeCompression);
    if (!quiet) for (const w of result.warnings) console.error(`warning: ${w}`);
    if (files[1] === '-') {
      process.stdout.write(result.text);
      return 0;
    }
    const output = files[1] ?? join(dirname(input), basename(input, extname(input)) + (textFormat === 'text' ? '.txt' : '.md'));
    await writeText(input, output, result.text);
    console.log(output);
    return 0;
  }
  if (files[1] === '-') {
    console.error('Writing to standard output (-) only applies with --text or --markdown');
    return 2;
  }
  if (includeHeadersFooters || !notes) {
    console.error('--headers and --no-notes only apply with --text or --markdown');
    return 2;
  }
  // The renderer loads only for PDF and image output, so text extraction starts faster.
  const { convertDocx, convertDocxToImages } = await import('./node.js');
  const options = { fonts, systemFonts, imageQuality, wordPlatform: mac ? 'mac' : 'windows' } as const;
  if (imageDir !== undefined) {
    if (files.length > 1) {
      console.error(`Pass either an output PDF or --${imageFormat} <dir>, not both`);
      return 2;
    }
    if (imageQuality !== undefined) {
      console.error('--image-quality only applies to PDF output; use --quality with --jpeg');
      return 2;
    }
    let pages: number[] | undefined;
    if (pageList !== undefined) {
      const chosen = new Set<number>();
      for (const part of pageList.split(',')) {
        const [from, to = from] = part.split('-').map(Number);
        if (from < 1 || to < from || to > 100_000) {
          console.error(`--pages: ${part} is not a page range`);
          return 2;
        }
        for (let n = from; n <= to; n++) chosen.add(n);
      }
      pages = [...chosen].sort((a, b) => a - b);
    }
    const prefix = basename(input, extname(input));
    const ext = imageFormat === 'png' ? 'png' : 'jpg';
    // Each image is written as soon as it is drawn, so only the pages in flight stay in memory.
    const onPage = async (image: { page: number; data: Uint8Array }, pageCount: number) => {
      await mkdir(imageDir, { recursive: true });
      await writeFile(join(imageDir, `${prefix}-${String(image.page).padStart(String(pageCount).length, '0')}.${ext}`), image.data);
    };
    const result = await convertDocxToImages(input, { fonts, systemFonts, wordPlatform: options.wordPlatform, format: imageFormat, dpi, quality, pages, onPage });
    await mkdir(imageDir, { recursive: true });
    if (!quiet) for (const w of result.warnings) console.error(`warning: ${w}`);
    const n = result.images.length;
    console.log(`${imageDir}: ${n} page image${n === 1 ? '' : 's'}`);
    return 0;
  }
  const output = files[1] ?? join(dirname(input), basename(input, extname(input)) + '.pdf');
  const source = await open(input, 'r');
  let result;
  try {
    const sourceStat = await source.stat({ bigint: true });
    result = await convertDocx(await readInputFile(source), options);
    // Compare the open files before truncating, including symlinks and hard links.
    const target = await open(output, constants.O_WRONLY | constants.O_CREAT);
    try {
      const targetStat = await target.stat({ bigint: true });
      if (sourceStat.dev === targetStat.dev && sourceStat.ino === targetStat.ino) throw new Error('Input and output refer to the same file');
      await target.truncate();
      await target.writeFile(result.pdf);
    } finally {
      await target.close();
    }
  } finally {
    await source.close();
  }
  if (!quiet) for (const w of result.warnings) console.error(`warning: ${w}`);
  console.log(`${output}: ${result.pageCount} page${result.pageCount === 1 ? '' : 's'}`);
  return 0;
}

/** Write text to a file, refusing to overwrite the input (through a link or otherwise). */
async function writeText(input: string, output: string, text: string): Promise<void> {
  const source = await open(input, 'r');
  try {
    const sourceStat = await source.stat({ bigint: true });
    const target = await open(output, constants.O_WRONLY | constants.O_CREAT);
    try {
      const targetStat = await target.stat({ bigint: true });
      if (sourceStat.dev === targetStat.dev && sourceStat.ino === targetStat.ino) throw new Error('Input and output refer to the same file');
      await target.truncate();
      await target.writeFile(text);
    } finally {
      await target.close();
    }
  } finally {
    await source.close();
  }
}

main(process.argv.slice(2)).then(
  (code) => process.exit(code),
  (err) => {
    console.error(err instanceof Error ? err.message : err);
    process.exit(1);
  },
);
