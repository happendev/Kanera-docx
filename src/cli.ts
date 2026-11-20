#!/usr/bin/env node
import { open, readFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { basename, dirname, extname, join } from 'node:path';
import { convertDocx } from './node.js';
import type { FontSource } from './core.js';

const USAGE = `Usage: docx2pdf <input.docx> [output.pdf] [options]

Options:
  --font <file>      Use this TTF/OTF/TTC font (repeatable)
  --system-fonts     Also use fonts installed on this machine
  --mac              Match Word for Mac's PDFs (text on a 0.24pt grid)
  --image-quality <1-100|off>  JPEG quality (default: automatic selection at quality 80)
  --quiet            Do not print warnings
  -h, --help         Show this help`;

async function main(argv: string[]): Promise<number> {
  const files: string[] = [];
  const fonts: FontSource[] = [];
  let systemFonts = false;
  let quiet = false;
  let mac = false;
  let imageQuality: number | false | undefined;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '-h' || a === '--help') {
      console.log(USAGE);
      return 0;
    }
    if (a === '--font') fonts.push({ data: await readFile(argv[++i]) });
    else if (a === '--system-fonts') systemFonts = true;
    else if (a === '--image-quality') {
      const value = argv[++i];
      if (value === 'off') imageQuality = false;
      else if (value && /^[0-9]+$/.test(value) && Number(value) >= 1 && Number(value) <= 100) imageQuality = Number(value);
      else {
        console.error('--image-quality requires an integer from 1 to 100 or off');
        return 2;
      }
    } else if (a === '--quiet') quiet = true;
    else if (a === '--mac') mac = true;
    else if (a.startsWith('-')) {
      console.error(`Unknown option ${a}\n\n${USAGE}`);
      return 2;
    } else files.push(a);
  }
  if (!files.length) {
    console.error(USAGE);
    return 2;
  }
  const input = files[0];
  const output = files[1] ?? join(dirname(input), basename(input, extname(input)) + '.pdf');
  const source = await open(input, 'r');
  let result;
  try {
    const sourceStat = await source.stat({ bigint: true });
    result = await convertDocx(await source.readFile(), { fonts, systemFonts, imageQuality, wordPlatform: mac ? 'mac' : 'windows' });
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

main(process.argv.slice(2)).then(
  (code) => process.exit(code),
  (err) => {
    console.error(err instanceof Error ? err.message : err);
    process.exit(1);
  },
);
