#!/usr/bin/env node
import { readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join } from 'node:path';
import { convertDocx } from './node.js';
import type { FontSource } from './core.js';

const USAGE = `Usage: docx2pdf <input.docx> [output.pdf] [options]

Options:
  --font <file>      Use this TTF/OTF/TTC font (repeatable)
  --system-fonts     Also use fonts installed on this machine
  --mac              Match Word for Mac's PDFs (text on a 0.24pt grid)
  --image-quality <1-100|off>  JPEG quality (default: automatic 80)
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
  const result = await convertDocx(input, { fonts, systemFonts, imageQuality, wordPlatform: mac ? 'mac' : 'windows' });
  await writeFile(output, result.pdf);
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
