/**
 * Convert .docx files and optionally rasterise the PDFs with Ghostscript for visual checks.
 *   pnpm render <file.docx>... [--out dir] [--png] [--dpi 72] [--system-fonts]
 */
import { execFileSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { basename, extname, join } from 'node:path';
import { convertDocx } from '../src/node.js';

const args = process.argv.slice(2);
const files: string[] = [];
let out = 'out';
let png = false;
let dpi = 72;
let systemFonts = false;
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--out') out = args[++i];
  else if (args[i] === '--png') png = true;
  else if (args[i] === '--dpi') dpi = Number(args[++i]);
  else if (args[i] === '--system-fonts') systemFonts = true;
  else files.push(args[i]);
}
mkdirSync(out, { recursive: true });
for (const f of files) {
  const name = basename(f, extname(f));
  const t0 = performance.now();
  const r = await convertDocx(f, { systemFonts });
  const ms = Math.round(performance.now() - t0);
  const pdfPath = join(out, `${name}.pdf`);
  writeFileSync(pdfPath, r.pdf);
  console.log(`${pdfPath}: ${r.pageCount} pages, ${ms} ms, breaks ${r.stats.matchedBreaks}/${r.stats.renderedBreaks}`);
  for (const w of r.warnings) console.log(`  warning: ${w}`);
  if (process.env.BREAKS) console.log('  breaks:', r.stats.breakLog.map((b, i) => `#${i + 1}[${b.block}]:p${b.page + 1}${b.top ? '' : '@' + b.y}`).join(' '));
  if (png) {
    execFileSync('gs', ['-q', '-dNOPAUSE', '-dBATCH', '-dSAFER', '-sDEVICE=png16m', `-r${dpi}`, '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4', `-sOutputFile=${join(out, `${name}-%02d.png`)}`, pdfPath], { stdio: 'inherit' });
  }
}
