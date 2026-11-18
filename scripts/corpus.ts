/**
 * Accuracy check over a local corpus of real Word documents (never committed).
 *   pnpm corpus <list.txt> [--root dir] [--out dir]
 * Compares page counts with Word's (docProps/app.xml) and our page breaks with the breaks Word
 * recorded when it last saved (w:lastRenderedPageBreak).
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import { ZipArchive, decodeXmlText } from '../src/util/zip.js';
import { convertDocx } from '../src/node.js';

const args = process.argv.slice(2);
let root = '.';
let out = '';
const lists: string[] = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--root') root = args[++i];
  else if (args[i] === '--out') out = args[++i];
  else lists.push(args[i]);
}
const files = lists.flatMap((l) => readFileSync(l, 'utf8').split('\n').map((s) => s.trim()).filter(Boolean));
if (out) mkdirSync(out, { recursive: true });

let ok = 0;
let failed = 0;
let pagesMatch = 0;
let pagesKnown = 0;
let breaks = 0;
let matched = 0;
const rows: string[] = [];
for (const rel of files) {
  const path = join(root, rel);
  const data = readFileSync(path);
  let wordPages: number | null = null;
  try {
    const zip = new ZipArchive(new Uint8Array(data));
    const app = zip.read('docProps/app.xml');
    const m = app ? /<Pages>(\d+)<\/Pages>/.exec(decodeXmlText(app)) : null;
    if (m) wordPages = Number(m[1]);
  } catch {
    // not a zip
  }
  const t0 = performance.now();
  try {
    const r = await convertDocx(new Uint8Array(data));
    const ms = Math.round(performance.now() - t0);
    ok++;
    breaks += r.stats.renderedBreaks;
    matched += r.stats.matchedBreaks;
    if (wordPages !== null) {
      pagesKnown++;
      if (wordPages === r.pageCount) pagesMatch++;
    }
    const id = basename(rel).slice(0, 8);
    const missing = r.warnings.filter((w) => w.startsWith('Font "')).map((w) => /Font "([^"]+)"/.exec(w)![1]);
    rows.push(`${id}  pages ${String(r.pageCount).padStart(3)} / word ${String(wordPages ?? '?').padStart(3)}  breaks ${String(r.stats.matchedBreaks).padStart(2)}/${String(r.stats.renderedBreaks).padEnd(2)}  ${String(ms).padStart(4)}ms  ${missing.length ? 'missing: ' + missing.join(', ') : ''}`);
    if (process.env.BREAKS && r.stats.renderedBreaks) rows.push('    ' + r.stats.breakLog.map((b, i) => `#${i + 1}:p${b.page + 1}${b.top ? '' : '@' + b.y}`).join(' '));
    if (out) writeFileSync(join(out, `${id}.pdf`), r.pdf);
  } catch (e) {
    failed++;
    rows.push(`${basename(rel).slice(0, 8)}  FAILED ${(e as Error).stack?.split('\n').slice(0, 3).join(' | ')}`);
  }
}
console.log(rows.join('\n'));
console.log(`\n${ok} converted, ${failed} failed; page count matches Word ${pagesMatch}/${pagesKnown}; recorded breaks reproduced ${matched}/${breaks} (${breaks ? ((100 * matched) / breaks).toFixed(1) : '-'}%)`);
