// Compare the production Node loader with full reads of the same font files.
import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { convertDocx } from '../../dist/node.js';

const [mode, input, output] = process.argv.slice(2);
if (!['full', 'sparse'].includes(mode) || !input || !output) {
  throw new Error('Usage: node cjk-fonts.mjs <full|sparse> <input.docx> <output.pdf> or --batch <jobs.json>');
}
const options = mode === 'full' ? { fontDir: fileURLToPath(new URL('../../fonts/', import.meta.url)) } : {};

async function convert(source, dest) {
  const result = await convertDocx(await readFile(source), options);
  await writeFile(dest, result.pdf);
  return { pageCount: result.pageCount, warnings: result.warnings, stats: result.stats };
}

if (input === '--batch') {
  const jobs = JSON.parse(await readFile(output, 'utf8'));
  const times = [];
  for (const [source, dest] of jobs) {
    const start = performance.now();
    await convert(source, dest);
    times.push(performance.now() - start);
  }
  console.log(JSON.stringify(times));
} else {
  console.log(JSON.stringify(await convert(input, output)));
}
