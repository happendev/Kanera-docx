// Warm page-image timing for Kanera: one Node process renders every job in order.
// Usage: node kanera.mjs --batch <jobs.json> <dpi>, where each job is [input.docx, output directory].
// Each job writes <output directory>/page-<n>.png with the options the docx2pdf CLI uses for --png.
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { convertDocxToImages } from '../../dist/node.js';

const [flag, jobsFile, dpiArg] = process.argv.slice(2);
const dpi = Number(dpiArg);
if (flag !== '--batch' || !jobsFile || !Number.isInteger(dpi) || dpi < 1) {
  throw new Error('Usage: node kanera.mjs --batch <jobs.json> <dpi>');
}

const jobs = JSON.parse(await readFile(jobsFile, 'utf8'));
const times = [];
for (const [input, output] of jobs) {
  const start = performance.now();
  const result = await convertDocxToImages(await readFile(input), { format: 'png', dpi, wordPlatform: 'windows' });
  await mkdir(output, { recursive: true });
  for (const image of result.images) await writeFile(join(output, `page-${image.page}.png`), image.data);
  times.push(performance.now() - start);
}
console.log(JSON.stringify(times));
