// Converts DOCX to Markdown with a JavaScript engine.
// node worker.mjs <kanera|mammoth> <input.docx> <output.md>
// node worker.mjs <kanera|mammoth> --batch <jobs.json>   (prints one time in ms per job, or null)
import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';

const [engine, inputPath, outputPath] = process.argv.slice(2);
const require = createRequire(import.meta.url);
let render;
if (engine === 'kanera') {
  const { docxToMarkdown } = await import('../../dist/node.js');
  render = docxToMarkdown;
} else if (engine === 'mammoth') {
  // The usual JavaScript route: mammoth's HTML, then turndown with its GFM plugin.
  const mammoth = require('mammoth');
  const TurndownService = require('turndown');
  const { gfm } = require('turndown-plugin-gfm');
  const turndown = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-' });
  turndown.use(gfm);
  render = async (input) => turndown.turndown((await mammoth.convertToHtml({ buffer: input })).value);
} else {
  throw new Error('Usage: node worker.mjs <kanera|mammoth> <input.docx> <output.md> or --batch <jobs.json>');
}

async function convert(input, output) {
  await writeFile(output, await render(await readFile(input)));
}

if (inputPath === '--batch') {
  const times = [];
  for (const [input, output] of JSON.parse(await readFile(outputPath, 'utf8'))) {
    const start = performance.now();
    try {
      await convert(input, output);
      times.push(performance.now() - start);
    } catch {
      times.push(null);
    }
  }
  console.log(JSON.stringify(times));
} else {
  await convert(inputPath, outputPath);
}
