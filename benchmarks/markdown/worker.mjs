// Converts DOCX to Markdown, or to plain text for a .txt output, with a JavaScript engine.
// node worker.mjs <kanera|mammoth> <input.docx> <output.md|output.txt>
// node worker.mjs <kanera|mammoth> --batch <jobs.json>   (prints one time in ms per job, or null)
import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';

const [engine, inputPath, outputPath] = process.argv.slice(2);
const require = createRequire(import.meta.url);
let render;
if (engine === 'kanera') {
  const { docxToMarkdown, docxToText } = await import('../../dist/node.js');
  render = (input, text) => (text ? docxToText(input) : docxToMarkdown(input));
} else if (engine === 'mammoth') {
  // The usual JavaScript route to Markdown: mammoth's HTML, then turndown with its GFM plugin.
  // Plain text is mammoth's own raw text.
  const mammoth = require('mammoth');
  const TurndownService = require('turndown');
  const { gfm } = require('turndown-plugin-gfm');
  const turndown = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-' });
  turndown.use(gfm);
  render = async (input, text) => text
    ? (await mammoth.extractRawText({ buffer: input })).value
    : turndown.turndown((await mammoth.convertToHtml({ buffer: input })).value);
} else {
  throw new Error('Usage: node worker.mjs <kanera|mammoth> <input.docx> <output.md|output.txt> or --batch <jobs.json>');
}

async function convert(input, output) {
  await writeFile(output, await render(await readFile(input), output.endsWith('.txt')));
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
