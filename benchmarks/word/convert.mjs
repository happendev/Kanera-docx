import { readFile, writeFile } from 'node:fs/promises';

const [engine, inputPath, outputPath] = process.argv.slice(2);
if (!['ours', 'eigenpal'].includes(engine) || !inputPath || !outputPath) {
  throw new Error('Usage: node convert.mjs <ours|eigenpal> <input.docx> <output.pdf> or --batch <jobs.json>');
}

let render;
if (engine === 'ours') {
  const { convertDocx } = await import('../../dist/node.js');
  render = convertDocx;
} else {
  const { exportPdf } = await import('@docx-editor.dev/docx-to-pdf');
  render = (input) => exportPdf(input, {
    displayMode: 'proposed',
    comments: false,
    fidelityPolicy: 'best-effort',
  });
}

async function convert(input, output) {
  const result = await render(await readFile(input));
  await writeFile(output, result.pdf ?? result.bytes);
  return { warnings: result.warnings, diagnostics: result.diagnostics };
}

if (inputPath === '--batch') {
  const jobs = JSON.parse(await readFile(outputPath, 'utf8'));
  const times = [];
  for (const [input, output] of jobs) {
    const start = performance.now();
    await convert(input, output);
    times.push(performance.now() - start);
  }
  console.log(JSON.stringify(times));
} else {
  console.log(JSON.stringify(await convert(inputPath, outputPath)));
}
