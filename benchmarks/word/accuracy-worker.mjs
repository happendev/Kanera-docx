import { readFile, writeFile } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import { convertDocx } from './dist/node.js';

const requests = createInterface({ input: process.stdin, crlfDelay: Infinity });

for await (const line of requests) {
  try {
    const [input, output] = JSON.parse(line);
    const result = await convertDocx(await readFile(input), {
      fonts: [],
      systemFonts: false,
      wordPlatform: 'windows',
    });
    await writeFile(output, result.pdf);
    process.stdout.write('{}\n');
  } catch (error) {
    process.stdout.write(JSON.stringify({ error: error instanceof Error ? error.message : String(error) }) + '\n');
  }
}
