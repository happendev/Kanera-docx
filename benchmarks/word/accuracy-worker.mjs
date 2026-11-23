import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createInterface } from 'node:readline';
import { convertDocx, convertDocxToImages } from './dist/node.js';

const requests = createInterface({ input: process.stdin, crlfDelay: Infinity });

for await (const line of requests) {
  try {
    const [input, output, dpi] = JSON.parse(line);
    const options = { fonts: [], systemFonts: false, wordPlatform: 'windows' };
    if (dpi) {
      // Page images straight from the layout, written as <output>/page-<n>.png.
      const result = await convertDocxToImages(await readFile(input), { ...options, dpi });
      await mkdir(output, { recursive: true });
      for (const image of result.images) await writeFile(join(output, `page-${image.page}.png`), image.data);
      process.stdout.write(JSON.stringify({ pages: result.images.length }) + '\n');
    } else {
      const result = await convertDocx(await readFile(input), options);
      await writeFile(output, result.pdf);
      process.stdout.write('{}\n');
    }
  } catch (error) {
    process.stdout.write(JSON.stringify({ error: error instanceof Error ? error.message : String(error) }) + '\n');
  }
}
