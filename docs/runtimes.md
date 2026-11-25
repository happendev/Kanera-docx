# Node, browser and other runtimes

The TypeScript renderer uses Node’s built-in zlib for compression. The browser and runtime-neutral
core use pako, bundled with the package. Conversion runs without LibreOffice, native add-ons or
WebAssembly.

## Node

Use Node.js 22 or later. `@kanera/docx2pdf` resolves to the Node build there; `@kanera/docx2pdf/node`
names it explicitly.

The package is ES modules only; CommonJS can `require()` it on Node 22.12+, or use
`await import()`. If you bundle your server, copy the package's `fonts/` folder into the build
output and pass its path as `fontDir`.

### In a web server

```ts
import { createServer } from 'node:http';
import { convertDocx } from '@kanera/docx2pdf';

const MAX_UPLOAD = 20 * 1024 * 1024;

createServer(async (req, res) => {
  if (req.method !== 'POST') return void res.writeHead(405).end();
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_UPLOAD) return void res.writeHead(413).end();
    chunks.push(chunk);
  }
  try {
    const { pdf } = await convertDocx(Buffer.concat(chunks));
    res.writeHead(200, { 'content-type': 'application/pdf' }).end(pdf);
  } catch (err) {
    res.writeHead(422, { 'content-type': 'text/plain' }).end(`Could not convert: ${(err as Error).message}`);
  }
}).listen(3000);
```

Conversion is synchronous, so a large document blocks other requests. For documents from users,
see [Converting untrusted documents](options.md#converting-untrusted-documents).

## Browser

Conversion runs in the browser. Browser conversion requires ES2022, `fetch` and `TextDecoder`.
To try the demo locally, run `pnpm demo:serve` from a checkout
with dependencies installed.
The demo previews and downloads PDFs, plain text (`.txt`), and Markdown (`.md`). In PDF mode, you can compare the included benchmark documents with Word's output. [Demo source](../demo/)

For a bundled app that serves `public/` at the site root, copy the bundled fonts into that folder
after installing the package.

```bash
node -e "require('node:fs').cpSync('node_modules/@kanera/docx2pdf/fonts', 'public/docx2pdf-fonts', { recursive: true })"
```

Pass that directory's public URL as `fontBaseUrl` and supply a `File` from a file input or drop
handler. The browser fetches individual font faces on demand.

```ts
import { convertDocx } from '@kanera/docx2pdf'; // resolves to the browser build

const { pdf } = await convertDocx(file, { fontBaseUrl: '/docx2pdf-fonts/' });
const pdfBlob = new Blob([pdf], { type: 'application/pdf' });
```

If you serve the package as-is, the default `fontBaseUrl` resolves to its bundled `fonts/` directory.

A plain Arial document downloads one face (about 400 KB); CJK text can fetch a Noto CJK face of
over 20 MB. Pass `fetch` to add caching or auth, or `fontBaseUrl: false` to use only embedded and
supplied fonts. For a smaller asset set, serve only the faces in `commonFontFiles` (exported from
`@kanera/docx2pdf/browser`) plus `fonts/licenses/`, and pass `fontPack: 'common'`.

Conversion runs on the calling thread. Use a Web Worker to keep the page responsive, and reuse the
worker across conversions so downloaded font faces stay cached.

```ts
// convert.worker.ts
import { convertDocx } from '@kanera/docx2pdf';

self.onmessage = async (e: MessageEvent<File>) => {
  try {
    self.postMessage(await convertDocx(e.data, { fontBaseUrl: '/docx2pdf-fonts/' }));
  } catch (err) {
    self.postMessage({ error: String(err) });
  }
};
```

```ts
// main.ts
const worker = new Worker(new URL('./convert.worker.ts', import.meta.url), { type: 'module' });
worker.onmessage = (e) => {
  if (e.data.error) return console.error(e.data.error);
  window.open(URL.createObjectURL(new Blob([e.data.pdf], { type: 'application/pdf' })));
};
worker.postMessage(file);
```

## Core (any runtime)

`@kanera/docx2pdf/core` loads no fonts itself; pass every font through `fonts`.

```ts
import { convertDocx } from '@kanera/docx2pdf/core';

const { pdf } = await convertDocx(bytes, { fonts: [{ data: regular }, { data: bold }] });
```
