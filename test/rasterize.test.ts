import { spawnSync } from 'node:child_process';
import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, it } from 'vitest';
import { convertDocx } from '../src/node.js';
import { rasterizePdf } from '../src/rasterize.js';
import { makeDocx, p } from './helpers/docx.js';

const installed = ['pdftoppm', 'mutool', process.platform === 'win32' ? 'gswin64c' : 'gs']
  .some((tool) => !spawnSync(tool, ['-v']).error);

it.skipIf(!installed)('writes one zero-padded PNG per page in page order', async () => {
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const body = Array.from({ length: 10 }, (_, i) => p(`Page ${i + 1}`)).join(pageBreak);
  const result = await convertDocx(makeDocx({ body }));
  expect(result.pageCount).toBe(10);
  const dir = mkdtempSync(join(tmpdir(), 'docx2pdf-png-'));
  try {
    const out = join(dir, 'nested', 'pages');
    const written = await rasterizePdf(result.pdf, { dir: out, prefix: 'doc', dpi: 36, pageCount: result.pageCount });
    const names = Array.from({ length: 10 }, (_, i) => `doc-${String(i + 1).padStart(2, '0')}.png`);
    expect(written).toEqual(names.map((name) => join(out, name)));
    expect(readdirSync(out).sort()).toEqual(names);
    const png = readFileSync(written[0]);
    expect(png.subarray(1, 4).toString('latin1')).toBe('PNG');
    // A4 at 36 dpi is 298 pixels wide (rasterisers round 297.6 differently, so allow either side).
    expect(Math.abs(png.readUInt32BE(16) - 298)).toBeLessThanOrEqual(1);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
