/**
 * End-to-end conversions through the public entry points. Ghostscript, when installed, reads the
 * output back as an independent check that the PDF is valid and its text is extractable.
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { convertDocx as convertInBrowser } from '../src/browser.js';
import { convertDocx as convertCore } from '../src/index.js';
import { convertDocx, docxToPdf } from '../src/node.js';
import { makeDocx, p } from './helpers/docx.js';
import { inlineImage, sampleDocx } from './helpers/sample.js';
import { ghostscript, ghostscriptRenders, pdfPageCount, pdfPagesText, pdfStreams, pdfText } from './helpers/pdf.js';

describe('convertDocx (node)', () => {
  it('produces a well-formed PDF with embedded subset fonts, images and links', async () => {
    const result = await convertDocx(sampleDocx());
    expect(result.pageCount).toBe(2);
    // Calibri Light has no metric-compatible open substitute, and says so.
    expect(result.warnings).toEqual(['Font "Calibri Light" not available; using Carlito']);
    const src = pdfText(result.pdf);
    expect(src.startsWith('%PDF-1.7')).toBe(true);
    expect(src.trimEnd().endsWith('%%EOF')).toBe(true);
    expect(pdfPageCount(result.pdf)).toBe(2);
    expect(src).toContain('/FontFile2');
    expect(src).toContain('/ToUnicode');
    expect(src).toContain('/Subtype /Image');
    expect(src).toContain('/URI (https://example.com/)');
    // Fonts are subset: the whole file stays small.
    expect(result.pdf.length).toBeLessThan(200_000);
    // Table shading and the red underline are drawn.
    const content = pdfStreams(result.pdf).join('\n');
    expect(content).toMatch(/0\.851 0\.886 0\.953 rg/);
    expect(content).toMatch(/0\.753 0 0 RG/);
  });

  it.skipIf(!ghostscript)('is readable by Ghostscript with the document text intact', async () => {
    const { pdf } = await convertDocx(sampleDocx());
    ghostscriptRenders(pdf);
    const [one, two] = pdfPagesText(pdf).map((t) => t.replace(/\s+/g, ' '));
    for (const s of ['Quarterly report', 'Plain, bold, italic, underlined red', 'Grüße — “quoted” € 42 ½ Ωmega', '• First bullet', 'Region', '1,204', 'Visit example.com', 'Footnote text.', 'ACME Corp — Confidential', 'Page 1']) {
      expect(one).toContain(s);
    }
    expect(two).toContain('Second page');
    expect(two).toContain('Page 2');
  });

  it('accepts ArrayBuffers and file paths', async () => {
    const bytes = sampleDocx();
    const fromBuffer = await docxToPdf(new Uint8Array(bytes).buffer);
    expect(pdfPageCount(fromBuffer)).toBe(2);
    const { writeFileSync, mkdtempSync, rmSync } = await import('node:fs');
    const { join } = await import('node:path');
    const { tmpdir } = await import('node:os');
    const dir = mkdtempSync(join(tmpdir(), 'folio-test-'));
    try {
      writeFileSync(join(dir, 'in.docx'), bytes);
      expect((await convertDocx(join(dir, 'in.docx'))).pageCount).toBe(2);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it('warns about content it cannot reproduce instead of failing', async () => {
    const emf = new Uint8Array(100);
    emf.set([1, 0, 0, 0], 0);
    emf.set([0x20, 0x45, 0x4d, 0x46], 40);
    const docx = makeDocx({ body: `<w:p>${inlineImage('img1', 50, 50)}</w:p>`, media: { img1: { name: 'chart.emf', data: emf } } });
    const result = await convertDocx(docx);
    expect(result.pageCount).toBe(1);
    expect(result.warnings.join('\n')).toMatch(/EMF/i);
  });

  it('rejects input that is not a .docx', async () => {
    await expect(convertDocx(new Uint8Array([1, 2, 3]))).rejects.toThrow();
  });
});

describe('other entry points', () => {
  it('the core entry converts with caller-supplied fonts only', async () => {
    const font = readFileSync(new URL('../fonts/LiberationSans-Regular.ttf', import.meta.url));
    const result = await convertCore(makeDocx({ body: p('Hello') }), { fonts: [{ data: font }] });
    expect(result.pageCount).toBe(1);
    expect(pdfText(result.pdf)).toContain('LiberationSans');
  });

  it('the browser entry fetches bundled fonts from fontBaseUrl', async () => {
    const requested: string[] = [];
    const fakeFetch = (async (url: string | URL) => {
      requested.push(String(url));
      const file = readFileSync(new URL(String(url).replace('https://cdn.test/fonts/', '../fonts/'), import.meta.url));
      return new Response(file);
    }) as typeof fetch;
    const result = await convertInBrowser(new Blob([new Uint8Array(sampleDocx())]), { fontBaseUrl: 'https://cdn.test/fonts/', fetch: fakeFetch });
    expect(result.pageCount).toBe(2);
    expect(requested.every((u) => u.startsWith('https://cdn.test/fonts/'))).toBe(true);
    // Families the document does not use are never downloaded.
    expect(requested.some((u) => u.includes('Carlito'))).toBe(true);
    expect(requested.filter((u) => /Roboto|Gelasio|Caladea|Selawik|Mono/.test(u))).toEqual([]);
  });
});
