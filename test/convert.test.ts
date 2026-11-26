/**
 * End-to-end conversions through the public entry points. Ghostscript, when installed, reads the
 * output back as an independent check that the PDF is valid and its text is extractable.
 */
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { convertDocx as convertInBrowser } from '../src/browser.js';
import { decodeJpegPixels } from '../src/image/jpeg-decoder.js';
import { encodeJpeg } from '../src/image/jpeg-encoder.js';
import { convertDocx as convertCore } from '../src/index.js';
import { convertDocx, docxToPdf } from '../src/node.js';
import { makeDocx, p } from './helpers/docx.js';
import { inlineImage, redPng, sampleDocx } from './helpers/sample.js';
import { ghostscript, ghostscriptRenders, pdfPageCount, pdfPagesText, pdfStreams, pdfText } from './helpers/pdf.js';

describe('convertDocx (node)', () => {
  it('uses automatic photo quality and accepts an explicit quality or no lossy re-encoding', async () => {
    const width = 192, height = 128;
    const data = new Uint8Array(width * height * 4);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        data.set([80 + x / 2, 100 + y / 2, 120 + 17 * Math.sin((x + y) / 9), 255], i);
      }
    }
    const jpeg = Buffer.from(encodeJpeg(width, height, data, 100, { channels: 4 }));
    const docx = makeDocx({ body: `<w:p>${inlineImage('photo', 96, 64)}</w:p>`, media: { photo: { name: 'photo.jpg', data: jpeg } } });
    const original = await docxToPdf(docx, { imageQuality: false });
    const automatic = await docxToPdf(docx);
    const smaller = await docxToPdf(docx, { imageQuality: 40 });
    expect(Buffer.from(original).includes(jpeg)).toBe(true);
    expect(automatic.length).toBeLessThan(original.length);
    expect(smaller.length).toBeLessThan(automatic.length);
    await expect(docxToPdf(docx, { imageQuality: 101 })).rejects.toThrow(/image quality/i);
  });

  it('preserves a coarsely quantized JPEG by default and honors explicit quality even when output grows', async () => {
    const width = 16, height = 16;
    const data = Uint8Array.from({ length: width * height * 4 }, (_, i) => i % 4 === 3 ? 255 : (i * 37) % 251);
    const jpeg = Buffer.from(encodeJpeg(width, height, data, 15, { channels: 4 }));
    expect(jpeg.length).toBeLessThan(1024);
    const pixels = decodeJpegPixels(jpeg);
    const doc = makeDocx({ body: `<w:p>${inlineImage('img', 16, 16)}</w:p>`, media: { img: { name: 'small.jpg', data: jpeg } } });
    expect(Buffer.from(await docxToPdf(doc)).includes(jpeg)).toBe(true);
    for (const quality of [80, 100]) {
      const expected = Buffer.from(encodeJpeg(pixels.width, pixels.height, pixels.data, quality, { channels: 4 }));
      expect(expected.length).toBeGreaterThan(jpeg.length);
      expect(Buffer.from(await docxToPdf(doc, { imageQuality: quality })).includes(expected)).toBe(true);
    }
  });

  it('writes supported Word SVGs as reusable PDF forms and uses the preview for unsupported features', async () => {
    const blip = '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="vector"/>';
    const drawing = inlineImage('preview', 96, 64).replace('<a:blip r:embed="preview"/>',
      `<a:blip r:embed="preview"><a:extLst><a:ext>${blip}</a:ext></a:extLst></a:blip>`);
    const convert = (shape: string) => convertDocx(makeDocx({
      body: `<w:p>${drawing}${drawing}</w:p>`,
      media: {
        vector: { name: 'vector.svg', data: new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 8"><metadata>editor metadata</metadata>${shape}</svg>`) },
        preview: { name: 'preview.png', data: redPng() },
      },
    }), { compress: false });
    const vector = await convert('<path d="M0 0h12v8h-12z" fill="#00f"/>');
    const source = pdfText(vector.pdf);
    expect(source.match(/\/Subtype \/Form\b/g)).toHaveLength(1);
    expect(source).toContain('0 0 1 rg');
    expect(source).not.toContain('editor metadata');
    expect(source).not.toContain('/Subtype /Image');
    const preview = await convert('<text>Use Word’s preview</text>');
    expect(preview.warnings).toEqual([]);
    expect(pdfText(preview.pdf).match(/\/Subtype \/Image\b/g)).toHaveLength(1);
  });

  it('washes out and recontrasts pictures as Word does, embedding each setting once', async () => {
    // Washout, in DrawingML (brightness 70%, contrast -70%) or VML (gain 0.3, blacklevel 0.35),
    // draws each colour component c as 0.3c + 0.805; VML gain 2 alone draws 2c - 0.5. Word's PDFs
    // agree. A cropped picture keeps its crop.
    const washed = inlineImage('red', 96, 64).replace('<a:blip r:embed="red"/>',
      '<a:blip r:embed="red"><a:lum bright="70000" contrast="-70000"/></a:blip><a:srcRect l="25000"/>');
    const vml = (attrs: string) => `<w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:72pt;height:48pt"><v:imagedata r:id="red" ${attrs}/></v:shape></w:pict></w:r>`;
    const { pdf } = await convertDocx(makeDocx({
      body: `<w:p>${inlineImage('red', 96, 64)}${washed}${vml('gain="19661f" blacklevel="22938f"')}${vml('gain="2"')}</w:p>`,
      media: { red: { name: 'red.png', data: redPng() } },
    }), { compress: false, imageLimits: { maxTotalPixels: 8 } });
    const images = pdfText(pdf).match(/\/Subtype \/Image\b[^]*?stream/g) ?? [];
    expect(images.map((d) => /\/Decode \[([^\]]*)\]/.exec(d)?.[1] ?? 'none').sort())
      .toEqual(['-0.5 1.5 -0.5 1.5 -0.5 1.5', '0.805 1.105 0.805 1.105 0.805 1.105', 'none']);
  });

  it('produces a readable PDF with embedded subset fonts, images and links', async () => {
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
    // Table shading and the red underline are drawn, the underline as a filled rectangle as in
    // Word's PDFs.
    const content = pdfStreams(result.pdf).join('\n');
    expect(content).toMatch(/0\.851 0\.886 0\.953 rg/);
    expect(content).toMatch(/0\.753 0 0 rg [\d. ]+ re f/);
    expect(content).not.toMatch(/ l S/);
    if (ghostscript) {
      ghostscriptRenders(result.pdf);
      const [one, two] = pdfPagesText(result.pdf).map((t) => t.replace(/\s+/g, ' '));
      for (const s of ['Quarterly report', 'Plain, bold, italic, underlined red', 'Grüße — “quoted” € 42 ½ Ωmega', '• First bullet', 'Region', '1,204', 'Visit example.com', 'Footnote text.', 'ACME Corp — Confidential', 'Page 1']) {
        expect(one).toContain(s);
      }
      expect(two).toContain('Second page');
      expect(two).toContain('Page 2');
    }
  });

  it('stops stroking text after a synthesised bold run', async () => {
    // Roboto Light has no bold face, so its bold is drawn with a stroke; the text render
    // mode outlives ET and must not carry over to the plain paragraph that follows.
    const body = p('Heading', '', '<w:rFonts w:ascii="Roboto Light" w:hAnsi="Roboto Light"/><w:b/>') + p('Plain');
    const { pdf } = await convertDocx(makeDocx({ body }), { compress: false });
    const content = pdfStreams(pdf).join('\n');
    const stroked = content.split('\n').filter((l) => l.includes(' 2 Tr '));
    expect(stroked).toHaveLength(1);
    const at = content.indexOf(stroked[0]);
    expect(content.slice(0, at).trimEnd().endsWith('q')).toBe(true);
    expect(content.slice(at + stroked[0].length).trimStart().startsWith('Q')).toBe(true);
  });

  it('snaps text to Word for Mac’s 0.24pt grid only when asked, moving a line as a whole', async () => {
    // 11pt text indented a quarter point: Word for Mac draws it at 11.04pt, starts the line at
    // 72.24 and puts the baseline on the grid; the bold run after it keeps its distance.
    const run = (text: string, rPr: string) => `<w:r><w:rPr>${rPr}<w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const body = `<w:p><w:pPr><w:ind w:left="5"/></w:pPr>${run('Plain ', '')}${run('bold', '<w:b/>')}</w:p>`;
    const runs = async (wordPlatform?: 'windows' | 'mac') => {
      const { pdf } = await convertDocx(makeDocx({ body }), { compress: false, wordPlatform });
      const tm = /\/F\d+ ([\d.]+) Tf .*? ([\d.]+) ([\d.]+) Tm/g;
      // The A4 page is 841.9pt tall; the grid runs down from its top.
      return [...pdfStreams(pdf).join('\n').matchAll(tm)].map((m) => ({ size: +m[1], x: +m[2], y: 841.9 - +m[3] }));
    };
    const win = await runs();
    const mac = await runs('mac');
    expect(win.map((r) => r.size)).toEqual([11, 11]);
    expect(win[0].x).toBeCloseTo(72.25, 3);
    expect(mac.map((r) => r.size)).toEqual([11.04, 11.04]);
    expect(mac[0].x).toBeCloseTo(72.24, 3);
    expect(mac[1].x - mac[0].x).toBeCloseTo(win[1].x - win[0].x, 3);
    for (const r of mac) expect(Math.abs(r.y / 0.24 - Math.round(r.y / 0.24))).toBeLessThan(0.01);
    expect(Math.abs(mac[0].y - win[0].y)).toBeLessThanOrEqual(0.12);
  });

  it('accepts ArrayBuffers and file paths', async () => {
    const bytes = sampleDocx();
    const fromBuffer = await docxToPdf(new Uint8Array(bytes).buffer);
    expect(pdfPageCount(fromBuffer)).toBe(2);
    const { writeFileSync, mkdtempSync, rmSync } = await import('node:fs');
    const { join } = await import('node:path');
    const { tmpdir } = await import('node:os');
    const dir = mkdtempSync(join(tmpdir(), 'docx2pdf-test-'));
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
    expect(requested.filter((u) => /Roboto|Gelasio|Caladea|Selawik|Morrow|Mono/.test(u))).toEqual([]);
  });

  it('embeds the default Aptos Narrow and Aptos Light substitutes in Node and the browser', async () => {
    const faces = [
      ['Aptos Narrow', '', 'MorrowSans-Narrow-Regular.ttf', 'MorrowSans-Narrow'],
      ['Aptos Narrow', '<w:b/>', 'MorrowSans-Narrow-Bold.ttf', 'MorrowSans-Narrow-Bold'],
      ['Aptos Narrow', '<w:i/>', 'MorrowSans-Narrow-Italic.ttf', 'MorrowSans-Narrow-Italic'],
      ['Aptos Narrow Bold Italic', '', 'MorrowSans-Narrow-BoldItalic.ttf', 'MorrowSans-Narrow-Bold-Italic'],
      ['Aptos Light', '', 'MorrowSans-Light.ttf', 'MorrowSans-Light'],
      ['Aptos Light Italic', '', 'MorrowSans-LightItalic.ttf', 'MorrowSans-LightItalic'],
    ];
    const body = faces.map(([family, style]) => p('AV café Ω Ж', '', `<w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>${style}`)).join('');
    // An empty styles part keeps Word's built-in Aptos defaults (and their Morrow Sans) out of it.
    const docx = makeDocx({ styles: '', body });
    const fetched: string[] = [];
    const browser = await convertInBrowser(docx, {
      fontBaseUrl: 'https://cdn.test/aptos/',
      fetch: async (url) => {
        const file = new URL(String(url)).pathname.split('/').at(-1)!;
        fetched.push(file);
        return new Response(readFileSync(new URL(`../fonts/${file}`, import.meta.url)));
      },
    });
    expect(fetched.filter((file) => file.startsWith('Morrow')).sort()).toEqual(faces.map(([, , file]) => file).sort());
    for (const result of [await convertDocx(docx), browser]) {
      expect(result.warnings).toEqual([]);
      const pdf = pdfText(result.pdf);
      for (const [, , , name] of faces) expect(pdf).toMatch(new RegExp(`/BaseFont /[^\\s]+\\+${name}(?=\\s)`));
    }
  });
});
