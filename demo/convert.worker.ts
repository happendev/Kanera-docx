/// <reference lib="webworker" />
// Runs the converter off the main thread. Font faces are fetched on demand and stay cached in
// this worker for the page's lifetime, so later conversions only download faces they haven't seen.
import { convertDocx, extractText, type BrowserConvertOptions, type TextFormat } from '../src/browser.js';

declare const self: DedicatedWorkerGlobalScope;

export type OutputFormat = 'pdf' | TextFormat;

export type ConvertRequest = {
  id: number;
  docx: ArrayBuffer;
} & (
  | { format: 'pdf'; options: Pick<BrowserConvertOptions, 'fontPack' | 'wordPlatform' | 'imageQuality' | 'embeddedFonts'> }
  | { format: TextFormat }
);

export type ConvertResponse =
  | {
      id: number;
      ok: true;
      format: 'pdf';
      pdf: ArrayBuffer;
      pageCount: number;
      warnings: string[];
      renderedBreaks: number;
      matchedBreaks: number;
      ms: number;
      fonts: { file: string; bytes: number }[];
    }
  | { id: number; ok: true; format: TextFormat; text: string; warnings: string[]; ms: number }
  | { id: number; ok: false; error: string };

const fontBaseUrl = new URL('../fonts/', self.location.href);
let downloaded: { file: string; bytes: number }[] = [];

const countingFetch: typeof fetch = async (input, init) => {
  const res = await fetch(input, init);
  if (!res.ok) return res;
  const body = await res.arrayBuffer();
  const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
  downloaded.push({ file: decodeURIComponent(url.slice(url.lastIndexOf('/') + 1)), bytes: body.byteLength });
  return new Response(body, { status: res.status, headers: res.headers });
};

// PDF requests share the font cache and download counter. Text extraction can run independently.
let pdfQueue = Promise.resolve();
self.onmessage = (e: MessageEvent<ConvertRequest>) => {
  const request = e.data;
  if (request.format === 'pdf') pdfQueue = pdfQueue.then(() => convert(request));
  else void convert(request);
};

async function convert(request: ConvertRequest) {
  const { id, docx, format } = request;
  const start = performance.now();
  try {
    if (format !== 'pdf') {
      const result = await extractText(docx, { format });
      const msg: ConvertResponse = { id, ok: true, format, ...result, ms: performance.now() - start };
      self.postMessage(msg);
      return;
    }

    downloaded = [];
    const result = await convertDocx(docx, { ...request.options, fontBaseUrl, fetch: countingFetch });
    const ms = performance.now() - start;
    const bytes = result.pdf;
    const pdf = bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : bytes.slice().buffer;
    const msg: ConvertResponse = {
      id,
      ok: true,
      format,
      pdf,
      pageCount: result.pageCount,
      warnings: result.warnings,
      renderedBreaks: result.stats.renderedBreaks,
      matchedBreaks: result.stats.matchedBreaks,
      ms,
      fonts: downloaded,
    };
    self.postMessage(msg, [pdf]);
  } catch (err) {
    const msg: ConvertResponse = { id, ok: false, error: err instanceof Error ? err.message : String(err) };
    self.postMessage(msg);
  }
}
