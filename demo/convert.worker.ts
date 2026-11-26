/// <reference lib="webworker" />
// Runs the converter off the main thread. Font faces are fetched on demand and stay cached in
// this worker for the page's lifetime, so later conversions only download faces they haven't seen.
import { extractText, type BrowserConvertOptions, type TextFormat } from '../src/browser.js';
import { fontProviders } from '../src/browser-fonts.js';
import { convertWithCompression, prepareDocxWithCompression } from '../src/core.js';
import { pakoCompression } from '../src/util/pako.js';

declare const self: DedicatedWorkerGlobalScope;

export type OutputFormat = 'pdf' | TextFormat;

export type ConvertRequest = {
  id: number;
  docx: ArrayBuffer;
} & (
  | { format: 'pdf'; options: Pick<BrowserConvertOptions, 'fontPack' | 'wordPlatform' | 'imageQuality' | 'useEmbeddedFonts'> }
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
const fontBytes = new Map<string, Uint8Array>();
let downloads = new Map<string, Promise<Uint8Array>>();
let downloaded: { file: string; bytes: number }[] = [];

async function readFontFromMemory(url: string): Promise<Uint8Array> {
  const bytes = fontBytes.get(url);
  if (bytes) return bytes;
  const attempt = downloads.get(url);
  if (attempt) return attempt;
  throw new Error(`Font was not loaded before conversion: ${url}`);
}

async function downloadFont(url: string): Promise<Uint8Array> {
  const bytes = fontBytes.get(url);
  if (bytes) return bytes;
  let attempt = downloads.get(url);
  if (!attempt) {
    const fonts = downloaded;
    attempt = (async () => {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status} for ${url}`);
      const data = new Uint8Array(await response.arrayBuffer());
      fontBytes.set(url, data);
      fonts.push({ file: decodeURIComponent(url.slice(url.lastIndexOf('/') + 1)), bytes: data.byteLength });
      return data;
    })();
    downloads.set(url, attempt);
  }
  return attempt;
}

// PDF requests share the font cache and download counter. Text extraction can run independently.
let pdfQueue = Promise.resolve();
self.onmessage = (e: MessageEvent<ConvertRequest>) => {
  const request = e.data;
  if (request.format === 'pdf') pdfQueue = pdfQueue.then(() => convert(request)).catch(reportResponseError);
  else void convert(request).catch(reportResponseError);
};

function reportResponseError(err: unknown) {
  console.error('Could not send the conversion response:', err);
}

async function convert(request: ConvertRequest) {
  const { id, docx, format } = request;
  try {
    if (format !== 'pdf') {
      const start = performance.now();
      const result = await extractText(docx, { format });
      const msg: ConvertResponse = { id, ok: true, format, ...result, ms: performance.now() - start };
      self.postMessage(msg);
      return;
    }

    downloaded = [];
    downloads = new Map();
    const options = { ...request.options, fontBaseUrl };
    // Resolve the font requests first, including any speculative reads, before timing a
    // conversion that can only read font bytes already in memory.
    try {
      await prepareDocxWithCompression(docx, options, fontProviders(options, downloadFont), pakoCompression);
    } finally {
      await Promise.allSettled(downloads.values());
    }

    const start = performance.now();
    const result = await convertWithCompression(docx, options, fontProviders(options, readFontFromMemory), pakoCompression);
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
