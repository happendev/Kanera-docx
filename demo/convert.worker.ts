/// <reference lib="webworker" />
// Runs the converter off the main thread. Font faces are fetched on demand and stay cached in
// this worker for the page's lifetime, so later conversions only download faces they haven't seen.
import { convertDocx, type BrowserConvertOptions } from '../src/browser.js';

export interface ConvertRequest {
  id: number;
  docx: ArrayBuffer;
  options: Pick<BrowserConvertOptions, 'fontPack' | 'wordPlatform' | 'imageQuality' | 'embeddedFonts'>;
}

export type ConvertResponse =
  | {
      id: number;
      ok: true;
      pdf: ArrayBuffer;
      pageCount: number;
      warnings: string[];
      renderedBreaks: number;
      matchedBreaks: number;
      ms: number;
      fonts: { file: string; bytes: number }[];
    }
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

self.onmessage = async (e: MessageEvent<ConvertRequest>) => {
  const { id, docx, options } = e.data;
  downloaded = [];
  const start = performance.now();
  try {
    const result = await convertDocx(new Uint8Array(docx), { ...options, fontBaseUrl, fetch: countingFetch });
    const ms = performance.now() - start;
    const bytes = result.pdf;
    const pdf = bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : bytes.slice().buffer;
    const msg: ConvertResponse = {
      id,
      ok: true,
      pdf,
      pageCount: result.pageCount,
      warnings: result.warnings,
      renderedBreaks: result.stats.renderedBreaks,
      matchedBreaks: result.stats.matchedBreaks,
      ms,
      fonts: downloaded,
    };
    (self as unknown as DedicatedWorkerGlobalScope).postMessage(msg, [pdf]);
  } catch (err) {
    const msg: ConvertResponse = { id, ok: false, error: err instanceof Error ? err.message : String(err) };
    (self as unknown as DedicatedWorkerGlobalScope).postMessage(msg);
  }
};
