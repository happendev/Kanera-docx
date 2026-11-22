/**
 * Runtime-neutral entry point. It ships no fonts: pass them in `options.fonts`. The `node` and
 * `browser` entry points (picked automatically through package exports) add the bundled fonts.
 */
import { pakoCompression } from './util/pako.js';
import { convertWithCompression, type ConvertOptions, type ConvertResult, type FontProvider } from './core.js';
import { extractTextWithCompression, type ExtractTextOptions, type ExtractTextResult } from './text/index.js';

export function convertDocx(input: Uint8Array | ArrayBuffer, options: ConvertOptions = {}, providers: FontProvider[] = []): Promise<ConvertResult> {
  return convertWithCompression(input, options, providers, pakoCompression);
}
export type { ConvertOptions, ConvertResult, ExtractTextOptions, ExtractTextResult, FontSource, LayoutStats } from './core.js';
export type { TextFormat } from './text/extract.js';

/** Convert a .docx file to PDF bytes. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer, options: ConvertOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  return (await convertDocx(input, options)).pdf;
}

/** Extract a .docx's text as plain text or Markdown (`format: 'markdown'`), with warnings. */
export async function extractText(input: Uint8Array | ArrayBuffer, options: ExtractTextOptions = {}): Promise<ExtractTextResult> {
  return extractTextWithCompression(input, options, pakoCompression);
}

/** A .docx's text as plain text. */
export async function docxToText(input: Uint8Array | ArrayBuffer, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'text' })).text;
}

/** A .docx's text as GitHub-flavoured Markdown. */
export async function docxToMarkdown(input: Uint8Array | ArrayBuffer, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'markdown' })).text;
}
