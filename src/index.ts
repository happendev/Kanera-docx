/**
 * Runtime-neutral entry point. It ships no fonts: pass them in `options.fonts`. The `node` and
 * `browser` entry points (picked automatically through package exports) add the bundled fonts.
 */
import { pakoCompression } from './util/pako.js';
import { convertWithCompression, type ConvertOptions, type ConvertResult, type FontProvider } from './core.js';

export function convertDocx(input: Uint8Array | ArrayBuffer, options: ConvertOptions = {}, providers: FontProvider[] = []): Promise<ConvertResult> {
  return convertWithCompression(input, options, providers, pakoCompression);
}
export type { ConvertOptions, ConvertResult, FontSource, LayoutStats } from './core.js';

/** Convert a .docx file to PDF bytes. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer, options: ConvertOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  return (await convertDocx(input, options)).pdf;
}
