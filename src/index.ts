/**
 * Runtime-neutral entry point. It ships no fonts: pass them in `options.fonts`. The `node` and
 * `browser` entry points (picked automatically through package exports) add the bundled fonts.
 */
import { convertDocx, type ConvertOptions, type ConvertResult } from './core.js';

export { convertDocx };
export type { ConvertOptions, ConvertResult, FontSource, LayoutStats } from './core.js';

/** Convert a .docx file to PDF bytes. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer, options: ConvertOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  return (await convertDocx(input, options)).pdf;
}
