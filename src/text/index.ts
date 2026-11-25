/**
 * Text extraction without the renderer: only the parts that hold text are parsed, no fonts are
 * loaded, pictures are not decoded and nothing is laid out.
 */
import { parseDocx } from '../docx/parser.js';
import type { InflateBytes } from '../util/compression.js';
import type { ArchiveLimits } from '../util/zip.js';
import { documentToText, type TextOptions } from './extract.js';

export type { TextFormat, TextOptions } from './extract.js';

export interface ExtractTextOptions extends TextOptions {
  /** Maximum uncompressed bytes per ZIP part and across all parts read. */
  archiveLimits?: ArchiveLimits;
  /** Stops extraction, rejecting with the signal's reason, if aborted before it starts. */
  signal?: AbortSignal;
}

export interface ExtractTextResult {
  /** The document's text, or Markdown when `format` is 'markdown'. Ends with a newline unless empty. */
  text: string;
  /** Content that was skipped (unsupported imports and the like). */
  warnings: string[];
}

/** Extract a .docx's text as plain text or Markdown. */
export function extractTextWithCompression(input: Uint8Array | ArrayBuffer, options: ExtractTextOptions, compression: { inflateRaw: InflateBytes }): ExtractTextResult {
  options.signal?.throwIfAborted();
  const data = input instanceof Uint8Array ? input : new Uint8Array(input);
  const doc = parseDocx(data, compression.inflateRaw, options.archiveLimits, false, true);
  return { text: documentToText(doc, options), warnings: [...new Set(doc.warnings)] };
}
