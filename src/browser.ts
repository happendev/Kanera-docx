import { readInputBlob } from './util/input.js';
import { convertDocx as convertCore, convertDocxToImages as imagesCore, extractText as extractCore } from './index.js';
import type { ConvertOptions, ConvertResult, ExtractTextOptions, ExtractTextResult, ImageOptions, ImagesResult, PageImage } from './core.js';
import { COMMON_FONTS } from './fonts/bundled.js';
import { fontProviders } from './browser-fonts.js';

export type { ConvertOptions, ConvertResult, ExtractTextOptions, ExtractTextResult, FontSource, ImageFormat, ImageOptions, ImagesResult, LayoutStats, PageImage } from './core.js';
export type { TextFormat } from './text/extract.js';

/** Files to copy from the package's fonts/ directory for the common browser pack. */
export const commonFontFiles: readonly string[] = COMMON_FONTS.map(({ file }) => file);

export interface BrowserFontOptions {
  /**
   * URL of a folder serving the package's `fonts/` files. Defaults to the folder next to this
   * module, which works when the package is served as-is (not bundled). With a bundler, copy the
   * selected pack's files from `node_modules/@kanera/docx2pdf/fonts` to your static assets.
   * Set to `false` to disable loading fonts from URLs automatically.
   */
  fontBaseUrl?: string | URL | false;
  /** Defaults to 'full'. Set to 'common' to limit bundled faces to commonFontFiles. */
  fontPack?: 'common' | 'full';
  /** Custom fetch (for caching or auth). */
  fetch?: typeof fetch;
}

export interface BrowserConvertOptions extends ConvertOptions, BrowserFontOptions {}

export interface BrowserImageOptions extends ImageOptions, BrowserFontOptions {}

/** Convert a .docx to PDF in the browser. Bundled fonts are fetched on demand. */
export async function convertDocx(input: Uint8Array | ArrayBuffer | Blob, options: BrowserConvertOptions = {}): Promise<ConvertResult> {
  const data = input instanceof Blob ? await readInputBlob(input, options) : input;
  return convertCore(data, options, fontProviders(options));
}

/** Render a .docx's pages as PNG (default) or JPEG images in the browser. Bundled fonts are fetched on demand. */
export async function convertDocxToImages(input: Uint8Array | ArrayBuffer | Blob, options: BrowserImageOptions = {}): Promise<ImagesResult> {
  const data = input instanceof Blob ? await readInputBlob(input, options) : input;
  return imagesCore(data, options, fontProviders(options));
}

/** Render a .docx's pages as PNG (default) or JPEG images in the browser. */
export async function docxToImages(input: Uint8Array | ArrayBuffer | Blob, options: BrowserImageOptions = {}): Promise<PageImage[]> {
  return (await convertDocxToImages(input, options)).images;
}

/** Convert a .docx to PDF bytes in the browser. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer | Blob, options: BrowserConvertOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  return (await convertDocx(input, options)).pdf;
}

/** Extract a .docx's text as plain text or Markdown (`format: 'markdown'`), with warnings. */
export async function extractText(input: Uint8Array | ArrayBuffer | Blob, options: ExtractTextOptions = {}): Promise<ExtractTextResult> {
  return extractCore(input instanceof Blob ? await readInputBlob(input, options) : input, options);
}

/** A .docx's text as plain text. */
export async function docxToText(input: Uint8Array | ArrayBuffer | Blob, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'text' })).text;
}

/** A .docx's text as GitHub-flavoured Markdown. */
export async function docxToMarkdown(input: Uint8Array | ArrayBuffer | Blob, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'markdown' })).text;
}
