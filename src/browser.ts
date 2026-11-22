import { convertDocx as convertCore, extractText as extractCore } from './index.js';
import type { ConvertOptions, ConvertResult, ExtractTextOptions, ExtractTextResult, FontProvider } from './core.js';
import { BundledFontData } from './fonts/registry.js';
import { isFontCollection } from './fonts/sfnt.js';
import { COMMON_FONTS, FULL_FONTS } from './fonts/bundled.js';

export type { ConvertOptions, ConvertResult, ExtractTextOptions, ExtractTextResult, FontSource, LayoutStats } from './core.js';
export type { TextFormat } from './text/extract.js';

/** Files to copy from the package's fonts/ directory for the common browser pack. */
export const commonFontFiles: readonly string[] = COMMON_FONTS.map(({ file }) => file);

export interface BrowserConvertOptions extends ConvertOptions {
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

const fontCaches = new WeakMap<typeof fetch, Map<string, Promise<BundledFontData | Uint8Array>>>();

function defaultBase(): URL {
  // Kept out of a literal `new URL('…', import.meta.url)` so bundlers don't try to inline a folder.
  const rel = ['..', 'fonts', ''].join('/');
  return new URL(rel, import.meta.url);
}

/** Convert a .docx to PDF in the browser. Bundled fonts are fetched on demand. */
export async function convertDocx(input: Uint8Array | ArrayBuffer | Blob, options: BrowserConvertOptions = {}): Promise<ConvertResult> {
  const data = input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : input;
  const providers: FontProvider[] = [];
  if (options.fontBaseUrl !== false) {
    const base = options.fontBaseUrl ? new URL(String(options.fontBaseUrl), globalThis.location?.href ?? import.meta.url) : defaultBase();
    if (!base.pathname.endsWith('/')) base.pathname += '/';
    const doFetch = options.fetch ?? fetch;
    let fontCache = fontCaches.get(doFetch);
    if (!fontCache) {
      fontCache = new Map();
      fontCaches.set(doFetch, fontCache);
    }
    const faces = options.fontPack === 'common' ? COMMON_FONTS : FULL_FONTS;
    providers.push({
      register(registry) {
        for (const b of faces) {
          const url = new URL(b.file, base).href;
          registry.add(
            {
              family: b.family,
              bold: b.bold,
              italic: b.italic,
              weight: b.weight, covers: b.covers,
              load: () => {
                let p = fontCache.get(url);
                if (!p) {
                  p = doFetch(url).then(async (r) => {
                    if (!r.ok) throw new Error(`HTTP ${r.status} for ${url}`);
                    const bytes = new Uint8Array(await r.arrayBuffer());
                    return isFontCollection(bytes) ? bytes : new BundledFontData(bytes);
                  });
                  p.catch(() => fontCache.delete(url));
                  fontCache.set(url, p);
                }
                return p;
              },
            },
            'bundled',
          );
        }
      },
    });
  }
  return convertCore(data, options, providers);
}

/** Convert a .docx to PDF bytes in the browser. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer | Blob, options: BrowserConvertOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  return (await convertDocx(input, options)).pdf;
}

/** Extract a .docx's text as plain text or Markdown (`format: 'markdown'`), with warnings. */
export async function extractText(input: Uint8Array | ArrayBuffer | Blob, options: ExtractTextOptions = {}): Promise<ExtractTextResult> {
  return extractCore(input instanceof Blob ? new Uint8Array(await input.arrayBuffer()) : input, options);
}

/** A .docx's text as plain text. */
export async function docxToText(input: Uint8Array | ArrayBuffer | Blob, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'text' })).text;
}

/** A .docx's text as GitHub-flavoured Markdown. */
export async function docxToMarkdown(input: Uint8Array | ArrayBuffer | Blob, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'markdown' })).text;
}
