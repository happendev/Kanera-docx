import { convertDocx as convertCore, type ConvertOptions, type ConvertResult, type FontProvider } from './core.js';
import { BUNDLED_FONTS } from './fonts/bundled.js';

export type { ConvertOptions, ConvertResult, FontSource, LayoutStats } from './core.js';

export interface BrowserConvertOptions extends ConvertOptions {
  /**
   * URL of a folder serving the package's `fonts/` files. Defaults to the folder next to this
   * module, which works when the package is served as-is (not bundled). With a bundler, copy
   * `node_modules/@happendev/folio/fonts` to your static assets and point this at it.
   * Set to `false` to use only `options.fonts`.
   */
  fontBaseUrl?: string | URL | false;
  /** Custom fetch (for caching or auth). */
  fetch?: typeof fetch;
}

const fontCache = new Map<string, Promise<ArrayBuffer>>();

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
    const base = options.fontBaseUrl ? new URL(String(options.fontBaseUrl), globalThis.location?.href) : defaultBase();
    const doFetch = options.fetch ?? fetch;
    providers.push({
      register(registry) {
        for (const b of BUNDLED_FONTS) {
          const url = new URL(b.file, base).href;
          registry.add(
            {
              family: b.family,
              bold: b.bold,
              italic: b.italic,
              weight: b.weight,
              load: () => {
                let p = fontCache.get(url);
                if (!p) {
                  p = doFetch(url).then((r) => {
                    if (!r.ok) throw new Error(`HTTP ${r.status} for ${url}`);
                    return r.arrayBuffer();
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
export async function docxToPdf(input: Uint8Array | ArrayBuffer | Blob, options: BrowserConvertOptions = {}): Promise<Uint8Array> {
  return (await convertDocx(input, options)).pdf;
}
