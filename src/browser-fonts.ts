import type { BrowserFontOptions } from './browser.js';
import type { FontProvider } from './core.js';
import { COMMON_FONTS, FULL_FONTS } from './fonts/bundled.js';
import { BundledFontData, BundledFontPack } from './fonts/registry.js';
import { isFontCollection } from './fonts/sfnt.js';

interface BrowserFontCache {
  fonts: Map<string, Promise<BundledFontData | Uint8Array>>;
  pack: BundledFontPack;
}

type ReadFont = (url: string) => Promise<Uint8Array>;

const fontCaches = new WeakMap<typeof fetch | ReadFont, BrowserFontCache>();

function defaultBase(): URL {
  // Kept out of a literal `new URL('…', import.meta.url)` so bundlers don't try to inline a folder.
  const rel = ['..', 'fonts', ''].join('/');
  return new URL(rel, import.meta.url);
}

export function fontProviders(options: BrowserFontOptions, readFont?: ReadFont): FontProvider[] {
  const providers: FontProvider[] = [];
  if (options.fontBaseUrl !== false) {
    const base = options.fontBaseUrl ? new URL(String(options.fontBaseUrl), globalThis.location?.href ?? import.meta.url) : defaultBase();
    if (!base.pathname.endsWith('/')) base.pathname += '/';
    const doFetch = options.fetch ?? fetch;
    const cacheKey = readFont ?? doFetch;
    let cache = fontCaches.get(cacheKey);
    if (!cache) {
      cache = { fonts: new Map(), pack: new BundledFontPack() };
      fontCaches.set(cacheKey, cache);
    }
    const { fonts, pack } = cache;
    const faces = options.fontPack === 'common' ? COMMON_FONTS : FULL_FONTS;
    providers.push({
      register(registry) {
        registry.addBundled(pack.sources(base.href, faces, file => {
          const url = new URL(file, base).href;
          return () => {
            let pending = fonts.get(url);
            if (!pending) {
              const bytes = readFont ? readFont(url) : fetchFont(url, doFetch);
              pending = bytes.then(data => isFontCollection(data) ? data : new BundledFontData(data));
              pending.catch(() => fonts.delete(url));
              fonts.set(url, pending);
            }
            return pending;
          };
        }));
      },
    });
  }
  return providers;
}

async function fetchFont(url: string, doFetch: typeof fetch): Promise<Uint8Array> {
  const response = await doFetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status} for ${url}`);
  return new Uint8Array(await response.arrayBuffer());
}
