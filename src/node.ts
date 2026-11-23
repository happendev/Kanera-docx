import { open, readdir, readFile, stat } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { nodeCompression } from './util/node-compression.js';
import { convertToImagesWithCompression, convertWithCompression, type ConvertOptions, type ConvertResult, type FontProvider, type ImageOptions, type ImagesResult, type PageImage } from './core.js';
import { extractTextWithCompression, type ExtractTextOptions, type ExtractTextResult } from './text/index.js';
import { CJK_FILES, fontVersion, loadCjkFont } from './fonts/node-cjk.js';
import { FULL_FONTS } from './fonts/bundled.js';
import { BundledFontData, type FontSource } from './fonts/registry.js';
import { collectionSize, isFontCollection, MAX_FONT_FACES, SfntFont } from './fonts/sfnt.js';
import { buildSfnt } from './fonts/subset.js';

export type { ConvertOptions, ConvertResult, ExtractTextOptions, ExtractTextResult, FontSource, ImageFormat, ImageOptions, ImagesResult, LayoutStats, PageImage } from './core.js';
export type { TextFormat } from './text/extract.js';

export interface NodeFontOptions {
  /**
   * Also use fonts installed on this machine (default false, so output does not depend on the
   * host). `true` scans the usual font folders; an array scans those folders instead.
   */
  systemFonts?: boolean | string[];
  /** Folder holding the full font pack (defaults to the package's `fonts/`). */
  fontDir?: string;
}

export interface NodeConvertOptions extends ConvertOptions, NodeFontOptions {}

export interface NodeImageOptions extends ImageOptions, NodeFontOptions {}

const MAX_CACHED_FONTS = 16;
type CachedFont = BundledFontData | Uint8Array;
const packagedFonts = new Map<string, CachedFont>();
const loadingFonts = new Map<string, Promise<CachedFont>>();

async function loadPackagedFont(path: string, checkVersion: boolean, sparse: boolean): Promise<CachedFont> {
  const version = checkVersion || sparse ? fontVersion(await stat(path, { bigint: true })) : '';
  const key = path + '\0' + version + (sparse ? '\0outlines' : '');
  const cached = packagedFonts.get(key);
  if (cached) {
    packagedFonts.delete(key);
    packagedFonts.set(key, cached);
    return cached;
  }
  const loading = loadingFonts.get(key);
  if (loading) return loading;
  const read = sparse ? loadCjkFont(path, version) : readFile(path).then(data =>
    isFontCollection(data) ? data : new BundledFontData(data));
  const pending = read.then(font => {
    packagedFonts.set(key, font);
    if (packagedFonts.size > MAX_CACHED_FONTS) packagedFonts.delete(packagedFonts.keys().next().value!);
    return font;
  });
  loadingFonts.set(key, pending);
  try {
    return await pending;
  } finally {
    loadingFonts.delete(key);
  }
}

function bundledProvider(dir?: string): FontProvider {
  return {
    register(registry) {
      for (const b of FULL_FONTS) {
        const path = resolve(dir ?? packageFontDir(), b.file);
        const sparse = dir === undefined && !registry.needsGlyphBounds && CJK_FILES.has(b.file);
        const load = () => loadPackagedFont(path, dir !== undefined, sparse);
        registry.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, covers: b.covers, load }, 'bundled');
      }
    },
  };
}

const DEFAULT_FONT_DIRS = (): string[] => {
  const home = homedir();
  if (process.platform === 'win32') return [join(process.env.WINDIR ?? 'C:\\Windows', 'Fonts'), join(process.env.LOCALAPPDATA ?? join(home, 'AppData', 'Local'), 'Microsoft', 'Windows', 'Fonts')];
  if (process.platform === 'darwin') return ['/System/Library/Fonts', '/Library/Fonts', join(home, 'Library', 'Fonts')];
  return ['/usr/share/fonts', '/usr/local/share/fonts', join(home, '.fonts'), join(home, '.local', 'share', 'fonts')];
};

async function listFontFiles(dir: string, out: string[], depth = 0): Promise<void> {
  if (depth > 6) return;
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const e of entries) {
    const p = join(dir, e.name);
    if (e.isDirectory()) await listFontFiles(p, out, depth + 1);
    else if (/\.(ttf|otf|ttc)$/i.test(e.name)) out.push(p);
  }
}

/** Read just the tables needed to describe a font file, without loading its outlines. */
async function describeFontFile(path: string): Promise<FontSource[]> {
  const fd = await open(path, 'r');
  try {
    const { size } = await fd.stat();
    const read = async (offset: number, length: number): Promise<Uint8Array> => {
      if (offset + length > size || length > 64 * 1024 * 1024) throw new Error('Invalid font table size');
      const b = new Uint8Array(length);
      let done = 0;
      while (done < length) {
        const { bytesRead } = await fd.read(b, done, length - done, offset + done);
        if (!bytesRead) throw new Error('Truncated font table');
        done += bytesRead;
      }
      return b;
    };
    const head = await read(0, 12);
    const u32 = (d: Uint8Array, o: number) => ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0;
    const u16 = (d: Uint8Array, o: number) => (d[o] << 8) | d[o + 1];
    let dirs = [0];
    if (isFontCollection(head)) {
      const n = collectionSize(head, size);
      if (n > MAX_FONT_FACES) throw new Error('Font collection face limit exceeded');
      const offs = await read(12, n * 4);
      dirs = [...Array(n).keys()].map((i) => u32(offs, i * 4));
    }
    const out: FontSource[] = [];
    for (const [index, dir] of dirs.entries()) {
      const dh = await read(dir, 12);
      const numTables = u16(dh, 4);
      const recs = await read(dir + 12, numTables * 16);
      const tables = new Map<string, Uint8Array>();
      for (let i = 0; i < numTables; i++) {
        const r = i * 16;
        const tag = String.fromCharCode(recs[r], recs[r + 1], recs[r + 2], recs[r + 3]);
        if (!['head', 'hhea', 'name', 'OS/2', 'cmap', 'maxp'].includes(tag)) continue;
        tables.set(tag, await read(u32(recs, r + 8), u32(recs, r + 12)));
      }
      if (!tables.has('head') || !tables.has('hhea')) continue;
      const font = new SfntFont(buildSfnt(tables, u32(dh, 0)));
      const families = [...new Set([font.names.family, font.names.typographicFamily, font.names.fullName, ...font.names.all].filter((f): f is string => !!f))];
      out.push({
        family: families,
        bold: font.isBold,
        italic: font.isItalic,
        weight: font.weightClass,
        collectionIndex: dirs.length > 1 ? index : undefined,
        covers: font.coverageTest(),
        load: () => readFile(path),
      });
    }
    return out;
  } finally {
    await fd.close();
  }
}

const systemCache = new Map<string, Promise<FontSource[]>>();

async function scanSystemFonts(dirs: string[]): Promise<FontSource[]> {
  dirs = dirs.map((dir) => resolve(dir));
  const key = dirs.join('\u0000');
  let p = systemCache.get(key);
  if (!p) {
    p = (async () => {
      const files: string[] = [];
      for (const d of dirs) await listFontFiles(d, files);
      const out: FontSource[] = [];
      for (const f of files) {
        try {
          out.push(...await describeFontFile(f));
        } catch {
          // Unreadable or unsupported font files are skipped.
        }
      }
      return out;
    })();
  }
  systemCache.delete(key);
  systemCache.set(key, p);
  if (systemCache.size > 4) systemCache.delete(systemCache.keys().next().value!);
  return p;
}

function packageFontDir(): string {
  return fileURLToPath(new URL('../fonts/', import.meta.url));
}

async function fontProviders(options: NodeFontOptions): Promise<FontProvider[]> {
  const providers: FontProvider[] = [];
  if (options.systemFonts) {
    const dirs = Array.isArray(options.systemFonts) ? options.systemFonts : DEFAULT_FONT_DIRS();
    const sources = await scanSystemFonts(dirs);
    providers.push({
      register(registry) {
        for (const s of sources) registry.add(s, 'system');
      },
    });
  }
  providers.push(bundledProvider(options.fontDir));
  return providers;
}

/** Convert a .docx (bytes or a file path) to PDF, with the bundled fonts available. */
export async function convertDocx(input: Uint8Array | ArrayBuffer | string, options: NodeConvertOptions = {}): Promise<ConvertResult> {
  const data = typeof input === 'string' ? await readFile(input) : input;
  return convertWithCompression(data, options, await fontProviders(options), nodeCompression);
}

/** Render a .docx's (bytes or a file path) pages as PNG (default) or JPEG images, with the bundled fonts available. */
export async function convertDocxToImages(input: Uint8Array | ArrayBuffer | string, options: NodeImageOptions = {}): Promise<ImagesResult> {
  const data = typeof input === 'string' ? await readFile(input) : input;
  return convertToImagesWithCompression(data, options, await fontProviders(options), nodeCompression);
}

/** Render a .docx's (bytes or a file path) pages as PNG (default) or JPEG images. */
export async function docxToImages(input: Uint8Array | ArrayBuffer | string, options: NodeImageOptions = {}): Promise<PageImage[]> {
  return (await convertDocxToImages(input, options)).images;
}

/** Convert a .docx (bytes or a file path) to PDF bytes. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer | string, options: NodeConvertOptions = {}): Promise<Uint8Array<ArrayBuffer>> {
  return (await convertDocx(input, options)).pdf;
}

/** Extract a .docx's (bytes or a file path) text as plain text or Markdown (`format: 'markdown'`), with warnings. */
export async function extractText(input: Uint8Array | ArrayBuffer | string, options: ExtractTextOptions = {}): Promise<ExtractTextResult> {
  const data = typeof input === 'string' ? await readFile(input) : input;
  return extractTextWithCompression(data, options, nodeCompression);
}

/** A .docx's (bytes or a file path) text as plain text. */
export async function docxToText(input: Uint8Array | ArrayBuffer | string, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'text' })).text;
}

/** A .docx's (bytes or a file path) text as GitHub-flavoured Markdown. */
export async function docxToMarkdown(input: Uint8Array | ArrayBuffer | string, options: Omit<ExtractTextOptions, 'format'> = {}): Promise<string> {
  return (await extractText(input, { ...options, format: 'markdown' })).text;
}
