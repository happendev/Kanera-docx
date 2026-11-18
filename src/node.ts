import { closeSync, fstatSync, openSync, readSync } from 'node:fs';
import { readdir, readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { convertDocx as convertCore, type ConvertOptions, type ConvertResult, type FontProvider } from './core.js';
import { BUNDLED_FONTS } from './fonts/bundled.js';
import type { FontRegistry, FontSource } from './fonts/registry.js';
import { SfntFont } from './fonts/sfnt.js';
import { buildSfnt } from './fonts/subset.js';

export type { ConvertOptions, ConvertResult, FontSource, LayoutStats } from './core.js';

export interface NodeConvertOptions extends ConvertOptions {
  /**
   * Also use fonts installed on this machine (default false, so output does not depend on the
   * host). `true` scans the usual font folders; an array scans those folders instead.
   */
  systemFonts?: boolean | string[];
  /** Folder holding the bundled fonts (defaults to the package's `fonts/`). */
  fontDir?: string;
}

function bundledProvider(dir: string): FontProvider {
  return {
    register(registry: FontRegistry) {
      for (const b of BUNDLED_FONTS) {
        registry.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, load: () => readFile(join(dir, b.file)) }, 'bundled');
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
function describeFontFile(path: string): FontSource[] {
  const fd = openSync(path, 'r');
  try {
    const size = fstatSync(fd).size;
    const read = (offset: number, length: number): Uint8Array => {
      const b = new Uint8Array(Math.max(0, Math.min(length, size - offset)));
      if (b.length) readSync(fd, b, 0, b.length, offset);
      return b;
    };
    const head = read(0, 12);
    const u32 = (d: Uint8Array, o: number) => ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0;
    const u16 = (d: Uint8Array, o: number) => (d[o] << 8) | d[o + 1];
    let dirs = [0];
    if (head[0] === 0x74 && head[1] === 0x74 && head[2] === 0x63 && head[3] === 0x66) {
      const n = u32(head, 8);
      const offs = read(12, n * 4);
      dirs = [...Array(n).keys()].map((i) => u32(offs, i * 4));
    }
    const out: FontSource[] = [];
    dirs.forEach((dir, index) => {
      const dh = read(dir, 12);
      const numTables = u16(dh, 4);
      const recs = read(dir + 12, numTables * 16);
      const tables = new Map<string, Uint8Array>();
      for (let i = 0; i < numTables; i++) {
        const r = i * 16;
        const tag = String.fromCharCode(recs[r], recs[r + 1], recs[r + 2], recs[r + 3]);
        if (!['head', 'hhea', 'name', 'OS/2', 'cmap', 'maxp'].includes(tag)) continue;
        tables.set(tag, read(u32(recs, r + 8), u32(recs, r + 12)));
      }
      if (!tables.has('head') || !tables.has('hhea')) return;
      const font = new SfntFont(buildSfnt(tables, u32(dh, 0)));
      const families = [...new Set([font.names.family, font.names.typographicFamily, font.names.fullName, ...font.names.all].filter((f): f is string => !!f))];
      out.push({
        family: families,
        bold: font.isBold,
        italic: font.isItalic,
        weight: font.weightClass,
        collectionIndex: dirs.length > 1 ? index : undefined,
        covers: (cp) => font.hasGlyph(cp),
        load: () => readFile(path),
      });
    });
    return out;
  } finally {
    closeSync(fd);
  }
}

const systemCache = new Map<string, Promise<FontSource[]>>();

async function scanSystemFonts(dirs: string[]): Promise<FontSource[]> {
  const key = dirs.join('\u0000');
  let p = systemCache.get(key);
  if (!p) {
    p = (async () => {
      const files: string[] = [];
      for (const d of dirs) await listFontFiles(d, files);
      const out: FontSource[] = [];
      for (const f of files) {
        try {
          out.push(...describeFontFile(f));
        } catch {
          // Unreadable or unsupported font files are skipped.
        }
      }
      return out;
    })();
    systemCache.set(key, p);
  }
  return p;
}

function packageFontDir(): string {
  return fileURLToPath(new URL('../fonts/', import.meta.url));
}

/** Convert a .docx (bytes or a file path) to PDF, with the bundled fonts available. */
export async function convertDocx(input: Uint8Array | ArrayBuffer | string, options: NodeConvertOptions = {}): Promise<ConvertResult> {
  const data = typeof input === 'string' ? await readFile(input) : input;
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
  providers.push(bundledProvider(options.fontDir ?? packageFontDir()));
  return convertCore(data instanceof Uint8Array ? data : new Uint8Array(data), options, providers);
}

/** Convert a .docx (bytes or a file path) to PDF bytes. */
export async function docxToPdf(input: Uint8Array | ArrayBuffer | string, options: NodeConvertOptions = {}): Promise<Uint8Array> {
  return (await convertDocx(input, options)).pdf;
}
