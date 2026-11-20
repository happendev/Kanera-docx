// Node-only loading for the packaged TrueType CJK fonts. Layout uses the small sfnt tables;
// PDF embedding opens the same file version and reads the outlines used by that conversion.
import type { BigIntStats } from 'node:fs';
import { open, type FileHandle } from 'node:fs/promises';
import { BundledFontData } from './registry.js';
import { glyphComponents, isFontCollection, type SfntFont, type TableRecord } from './sfnt.js';
import { buildSfnt, canSubsetTrueType } from './subset.js';

export const CJK_FILES = new Set(['NotoSansCJKjp-Regular.ttf', 'NotoSerifCJKsc-Regular.ttf']);

export function fontVersion(info: BigIntStats): string {
  return `${info.dev}:${info.ino}:${info.size}:${info.ctimeNs}:${info.mtimeNs}`;
}

async function readRange(file: FileHandle, offset: number, length: number): Promise<Buffer> {
  const bytes = Buffer.allocUnsafe(length);
  let done = 0;
  while (done < length) {
    const { bytesRead } = await file.read(bytes, done, length - done, offset + done);
    if (!bytesRead) throw new Error('Truncated bundled font');
    done += bytesRead;
  }
  return bytes;
}

async function checkVersion(file: FileHandle, expected: string): Promise<number> {
  const info = await file.stat({ bigint: true });
  if (fontVersion(info) !== expected) throw new Error('Bundled font changed during conversion');
  return Number(info.size);
}

async function readOutlines(path: string, version: string, font: SfntFont, glyf: TableRecord, gids: readonly number[]) {
  const file = await open(path, 'r');
  try {
    await checkVersion(file, version);
    const outlines = new Map<number, Uint8Array>();
    let pending = new Set([0, ...gids]);
    while (pending.size) {
      const ranges = Array.from(pending, gid => {
        const range = font.glyphRange(gid, glyf);
        if (!range) throw new Error('Invalid glyph in bundled font');
        return { gid, start: range[0], end: range[1] };
      }).sort((a, b) => a.start - b.start);
      pending = new Set();
      // Amortize I/O on adjacent glyphs without reading a multi-megabyte hole between them.
      for (let i = 0; i < ranges.length;) {
        const start = ranges[i].start;
        let end = ranges[i].end;
        let j = i + 1;
        while (j < ranges.length && ranges[j].start - end <= 4096 && ranges[j].end - start <= 256 * 1024) {
          end = Math.max(end, ranges[j++].end);
        }
        const bytes = await readRange(file, start, end - start);
        for (; i < j; i++) {
          const range = ranges[i];
          const outline = bytes.subarray(range.start - start, range.end - start);
          outlines.set(range.gid, outline);
          for (const child of glyphComponents(outline)) {
            if (!outlines.has(child.gid)) pending.add(child.gid);
          }
        }
      }
      for (const gid of pending) if (outlines.has(gid)) pending.delete(gid);
    }
    await checkVersion(file, version);
    return outlines;
  } finally {
    await file.close();
  }
}

export async function loadCjkFont(path: string, version: string): Promise<BundledFontData | Uint8Array> {
  const file = await open(path, 'r');
  try {
    const size = await checkVersion(file, version);
    const header = await readRange(file, 0, 12);
    const fullFont = async () => {
      const bytes = await file.readFile();
      await checkVersion(file, version);
      return isFontCollection(bytes) ? bytes : new BundledFontData(bytes);
    };
    if (isFontCollection(header)) return await fullFont();
    const count = header.readUInt16BE(4);
    const directory = await readRange(file, 12, count * 16);
    const tables = new Map<string, Uint8Array>();
    let glyf: TableRecord | undefined;
    let total = 0;
    for (let i = 0; i < directory.length; i += 16) {
      const tag = directory.toString('latin1', i, i + 4);
      const offset = directory.readUInt32BE(i + 8);
      const length = directory.readUInt32BE(i + 12);
      total += length;
      if (offset + length > size || total > size) throw new Error('Invalid bundled font table');
      if (tag === 'glyf') glyf = { offset, length };
      else tables.set(tag, await readRange(file, offset, length));
    }
    if (glyf) {
      const data = new BundledFontData(buildSfnt(tables, header.readUInt32BE(0)));
      const outlines = glyf;
      data.font.readSubsetGlyphs = gids => readOutlines(path, version, data.font, outlines, gids);
      if (canSubsetTrueType(data.font)) {
        await checkVersion(file, version);
        return data;
      }
    }
    // An updated font may require full embedding, e.g. its license disallows subsetting.
    return await fullFont();
  } finally {
    await file.close();
  }
}
