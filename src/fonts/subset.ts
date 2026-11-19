import type { SfntFont } from './sfnt.js';

/**
 * Subset a TrueType-outline font for PDF embedding.
 *
 * Glyph ids are preserved (unused glyphs become empty), so the PDF can use an Identity
 * CIDToGIDMap and the layout's glyph ids stay valid. Returns null for CFF-outline fonts, which are
 * embedded whole instead.
 */
export function subsetTrueType(font: SfntFont, usedGlyphs: Iterable<number>): Uint8Array | null {
  if (font.isCFF || !font.table('glyf') || !font.table('loca')) return null;
  const keep = new Set<number>([0]);
  const queue = [...usedGlyphs];
  while (queue.length) {
    const g = queue.pop()!;
    if (g >= font.numGlyphs || keep.has(g)) continue;
    keep.add(g);
    for (const c of font.compositeComponents(g)) if (!keep.has(c)) queue.push(c);
  }

  // glyf + loca (long format)
  const pieces: Uint8Array[] = [];
  const loca = new Uint8Array((font.numGlyphs + 1) * 4);
  const locaView = new DataView(loca.buffer);
  let off = 0;
  for (let g = 0; g < font.numGlyphs; g++) {
    locaView.setUint32(g * 4, off);
    if (!keep.has(g)) continue;
    const range = font.glyphRange(g);
    if (!range) continue;
    const bytes = font.data.subarray(range[0], range[1]);
    if (bytes.length === 0) continue;
    pieces.push(bytes);
    off += bytes.length;
    if (off % 4) {
      const pad = 4 - (off % 4);
      pieces.push(new Uint8Array(pad));
      off += pad;
    }
  }
  locaView.setUint32(font.numGlyphs * 4, off);
  const glyf = new Uint8Array(off);
  let p = 0;
  for (const piece of pieces) {
    glyf.set(piece, p);
    p += piece.length;
  }

  const head = new Uint8Array(font.tableBytes('head')!);
  // indexToLocFormat = 1 (long offsets); checkSumAdjustment zeroed until the end.
  head[50] = 0;
  head[51] = 1;
  head[8] = head[9] = head[10] = head[11] = 0;

  const tables = new Map<string, Uint8Array>();
  tables.set('head', head);
  tables.set('glyf', glyf);
  tables.set('loca', loca);
  for (const tag of ['hhea', 'hmtx', 'maxp', 'cvt ', 'fpgm', 'prep', 'OS/2', 'name']) {
    const b = font.tableBytes(tag);
    if (b) tables.set(tag, b);
  }
  const post = font.tableBytes('post');
  if (post && post.length >= 32) {
    // Version 3.0 drops glyph names, which the PDF does not need.
    const p3 = new Uint8Array(post.subarray(0, 32));
    p3[0] = 0;
    p3[1] = 3;
    p3[2] = 0;
    p3[3] = 0;
    tables.set('post', p3);
  }
  return buildSfnt(tables, 0x00010000);
}

function checksum(d: Uint8Array): number {
  let sum = 0;
  const n = d.length;
  for (let i = 0; i < n; i += 4) {
    const v = ((d[i] << 24) | ((d[i + 1] ?? 0) << 16) | ((d[i + 2] ?? 0) << 8) | (d[i + 3] ?? 0)) >>> 0;
    sum = (sum + v) >>> 0;
  }
  return sum;
}

export function buildSfnt(tables: Map<string, Uint8Array>, version: number): Uint8Array {
  const tags = [...tables.keys()].sort();
  const numTables = tags.length;
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= numTables) {
    searchRange *= 2;
    entrySelector++;
  }
  searchRange *= 16;
  const headerLen = 12 + numTables * 16;
  let total = headerLen;
  for (const t of tags) total += (tables.get(t)!.length + 3) & ~3;
  const out = new Uint8Array(total);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, version);
  dv.setUint16(4, numTables);
  dv.setUint16(6, searchRange);
  dv.setUint16(8, entrySelector);
  dv.setUint16(10, numTables * 16 - searchRange);
  let offset = headerLen;
  let headOffset = -1;
  tags.forEach((tag, i) => {
    const data = tables.get(tag)!;
    const rec = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[rec + k] = tag.charCodeAt(k);
    dv.setUint32(rec + 8, offset);
    dv.setUint32(rec + 12, data.length);
    out.set(data, offset);
    if (tag === 'head') {
      headOffset = offset;
      // Rebuilding a collection face may copy an old adjustment into the new file.
      dv.setUint32(offset + 8, 0);
    }
    dv.setUint32(rec + 4, checksum(out.subarray(offset, offset + data.length)));
    offset += (data.length + 3) & ~3;
  });
  if (headOffset >= 0) {
    const adj = (0xb1b0afba - checksum(out)) >>> 0;
    dv.setUint32(headOffset + 8, adj);
  }
  return out;
}
