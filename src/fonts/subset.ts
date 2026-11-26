import { glyphComponents, type SfntFont } from './sfnt.js';

export function canSubsetTrueType(font: SfntFont): boolean {
  if ((font.fsType & 0x100) || font.isCFF || (!font.table('glyf') && !font.readSubsetGlyphs) || !font.table('loca')) return false;
  const metrics = font.numberOfHMetrics;
  return metrics > 0 && metrics <= font.numGlyphs &&
    (font.table('hmtx')?.length ?? 0) >= 2 * (font.numGlyphs + metrics);
}

/** Glyph zero is reserved. Drawn glyphs follow, then composite dependencies. */
export function subsetTrueType(font: SfntFont, usedGlyphs: Iterable<number>, outlines?: ReadonlyMap<number, Uint8Array>): { data: Uint8Array; glyphIds: Map<number, number> } {
  const glyphIds = new Map<number, number>([[0, 0]]);
  for (const gid of usedGlyphs) {
    if (gid < 0 || gid >= font.numGlyphs) throw new Error('Invalid glyph ID for TrueType subset');
    if (!glyphIds.has(gid)) glyphIds.set(gid, glyphIds.size);
  }
  const pieces: Uint8Array[] = [];
  const components = new Map<number, ReturnType<typeof glyphComponents>>();
  let length = 0;
  // Iterating the map visits newly appended components as well, including nested composites.
  for (const gid of glyphIds.keys()) {
    const bytes = outlines ? outlines.get(gid)! : font.data.subarray(...font.glyphRange(gid)!);
    pieces.push(bytes);
    length += Math.ceil(bytes.length / 4) * 4;
    if (bytes.length < 10 || !(bytes[0] & 0x80)) continue;
    const children = glyphComponents(bytes);
    if (children.length) components.set(gid, children);
    for (const child of children) {
      if (child.gid >= font.numGlyphs) throw new Error('Invalid component ID for TrueType subset');
      if (!glyphIds.has(child.gid)) glyphIds.set(child.gid, glyphIds.size);
    }
  }

  const glyf = new Uint8Array(length);
  const glyph = new DataView(glyf.buffer);
  const loca = new Uint8Array((glyphIds.size + 1) * 4);
  const locaView = new DataView(loca.buffer);
  const hmtx = new Uint8Array(glyphIds.size * 4);
  const metrics = new DataView(hmtx.buffer);
  const original = font.tableBytes('hmtx')!;
  const originalMetrics = new DataView(original.buffer, original.byteOffset, original.byteLength);
  const longMetrics = font.numberOfHMetrics;
  let off = 0;
  for (const [gid, id] of glyphIds) {
    locaView.setUint32(id * 4, off);
    const bearing = gid < longMetrics ? gid * 4 + 2 : (longMetrics + gid) * 2;
    metrics.setUint16(id * 4, font.advance(gid));
    metrics.setInt16(id * 4 + 2, originalMetrics.getInt16(bearing));
    const bytes = pieces[id];
    glyf.set(bytes, off);
    const children = components.get(gid);
    if (children) for (const child of children) glyph.setUint16(off + child.offset, glyphIds.get(child.gid)!);
    off += Math.ceil(bytes.length / 4) * 4;
  }
  locaView.setUint32(glyphIds.size * 4, off);

  const head = new Uint8Array(font.tableBytes('head')!);
  // indexToLocFormat = 1 (long offsets); checkSumAdjustment zeroed until the end.
  head[50] = 0;
  head[51] = 1;
  head[8] = head[9] = head[10] = head[11] = 0;
  const hhea = new Uint8Array(font.tableBytes('hhea')!);
  new DataView(hhea.buffer).setUint16(34, glyphIds.size);
  const maxp = new Uint8Array(font.tableBytes('maxp')!);
  new DataView(maxp.buffer).setUint16(4, glyphIds.size);

  const tables = new Map<string, Uint8Array>();
  tables.set('head', head);
  tables.set('glyf', glyf);
  tables.set('loca', loca);
  tables.set('hhea', hhea);
  tables.set('hmtx', hmtx);
  tables.set('maxp', maxp);
  for (const tag of ['cvt ', 'fpgm', 'prep', 'OS/2', 'name']) {
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
  return { data: buildSfnt(tables, 0x00010000), glyphIds };
}

/** Sum padded, big-endian 32-bit words in the assembled sfnt. */
function checksum(view: DataView, start: number, end: number): number {
  let sum = 0;
  for (let i = start; i < end; i += 4) sum = (sum + view.getUint32(i)) >>> 0;
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
  let tableSum = 0;
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
    const sum = checksum(dv, offset, offset + data.length);
    dv.setUint32(rec + 4, sum);
    tableSum = (tableSum + sum) >>> 0;
    offset += (data.length + 3) & ~3;
  });
  if (headOffset >= 0) {
    const adj = (0xb1b0afba - tableSum - checksum(dv, 0, headerLen)) >>> 0;
    dv.setUint32(headOffset + 8, adj);
  }
  return out;
}
