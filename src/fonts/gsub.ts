/**
 * The single substitutions (GSUB lookup type 1, including those inside type 7 extensions) that an
 * OpenType feature makes, by glyph id. Empty when the table has no such feature or is malformed.
 */
export function singleSubstitutions(gsub: Uint8Array, feature: string): Map<number, number> {
  const out = new Map<number, number>();
  const v = new DataView(gsub.buffer, gsub.byteOffset, gsub.byteLength);
  const u16 = (o: number) => v.getUint16(o);
  try {
    const features = u16(6);
    const lookupList = u16(8);
    const lookups = new Set<number>();
    for (let i = 0, n = u16(features); i < n; i++) {
      const rec = features + 2 + 6 * i;
      const tag = String.fromCharCode(v.getUint8(rec), v.getUint8(rec + 1), v.getUint8(rec + 2), v.getUint8(rec + 3));
      if (tag !== feature) continue;
      const f = features + u16(rec + 4);
      for (let k = 0, c = u16(f + 2); k < c; k++) lookups.add(u16(f + 4 + 2 * k));
    }
    for (const index of lookups) {
      const lookup = lookupList + u16(lookupList + 2 + 2 * index);
      for (let s = 0, n = u16(lookup + 4); s < n; s++) {
        let sub = lookup + u16(lookup + 6 + 2 * s);
        let type = u16(lookup);
        if (type === 7) {
          type = u16(sub + 2);
          sub += v.getUint32(sub + 4);
        }
        if (type !== 1) continue;
        const glyphs = coverage(v, sub + u16(sub + 2));
        if (u16(sub) === 1) {
          const delta = v.getInt16(sub + 4);
          for (const g of glyphs) out.set(g, (g + delta) & 0xffff);
        } else if (u16(sub) === 2) {
          const count = u16(sub + 4);
          glyphs.forEach((g, i) => {
            if (i < count) out.set(g, u16(sub + 6 + 2 * i));
          });
        }
      }
    }
  } catch {
    // A table that ends early (a RangeError) substitutes nothing.
    return new Map();
  }
  return out;
}

/** The glyphs of a coverage table, in coverage index order. */
function coverage(v: DataView, at: number): number[] {
  const out: number[] = [];
  const count = v.getUint16(at + 2);
  if (v.getUint16(at) === 1) {
    for (let i = 0; i < count; i++) out.push(v.getUint16(at + 4 + 2 * i));
  } else if (v.getUint16(at) === 2) {
    for (let i = 0; i < count; i++) {
      const r = at + 4 + 6 * i;
      const start = v.getUint16(r);
      const index = v.getUint16(r + 4);
      for (let g = start; g <= v.getUint16(r + 2); g++) out[index + g - start] = g;
    }
  }
  return out;
}
