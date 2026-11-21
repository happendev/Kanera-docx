import { type FontBudget, FontLimitError } from './sfnt.js';

export type FeatureBudget = Pick<FontBudget, 'reserveFeatureWork'>;

/**
 * The single substitutions (GSUB lookup type 1, including those inside type 7 extensions) that an
 * OpenType feature makes, by glyph id. Empty when the table has no such feature or is malformed.
 */
export function singleSubstitutions(gsub: Uint8Array, feature: string, budget: FeatureBudget, numGlyphs = 65536): Map<number, number> {
  const out = new Map<number, number>();
  const v = new DataView(gsub.buffer, gsub.byteOffset, gsub.byteLength);
  const u16 = (o: number) => v.getUint16(o);
  try {
    for (const sub of featureSubtables(v, feature, 1, 7, budget)) {
      const glyphs = coverage(v, sub + u16(sub + 2), budget);
      if (u16(sub) === 1) {
        const delta = v.getInt16(sub + 4);
        for (const g of glyphs) {
          const target = (g + delta) & 0xffff;
          if (g < numGlyphs && target < numGlyphs) out.set(g, target);
        }
      } else if (u16(sub) === 2) {
        const count = u16(sub + 4);
        glyphs.forEach((g, i) => {
          if (i >= count || g >= numGlyphs) return;
          const target = u16(sub + 6 + 2 * i);
          if (target < numGlyphs) out.set(g, target);
        });
      }
    }
  } catch (error) {
    if (error instanceof FontLimitError) throw error;
    // A table that ends early (a RangeError) substitutes nothing.
    return new Map();
  }
  return out;
}

/** A ligature that replaces its first glyph and the glyphs in `rest` that follow it. */
export interface Ligature {
  rest: number[];
  glyph: number;
}

/**
 * The ligatures (GSUB lookup type 4, including those inside type 7 extensions) that an OpenType
 * feature makes, by first glyph id, longest first. Empty when the table has no such feature or is
 * malformed.
 */
export function ligatureSubstitutions(gsub: Uint8Array, feature: string, budget: FeatureBudget, numGlyphs = 65536): Map<number, Ligature[]> {
  const out = new Map<number, Ligature[]>();
  const v = new DataView(gsub.buffer, gsub.byteOffset, gsub.byteLength);
  const u16 = (o: number) => v.getUint16(o);
  try {
    for (const sub of featureSubtables(v, feature, 4, 7, budget)) {
      if (u16(sub) !== 1) continue;
      const glyphs = coverage(v, sub + u16(sub + 2), budget);
      const sets = u16(sub + 4);
      budget.reserveFeatureWork(sets);
      for (let i = 0; i < sets && i < glyphs.length; i++) {
        if (glyphs[i] >= numGlyphs) continue;
        const set = sub + u16(sub + 6 + 2 * i);
        const count = u16(set);
        budget.reserveFeatureWork(count);
        let list = out.get(glyphs[i]);
        if (!list) out.set(glyphs[i], (list = []));
        for (let k = 0; k < count; k++) {
          const lig = set + u16(set + 2 + 2 * k);
          const glyph = u16(lig);
          const components = u16(lig + 2);
          budget.reserveFeatureWork(components);
          const rest: number[] = [];
          for (let c = 1; c < components; c++) rest.push(u16(lig + 2 + 2 * c));
          if (glyph < numGlyphs && rest.every((g) => g < numGlyphs)) list.push({ rest, glyph });
        }
      }
    }
  } catch (error) {
    if (error instanceof FontLimitError) throw error;
    return new Map();
  }
  for (const list of out.values()) list.sort((a, b) => b.rest.length - a.rest.length);
  return out;
}

/**
 * The horizontal single adjustments (GPOS lookup type 1, including those inside type 9 extensions)
 * that an OpenType feature makes, by glyph id, in font units: `x` moves the glyph, `advance` changes
 * its advance. Empty when the table has no such feature or is malformed.
 */
export function singleAdjustments(gpos: Uint8Array, feature: string, budget: FeatureBudget): Map<number, { x: number; advance: number }> {
  const out = new Map<number, { x: number; advance: number }>();
  const v = new DataView(gpos.buffer, gpos.byteOffset, gpos.byteLength);
  const u16 = (o: number) => v.getUint16(o);
  // A value record holds the fields its format's bits name, two bytes each, in bit order.
  const value = (format: number, at: number) => {
    let x = 0;
    let advance = 0;
    for (let bit = 0, o = at; bit < 8; bit++) {
      if (!(format & (1 << bit))) continue;
      if (bit === 0) x = v.getInt16(o);
      if (bit === 2) advance = v.getInt16(o);
      o += 2;
    }
    return { x, advance };
  };
  const size = (format: number) => 2 * [...(format & 0xff).toString(2)].filter((b) => b === '1').length;
  try {
    for (const sub of featureSubtables(v, feature, 1, 9, budget)) {
      const glyphs = coverage(v, sub + u16(sub + 2), budget);
      const format = u16(sub + 4);
      if (!(format & 0x5)) continue;
      if (u16(sub) === 1) {
        const adjust = value(format, sub + 6);
        for (const g of glyphs) if (!out.has(g)) out.set(g, adjust);
      } else if (u16(sub) === 2) {
        const count = u16(sub + 6);
        const stride = size(format);
        glyphs.forEach((g, i) => {
          if (i < count && !out.has(g)) out.set(g, value(format, sub + 8 + i * stride));
        });
      }
    }
  } catch (error) {
    if (error instanceof FontLimitError) throw error;
    return new Map();
  }
  return out;
}

/** Offsets of the subtables of `type` in a feature's lookups, through extension lookups of type `extension`. */
function featureSubtables(v: DataView, feature: string, type: number, extension: number, budget: FeatureBudget): number[] {
  const u16 = (o: number) => v.getUint16(o);
  const features = u16(6);
  const lookupList = u16(8);
  const lookups = new Set<number>();
  budget.reserveFeatureWork(u16(features));
  for (let i = 0, n = u16(features); i < n; i++) {
    const rec = features + 2 + 6 * i;
    const tag = String.fromCharCode(v.getUint8(rec), v.getUint8(rec + 1), v.getUint8(rec + 2), v.getUint8(rec + 3));
    if (tag !== feature) continue;
    const f = features + u16(rec + 4);
    budget.reserveFeatureWork(u16(f + 2));
    for (let k = 0, c = u16(f + 2); k < c; k++) lookups.add(u16(f + 4 + 2 * k));
  }
  const out: number[] = [];
  for (const index of lookups) {
    if (index >= u16(lookupList)) throw new Error('Invalid font lookup index');
    const lookup = lookupList + u16(lookupList + 2 + 2 * index);
    budget.reserveFeatureWork(u16(lookup + 4));
    for (let s = 0, n = u16(lookup + 4); s < n; s++) {
      let sub = lookup + u16(lookup + 6 + 2 * s);
      let kind = u16(lookup);
      if (kind === extension) {
        kind = u16(sub + 2);
        sub += v.getUint32(sub + 4);
      }
      if (kind === type) out.push(sub);
    }
  }
  return out;
}

/** The glyphs of a coverage table, in coverage index order. */
function coverage(v: DataView, at: number, budget: FeatureBudget): number[] {
  const out: number[] = [];
  const count = v.getUint16(at + 2);
  budget.reserveFeatureWork(count);
  if (v.getUint16(at) === 1) {
    for (let i = 0; i < count; i++) out.push(v.getUint16(at + 4 + 2 * i));
  } else if (v.getUint16(at) === 2) {
    for (let i = 0; i < count; i++) {
      const r = at + 4 + 6 * i;
      const start = v.getUint16(r);
      const index = v.getUint16(r + 4);
      const end = v.getUint16(r + 2);
      if (end < start || index !== out.length || start <= (out[index - 1] ?? -1)) throw new Error('Invalid font coverage');
      budget.reserveFeatureWork(end - start + 1);
      for (let g = start; g <= end; g++) out.push(g);
    }
  }
  return out;
}
