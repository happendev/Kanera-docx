/**
 * TrueType / OpenType (sfnt) parser: just the tables needed for Word-compatible text measurement
 * and PDF embedding.
 */

export interface TableRecord {
  offset: number;
  length: number;
}

export interface FontNames {
  family: string;
  subfamily: string;
  fullName: string;
  postscriptName: string;
  typographicFamily?: string;
  typographicSubfamily?: string;
  /** Every family-like name in every language, for matching. */
  all: string[];
}

class Reader {
  constructor(public d: Uint8Array) {}
  u8(o: number): number {
    return this.d[o];
  }
  u16(o: number): number {
    return (this.d[o] << 8) | this.d[o + 1];
  }
  i16(o: number): number {
    const v = this.u16(o);
    return v & 0x8000 ? v - 0x10000 : v;
  }
  u32(o: number): number {
    return ((this.d[o] << 24) | (this.d[o + 1] << 16) | (this.d[o + 2] << 8) | this.d[o + 3]) >>> 0;
  }
  tag(o: number): string {
    return String.fromCharCode(this.d[o], this.d[o + 1], this.d[o + 2], this.d[o + 3]);
  }
}

export function isFontCollection(data: Uint8Array): boolean {
  return data.length > 12 && data[0] === 0x74 && data[1] === 0x74 && data[2] === 0x63 && data[3] === 0x66;
}

export function collectionSize(data: Uint8Array): number {
  if (!isFontCollection(data)) return 1;
  return new Reader(data).u32(8);
}

const MAC_ROMAN_HIGH =
  'ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ';

export class SfntFont {
  readonly r: Reader;
  readonly tables = new Map<string, TableRecord>();
  readonly sfntVersion: number;
  /** Offset of this font's table directory (non-zero inside a collection). */
  readonly dirOffset: number;

  unitsPerEm = 1000;
  numGlyphs = 0;
  indexToLocFormat = 0;
  macStyle = 0;
  bbox: [number, number, number, number] = [0, 0, 0, 0];

  hheaAscender = 0;
  hheaDescender = 0;
  hheaLineGap = 0;
  numberOfHMetrics = 0;

  winAscent = 0;
  winDescent = 0;
  typoAscender = 0;
  typoDescender = 0;
  typoLineGap = 0;
  fsSelection = 0;
  fsType = 0;
  weightClass = 400;
  capHeight = 0;
  xHeight = 0;
  strikeoutSize = 0;
  strikeoutPosition = 0;
  superscriptYSize = 0;
  superscriptYOffset = 0;
  subscriptYSize = 0;
  subscriptYOffset = 0;
  avgCharWidth = 0;
  panose: number[] = [];
  hasOS2 = false;

  italicAngle = 0;
  underlinePosition = 0;
  underlineThickness = 0;
  isFixedPitch = false;

  advances!: Uint16Array;
  names!: FontNames;
  private cmap = new Map<number, number>();
  /** True when the only usable cmap is the (3,0) symbol cmap with codes in U+F000..U+F0FF. */
  symbolCmap = false;
  private kernPairs: Map<number, number> | null = null;

  constructor(
    readonly data: Uint8Array,
    collectionIndex = 0,
  ) {
    this.r = new Reader(data);
    let dir = 0;
    if (isFontCollection(data)) {
      const n = this.r.u32(8);
      if (collectionIndex >= n) throw new Error('Font collection index out of range');
      dir = this.r.u32(12 + collectionIndex * 4);
    }
    this.dirOffset = dir;
    this.sfntVersion = this.r.u32(dir);
    if (this.r.tag(dir) === 'wOFF' || this.r.tag(dir) === 'wOF2') throw new Error('WOFF fonts are not supported; supply TTF/OTF');
    const numTables = this.r.u16(dir + 4);
    for (let i = 0; i < numTables; i++) {
      const rec = dir + 12 + i * 16;
      this.tables.set(this.r.tag(rec), { offset: this.r.u32(rec + 8), length: this.r.u32(rec + 12) });
    }
    this.parseHead();
    this.parseMaxp();
    this.parseHhea();
    this.parseHmtx();
    this.parseOS2();
    this.parsePost();
    this.names = this.parseNames();
    this.parseCmap();
  }

  get isCFF(): boolean {
    return this.tables.has('CFF ') || this.tables.has('CFF2');
  }

  /** OS/2 fsSelection USE_TYPO_METRICS: line metrics come from the typo fields, not win. */
  get useTypoMetrics(): boolean {
    return this.hasOS2 && (this.fsSelection & 0x80) !== 0;
  }

  get isBold(): boolean {
    if (this.hasOS2) return this.weightClass >= 600 || (this.fsSelection & 0x20) !== 0;
    return (this.macStyle & 1) !== 0;
  }

  get isItalic(): boolean {
    if (this.hasOS2 && (this.fsSelection & 0x01) !== 0) return true;
    return (this.macStyle & 2) !== 0 || this.italicAngle !== 0;
  }

  table(tag: string): TableRecord | undefined {
    return this.tables.get(tag);
  }

  tableBytes(tag: string): Uint8Array | undefined {
    const t = this.tables.get(tag);
    return t ? this.data.subarray(t.offset, t.offset + t.length) : undefined;
  }

  private parseHead(): void {
    const t = this.table('head');
    if (!t) throw new Error('Font has no head table');
    const o = t.offset;
    this.unitsPerEm = this.r.u16(o + 18) || 1000;
    this.bbox = [this.r.i16(o + 36), this.r.i16(o + 38), this.r.i16(o + 40), this.r.i16(o + 42)];
    this.macStyle = this.r.u16(o + 44);
    this.indexToLocFormat = this.r.i16(o + 50);
  }

  private parseMaxp(): void {
    const t = this.table('maxp');
    this.numGlyphs = t ? this.r.u16(t.offset + 4) : 0;
  }

  private parseHhea(): void {
    const t = this.table('hhea');
    if (!t) throw new Error('Font has no hhea table');
    const o = t.offset;
    this.hheaAscender = this.r.i16(o + 4);
    this.hheaDescender = this.r.i16(o + 6);
    this.hheaLineGap = this.r.i16(o + 8);
    this.numberOfHMetrics = this.r.u16(o + 34);
  }

  private parseHmtx(): void {
    const t = this.table('hmtx');
    this.advances = new Uint16Array(Math.max(this.numGlyphs, 1));
    if (!t) return;
    let last = 0;
    for (let g = 0; g < this.numGlyphs; g++) {
      if (g < this.numberOfHMetrics) last = this.r.u16(t.offset + g * 4);
      this.advances[g] = last;
    }
  }

  private parseOS2(): void {
    const t = this.table('OS/2');
    if (!t || t.length < 68) {
      this.winAscent = this.hheaAscender;
      this.winDescent = -this.hheaDescender;
      this.typoAscender = this.hheaAscender;
      this.typoDescender = this.hheaDescender;
      this.typoLineGap = this.hheaLineGap;
      this.capHeight = Math.round(this.unitsPerEm * 0.7);
      this.xHeight = Math.round(this.unitsPerEm * 0.5);
      return;
    }
    this.hasOS2 = true;
    const o = t.offset;
    const version = this.r.u16(o);
    this.avgCharWidth = this.r.i16(o + 2);
    this.weightClass = this.r.u16(o + 4);
    this.fsType = this.r.u16(o + 8);
    this.subscriptYSize = this.r.i16(o + 12);
    this.subscriptYOffset = this.r.i16(o + 16);
    this.superscriptYSize = this.r.i16(o + 20);
    this.superscriptYOffset = this.r.i16(o + 24);
    this.strikeoutSize = this.r.i16(o + 26);
    this.strikeoutPosition = this.r.i16(o + 28);
    this.panose = Array.from(this.data.subarray(o + 32, o + 42));
    this.fsSelection = this.r.u16(o + 62);
    this.typoAscender = this.r.i16(o + 68);
    this.typoDescender = this.r.i16(o + 70);
    this.typoLineGap = this.r.i16(o + 72);
    this.winAscent = this.r.u16(o + 74);
    this.winDescent = this.r.u16(o + 76);
    if (version >= 2 && t.length >= 96) {
      this.xHeight = this.r.i16(o + 86);
      this.capHeight = this.r.i16(o + 88);
    }
    if (!this.capHeight) this.capHeight = Math.round(this.unitsPerEm * 0.7);
    if (!this.xHeight) this.xHeight = Math.round(this.unitsPerEm * 0.5);
    if (!this.winAscent && !this.winDescent) {
      this.winAscent = this.hheaAscender;
      this.winDescent = -this.hheaDescender;
    }
  }

  private parsePost(): void {
    const t = this.table('post');
    if (!t) return;
    const o = t.offset;
    this.italicAngle = this.r.i16(o + 4) + this.r.u16(o + 6) / 65536;
    this.underlinePosition = this.r.i16(o + 8);
    this.underlineThickness = this.r.i16(o + 10);
    this.isFixedPitch = this.r.u32(o + 12) !== 0;
  }

  private parseNames(): FontNames {
    const t = this.table('name');
    const best = new Map<number, { score: number; value: string }>();
    const all = new Set<string>();
    if (t) {
      const o = t.offset;
      const count = this.r.u16(o + 2);
      const strOff = o + this.r.u16(o + 4);
      for (let i = 0; i < count; i++) {
        const rec = o + 6 + i * 12;
        const platform = this.r.u16(rec);
        const encoding = this.r.u16(rec + 2);
        const lang = this.r.u16(rec + 4);
        const nameId = this.r.u16(rec + 6);
        const len = this.r.u16(rec + 8);
        const off = strOff + this.r.u16(rec + 10);
        if (![1, 2, 4, 6, 16, 17, 21, 22].includes(nameId)) continue;
        let value: string | null = null;
        if (platform === 3 || platform === 0) {
          if (platform === 3 && encoding !== 1 && encoding !== 10 && encoding !== 0) continue;
          let s = '';
          for (let k = 0; k + 1 < len; k += 2) s += String.fromCharCode(this.r.u16(off + k));
          value = s;
        } else if (platform === 1 && encoding === 0) {
          let s = '';
          for (let k = 0; k < len; k++) {
            const b = this.data[off + k];
            s += b < 128 ? String.fromCharCode(b) : MAC_ROMAN_HIGH[b - 128] ?? '?';
          }
          value = s;
        }
        if (value === null) continue;
        value = value.replace(/\0/g, '').trim();
        if (!value) continue;
        if (nameId === 1 || nameId === 4 || nameId === 6 || nameId === 16 || nameId === 21) all.add(value);
        const score = (platform === 3 ? 4 : platform === 0 ? 2 : 1) + (lang === 0x409 || (platform === 1 && lang === 0) ? 8 : 0);
        const prev = best.get(nameId);
        if (!prev || score > prev.score) best.set(nameId, { score, value });
      }
    }
    const get = (id: number) => best.get(id)?.value;
    const family = get(1) ?? get(16) ?? get(4) ?? 'Unknown';
    return {
      family,
      subfamily: get(2) ?? 'Regular',
      fullName: get(4) ?? family,
      postscriptName: (get(6) ?? family).replace(/[^\x21-\x7e]|[[\](){}<>/%]/g, ''),
      typographicFamily: get(16) ?? get(21),
      typographicSubfamily: get(17) ?? get(22),
      all: [...all],
    };
  }

  private parseCmap(): void {
    const t = this.table('cmap');
    if (!t) return;
    const o = t.offset;
    const n = this.r.u16(o + 2);
    type Sub = { platform: number; encoding: number; offset: number; format: number };
    const subs: Sub[] = [];
    for (let i = 0; i < n; i++) {
      const rec = o + 4 + i * 8;
      const platform = this.r.u16(rec);
      const encoding = this.r.u16(rec + 2);
      const offset = o + this.r.u32(rec + 4);
      subs.push({ platform, encoding, offset, format: this.r.u16(offset) });
    }
    const rank = (s: Sub): number => {
      if (s.platform === 3 && s.encoding === 10) return 10;
      if (s.platform === 0 && (s.encoding === 4 || s.encoding === 6)) return 9;
      if (s.platform === 3 && s.encoding === 1) return 8;
      if (s.platform === 0) return 7;
      if (s.platform === 3 && s.encoding === 0) return 5;
      if (s.platform === 1 && s.encoding === 0) return 2;
      return 0;
    };
    const usable = subs.filter((s) => rank(s) > 0 && [0, 4, 6, 12, 13].includes(s.format)).sort((a, b) => rank(b) - rank(a));
    const chosen = usable[0];
    if (!chosen) return;
    this.symbolCmap = chosen.platform === 3 && chosen.encoding === 0;
    this.readCmapSubtable(chosen.offset, chosen.format, chosen.platform === 1);
  }

  private readCmapSubtable(o: number, format: number, macRoman: boolean): void {
    const map = this.cmap;
    const r = this.r;
    if (format === 0) {
      for (let c = 0; c < 256; c++) {
        const g = r.u8(o + 6 + c);
        if (!g) continue;
        const cp = macRoman && c >= 128 ? MAC_ROMAN_HIGH.charCodeAt(c - 128) : c;
        map.set(cp, g);
      }
    } else if (format === 4) {
      const segX2 = r.u16(o + 6);
      const ends = o + 14;
      const starts = ends + segX2 + 2;
      const deltas = starts + segX2;
      const rangeOffs = deltas + segX2;
      for (let s = 0; s < segX2; s += 2) {
        const end = r.u16(ends + s);
        const start = r.u16(starts + s);
        const delta = r.i16(deltas + s);
        const ro = r.u16(rangeOffs + s);
        if (start === 0xffff) continue;
        for (let c = start; c <= end; c++) {
          let g: number;
          if (ro === 0) g = (c + delta) & 0xffff;
          else {
            const addr = rangeOffs + s + ro + (c - start) * 2;
            if (addr + 1 >= this.data.length) continue;
            g = r.u16(addr);
            if (g) g = (g + delta) & 0xffff;
          }
          if (g) map.set(c, g);
        }
      }
    } else if (format === 6) {
      const first = r.u16(o + 6);
      const count = r.u16(o + 8);
      for (let i = 0; i < count; i++) {
        const g = r.u16(o + 10 + i * 2);
        if (g) map.set(first + i, g);
      }
    } else if (format === 12 || format === 13) {
      const groups = r.u32(o + 12);
      for (let i = 0; i < groups; i++) {
        const g0 = o + 16 + i * 12;
        const start = r.u32(g0);
        const end = Math.min(r.u32(g0 + 4), start + 0x20000);
        const gid = r.u32(g0 + 8);
        for (let c = start; c <= end; c++) map.set(c, format === 12 ? gid + (c - start) : gid);
      }
    }
  }

  /** Glyph id for a Unicode code point (0 = .notdef / missing). */
  glyphId(cp: number): number {
    const g = this.cmap.get(cp);
    if (g !== undefined) return g;
    if (this.symbolCmap) {
      if (cp < 0x100) return this.cmap.get(0xf000 + cp) ?? 0;
      if (cp >= 0xf000 && cp <= 0xf0ff) return this.cmap.get(cp - 0xf000) ?? 0;
    }
    return 0;
  }

  hasGlyph(cp: number): boolean {
    return this.glyphId(cp) !== 0;
  }

  advance(gid: number): number {
    return gid < this.advances.length ? this.advances[gid] : this.advances[this.advances.length - 1];
  }

  /** Byte range of a glyph inside the glyf table (TrueType outlines only). */
  glyphRange(gid: number): [number, number] | null {
    const loca = this.table('loca');
    const glyf = this.table('glyf');
    if (!loca || !glyf || gid >= this.numGlyphs) return null;
    let start: number;
    let end: number;
    if (this.indexToLocFormat === 0) {
      start = this.r.u16(loca.offset + gid * 2) * 2;
      end = this.r.u16(loca.offset + gid * 2 + 2) * 2;
    } else {
      start = this.r.u32(loca.offset + gid * 4);
      end = this.r.u32(loca.offset + gid * 4 + 4);
    }
    if (end <= start) return [glyf.offset + start, glyf.offset + start];
    return [glyf.offset + start, glyf.offset + end];
  }

  /** Components referenced by a composite glyph. */
  compositeComponents(gid: number): number[] {
    const range = this.glyphRange(gid);
    if (!range || range[1] - range[0] < 10) return [];
    const [start] = range;
    if (this.r.i16(start) >= 0) return [];
    const out: number[] = [];
    let p = start + 10;
    for (;;) {
      const flags = this.r.u16(p);
      out.push(this.r.u16(p + 2));
      p += 4;
      p += flags & 0x0001 ? 4 : 2;
      if (flags & 0x0008) p += 2;
      else if (flags & 0x0040) p += 4;
      else if (flags & 0x0080) p += 8;
      if (!(flags & 0x0020)) break;
    }
    return out;
  }

  /** Pair kerning in font units (legacy `kern` table, then GPOS pair adjustments). */
  kerning(left: number, right: number): number {
    if (!this.kernPairs) this.kernPairs = this.parseKerning();
    return this.kernPairs.get(left * 65536 + right) ?? 0;
  }

  private parseKerning(): Map<number, number> {
    const pairs = new Map<number, number>();
    const t = this.table('kern');
    if (t) {
      const o = t.offset;
      const version = this.r.u16(o);
      if (version === 0) {
        const nTables = this.r.u16(o + 2);
        let p = o + 4;
        for (let i = 0; i < nTables; i++) {
          const len = this.r.u16(p + 2);
          const coverage = this.r.u16(p + 4);
          const format = coverage >> 8;
          if (format === 0 && (coverage & 1) === 1) {
            const nPairs = this.r.u16(p + 6);
            for (let k = 0; k < nPairs; k++) {
              const q = p + 14 + k * 6;
              pairs.set(this.r.u16(q) * 65536 + this.r.u16(q + 2), this.r.i16(q + 4));
            }
          }
          p += len;
        }
      }
    }
    this.parseGposPairs(pairs);
    return pairs;
  }

  private parseGposPairs(pairs: Map<number, number>): void {
    const t = this.table('GPOS');
    if (!t) return;
    const r = this.r;
    const base = t.offset;
    const featureList = base + r.u16(base + 6);
    const lookupList = base + r.u16(base + 8);
    const kernLookups = new Set<number>();
    const fCount = r.u16(featureList);
    for (let i = 0; i < fCount; i++) {
      const rec = featureList + 2 + i * 6;
      if (r.tag(rec) !== 'kern') continue;
      const feat = featureList + r.u16(rec + 4);
      const n = r.u16(feat + 2);
      for (let k = 0; k < n; k++) kernLookups.add(r.u16(feat + 4 + k * 2));
    }
    const coverage = (off: number): number[] => {
      const fmt = r.u16(off);
      const out: number[] = [];
      if (fmt === 1) {
        const n = r.u16(off + 2);
        for (let i = 0; i < n; i++) out.push(r.u16(off + 4 + i * 2));
      } else if (fmt === 2) {
        const n = r.u16(off + 2);
        for (let i = 0; i < n; i++) {
          const rr = off + 4 + i * 6;
          const s = r.u16(rr);
          const e = r.u16(rr + 2);
          for (let g = s; g <= e; g++) out.push(g);
        }
      }
      return out;
    };
    const classDef = (off: number): Map<number, number> => {
      const m = new Map<number, number>();
      const fmt = r.u16(off);
      if (fmt === 1) {
        const start = r.u16(off + 2);
        const n = r.u16(off + 4);
        for (let i = 0; i < n; i++) m.set(start + i, r.u16(off + 6 + i * 2));
      } else if (fmt === 2) {
        const n = r.u16(off + 2);
        for (let i = 0; i < n; i++) {
          const rr = off + 4 + i * 6;
          const s = r.u16(rr);
          const e = r.u16(rr + 2);
          const c = r.u16(rr + 4);
          for (let g = s; g <= e; g++) m.set(g, c);
        }
      }
      return m;
    };
    const valueSize = (vf: number): number => {
      let n = 0;
      for (let b = 0; b < 8; b++) if (vf & (1 << b)) n += 2;
      return n;
    };
    const xAdvance = (p: number, vf: number): number => {
      // XAdvance is bit 2; preceded by XPlacement (bit 0) and YPlacement (bit 1).
      if (!(vf & 4)) return 0;
      let off = 0;
      if (vf & 1) off += 2;
      if (vf & 2) off += 2;
      return r.i16(p + off);
    };
    const handleSubtable = (st: number) => {
      const fmt = r.u16(st);
      const cov = coverage(st + r.u16(st + 2));
      const vf1 = r.u16(st + 4);
      const vf2 = r.u16(st + 6);
      const s1 = valueSize(vf1);
      const s2 = valueSize(vf2);
      if (fmt === 1) {
        const setCount = r.u16(st + 8);
        for (let i = 0; i < setCount && i < cov.length; i++) {
          const ps = st + r.u16(st + 10 + i * 2);
          const n = r.u16(ps);
          for (let k = 0; k < n; k++) {
            const pr = ps + 2 + k * (2 + s1 + s2);
            const second = r.u16(pr);
            const adj = xAdvance(pr + 2, vf1);
            const key = cov[i] * 65536 + second;
            if (adj && !pairs.has(key)) pairs.set(key, adj);
          }
        }
      } else if (fmt === 2) {
        const cd1 = classDef(st + r.u16(st + 8));
        const cd2 = classDef(st + r.u16(st + 10));
        const c1Count = r.u16(st + 12);
        const c2Count = r.u16(st + 14);
        const rec = 2 * 0 + s1 + s2;
        // Class-pair matrices can be large; expand only for glyphs with a nonzero adjustment.
        const byClass2 = new Map<number, number[]>();
        for (const [g, c] of cd2) {
          let arr = byClass2.get(c);
          if (!arr) byClass2.set(c, (arr = []));
          arr.push(g);
        }
        for (const g1 of cov) {
          const c1 = cd1.get(g1) ?? 0;
          if (c1 >= c1Count) continue;
          for (let c2 = 1; c2 < c2Count; c2++) {
            const p = st + 16 + (c1 * c2Count + c2) * rec;
            const adj = xAdvance(p, vf1);
            if (!adj) continue;
            for (const g2 of byClass2.get(c2) ?? []) {
              const key = g1 * 65536 + g2;
              if (!pairs.has(key)) pairs.set(key, adj);
            }
          }
        }
      }
    };
    for (const li of kernLookups) {
      const lookup = lookupList + r.u16(lookupList + 2 + li * 2);
      const type = r.u16(lookup);
      const n = r.u16(lookup + 4);
      for (let k = 0; k < n; k++) {
        let st = lookup + r.u16(lookup + 6 + k * 2);
        let stType = type;
        if (type === 9) {
          stType = r.u16(st + 2);
          st = st + r.u32(st + 4);
        }
        if (stType === 2) handleSubtable(st);
      }
    }
  }
}
