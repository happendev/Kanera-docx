import { SfntFont, collectionSize } from './sfnt.js';
import {
  GENERIC_FALLBACKS,
  GLYPH_FALLBACKS,
  genericFamily,
  normalizeFamily,
  substituteFamilies,
  type FontTableHints,
  type GenericFamily,
} from './substitutes.js';
import { mapSymbolChar, symbolFontKind, type SymbolFontKind } from './symbol-map.js';

export type FontOrigin = 'embedded' | 'user' | 'system' | 'bundled';

/** A font supplied to the converter. Either `data` or `load` (with `family`) is required. */
export interface FontSource {
  data?: Uint8Array | ArrayBuffer;
  load?: () => Promise<Uint8Array | ArrayBuffer>;
  /** Family names this font answers to. Read from the font when `data` is given. */
  family?: string | string[];
  bold?: boolean;
  italic?: boolean;
  weight?: number;
  /** Index inside a .ttc collection. All fonts in a collection are added when omitted. */
  collectionIndex?: number;
  /** Optional coverage test used to pick unloaded fonts for glyph fallback. */
  covers?: (codePoint: number) => boolean;
}

let nextFaceId = 1;

export class FontFace {
  readonly id: string;
  readonly families: string[];
  readonly bold: boolean;
  readonly italic: boolean;
  readonly weight: number;
  /** Word (Windows GDI) line metrics in em units. */
  readonly ascent: number;
  readonly descent: number;
  readonly externalLeading: number;

  constructor(
    readonly font: SfntFont,
    readonly origin: FontOrigin,
    extraFamilies: string[] = [],
  ) {
    this.id = `F${nextFaceId++}`;
    const names = new Set<string>();
    for (const n of [font.names.family, font.names.typographicFamily, font.names.fullName, font.names.postscriptName, ...font.names.all, ...extraFamilies])
      if (n) names.add(normalizeFamily(n));
    this.families = [...names];
    this.weight = font.weightClass;
    this.bold = font.isBold;
    this.italic = font.isItalic;
    const upm = font.unitsPerEm;
    this.ascent = font.winAscent / upm;
    this.descent = font.winDescent / upm;
    const hheaHeight = font.hheaAscender - font.hheaDescender;
    this.externalLeading = Math.max(0, font.hheaLineGap - (font.winAscent + font.winDescent - hheaHeight)) / upm;
  }

  get familyName(): string {
    return this.font.names.typographicFamily ?? this.font.names.family;
  }
}

interface PendingFont {
  families: string[];
  bold: boolean;
  italic: boolean;
  weight: number;
  origin: FontOrigin;
  load: () => Promise<Uint8Array | ArrayBuffer>;
  collectionIndex?: number;
  covers?: (cp: number) => boolean;
  promise?: Promise<void>;
  loaded?: boolean;
}

export interface ResolvedFace {
  face: FontFace;
  syntheticBold: boolean;
  syntheticItalic: boolean;
  /** Set when characters must be remapped from a legacy symbol font encoding to Unicode. */
  symbolMap?: SymbolFontKind;
}

const ORIGIN_RANK: Record<FontOrigin, number> = { embedded: 0, user: 1, system: 2, bundled: 3 };

function toBytes(d: Uint8Array | ArrayBuffer): Uint8Array {
  return d instanceof Uint8Array ? d : new Uint8Array(d);
}

export class FontRegistry {
  private faces: FontFace[] = [];
  private pending: PendingFont[] = [];
  private byFamily = new Map<string, FontFace[]>();
  private resolveCache = new Map<string, ResolvedFace>();
  private glyphCache = new Map<string, ResolvedFace>();
  readonly hints = new Map<string, FontTableHints>();
  readonly warnings = new Set<string>();

  add(source: FontSource, origin: FontOrigin): void {
    const extra = source.family ? (Array.isArray(source.family) ? source.family : [source.family]) : [];
    if (source.data) {
      const data = toBytes(source.data);
      const count = collectionSize(data);
      const indices = source.collectionIndex !== undefined ? [source.collectionIndex] : [...Array(count).keys()];
      for (const idx of indices) {
        try {
          this.addFace(new FontFace(new SfntFont(data, idx), origin, extra));
        } catch (e) {
          this.warnings.add(`Could not read font ${extra[0] ?? ''}: ${(e as Error).message}`);
        }
      }
      return;
    }
    if (source.load && extra.length) {
      this.pending.push({
        families: extra.map(normalizeFamily),
        bold: !!source.bold,
        italic: !!source.italic,
        weight: source.weight ?? (source.bold ? 700 : 400),
        origin,
        load: source.load,
        collectionIndex: source.collectionIndex,
        covers: source.covers,
      });
      return;
    }
    throw new Error('FontSource needs `data`, or `load` together with `family`');
  }

  private addFace(face: FontFace): void {
    this.faces.push(face);
    for (const fam of face.families) {
      let list = this.byFamily.get(fam);
      if (!list) this.byFamily.set(fam, (list = []));
      list.push(face);
      list.sort((a, b) => ORIGIN_RANK[a.origin] - ORIGIN_RANK[b.origin]);
    }
    this.resolveCache.clear();
    this.glyphCache.clear();
  }

  setHints(family: string, hints: FontTableHints): void {
    this.hints.set(normalizeFamily(family), hints);
  }

  /** Families tried for a requested font, in order: exact, alternate name, substitutes, generic. */
  candidateFamilies(family: string): string[] {
    const norm = normalizeFamily(family);
    const out = [norm];
    const hint = this.hints.get(norm);
    if (hint?.altName) out.push(normalizeFamily(hint.altName));
    const styleless = norm.replace(/\s+(bold|italic|bold italic|regular)$/, '');
    if (styleless !== norm) out.push(styleless);
    const sub = substituteFamilies(norm) ?? (styleless !== norm ? substituteFamilies(styleless) : null);
    if (sub) out.push(...sub.families);
    out.push(...GENERIC_FALLBACKS[this.generic(family)]);
    if (this.generic(family) !== 'sans') out.push(...GENERIC_FALLBACKS.sans);
    return [...new Set(out)];
  }

  generic(family: string): GenericFamily {
    return genericFamily(family, this.hints.get(normalizeFamily(family)));
  }

  /** Load the fonts needed to render `families` (any style) plus the generic fallbacks. */
  async prepare(families: Iterable<string>, codePoints?: Iterable<number>): Promise<void> {
    const wanted = new Set<string>();
    for (const f of families) for (const c of this.candidateFamilies(f)) wanted.add(c);
    for (const f of GENERIC_FALLBACKS.sans) wanted.add(f);
    wanted.add('dejavu sans');
    // Stop at the first candidate available for each request so unrelated fonts are not fetched.
    const load: PendingFont[] = [];
    for (const f of families) {
      for (const c of this.candidateFamilies(f)) {
        const pend = this.pending.filter((p) => !p.loaded && p.families.includes(c));
        if (this.byFamily.has(c) && pend.length === 0) break;
        if (pend.length) {
          load.push(...pend);
          break;
        }
      }
    }
    for (const fallback of ['liberation sans', 'dejavu sans']) load.push(...this.pending.filter((p) => p.families.includes(fallback)));
    await Promise.all([...new Set(load)].map((p) => this.loadPending(p)));

    // Glyph coverage: load fallback fonts for characters the loaded fonts lack.
    if (codePoints) {
      const missing: number[] = [];
      for (const cp of codePoints) {
        if (cp < 0x20) continue;
        if (!this.faces.some((f) => f.font.hasGlyph(cp))) missing.push(cp);
      }
      if (missing.length) {
        const extra: PendingFont[] = [];
        for (const fam of GLYPH_FALLBACKS) {
          for (const p of this.pending) {
            if (p.loaded || !p.families.includes(fam)) continue;
            if (p.covers && !missing.some((cp) => p.covers!(cp))) continue;
            extra.push(p);
          }
        }
        // Coverage-aware pending fonts (system fonts) for anything still unaccounted for.
        for (const p of this.pending) {
          if (!p.loaded && p.covers && missing.some((cp) => p.covers!(cp)) && !extra.includes(p) && !p.bold && !p.italic) extra.push(p);
        }
        await Promise.all(extra.slice(0, 8).map((p) => this.loadPending(p)));
      }
    }
  }

  private loadPending(p: PendingFont): Promise<void> {
    if (!p.promise) {
      p.promise = (async () => {
        try {
          const data = toBytes(await p.load());
          const count = collectionSize(data);
          const indices = p.collectionIndex !== undefined ? [p.collectionIndex] : [...Array(count).keys()];
          for (const idx of indices) this.addFace(new FontFace(new SfntFont(data, idx), p.origin, p.families));
        } catch (e) {
          this.warnings.add(`Could not load font ${p.families[0]}: ${(e as Error).message}`);
        }
        p.loaded = true;
      })();
    }
    return p.promise;
  }

  private pickStyle(list: FontFace[], bold: boolean, italic: boolean): FontFace {
    const targetWeight = bold ? 700 : 400;
    let best = list[0];
    let bestScore = Infinity;
    for (const f of list) {
      let score = ORIGIN_RANK[f.origin] * 0.01;
      score += Math.abs(f.weight - targetWeight) / 100;
      if (f.bold !== bold) score += 3;
      if (f.italic !== italic) score += 5;
      if (score < bestScore) {
        bestScore = score;
        best = f;
      }
    }
    return best;
  }

  private styled(face: FontFace, bold: boolean, italic: boolean, symbolMap?: SymbolFontKind): ResolvedFace {
    return {
      face,
      syntheticBold: bold && !face.bold && face.weight < 600,
      syntheticItalic: italic && !face.italic,
      symbolMap,
    };
  }

  /** Resolve the face Word would use for a family/style, substituting when it is unavailable. */
  resolve(family: string, bold: boolean, italic: boolean): ResolvedFace {
    const key = `${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}`;
    const cached = this.resolveCache.get(key);
    if (cached) return cached;
    const norm = normalizeFamily(family);
    let result: ResolvedFace | undefined;
    // A family name that carries the style ("Arial Bold") selects that style.
    const styleBold = bold || /\bbold\b/.test(norm);
    const styleItalic = italic || /\b(italic|oblique)\b/.test(norm);
    const symKind = symbolFontKind(norm);
    for (const cand of this.candidateFamilies(family)) {
      const list = this.byFamily.get(cand);
      if (!list || list.length === 0) continue;
      const face = this.pickStyle(list, styleBold, styleItalic);
      // A real symbol font keeps its encoding; any other face needs characters remapped.
      const isRealSymbol = symKind !== null && face.font.symbolCmap;
      result = this.styled(face, styleBold, styleItalic, symKind && !isRealSymbol ? symKind : undefined);
      if (cand !== norm && symKind === null && !this.isMetricSubstitute(norm, cand)) this.warnings.add(`Font "${family}" not available; using ${face.familyName}`);
      break;
    }
    if (!result) {
      if (this.faces.length === 0) throw new Error('No fonts available. Supply fonts or use the node/browser entry point that loads bundled fonts.');
      result = this.styled(this.faces[0], bold, italic, symKind ?? undefined);
      this.warnings.add(`Font "${family}" not available; using ${this.faces[0].familyName}`);
    }
    this.resolveCache.set(key, result);
    return result;
  }

  private isMetricSubstitute(norm: string, cand: string): boolean {
    const sub = substituteFamilies(norm);
    return !!sub && sub.metricCompatible && sub.families.includes(cand);
  }

  /**
   * The face that renders `cp` for the requested font: the resolved face when it has the glyph,
   * otherwise a fallback font that does (like Word's font linking).
   */
  faceForChar(family: string, bold: boolean, italic: boolean, cp: number): ResolvedFace {
    const primary = this.resolve(family, bold, italic);
    const mapped = primary.symbolMap ? mapSymbolChar(primary.symbolMap, cp) : cp;
    if (primary.face.font.hasGlyph(mapped) || cp < 0x20 || cp === 0xa0 || cp === 0x2028) return primary;
    const key = `${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}\u0000${cp}`;
    const cached = this.glyphCache.get(key);
    if (cached) return cached;
    let result: ResolvedFace | undefined;
    const tryFamilies = [...GENERIC_FALLBACKS[this.generic(family)], ...GLYPH_FALLBACKS];
    for (const fam of tryFamilies) {
      const list = this.byFamily.get(fam);
      if (!list) continue;
      const withGlyph = list.filter((f) => f.font.hasGlyph(mapped));
      if (withGlyph.length) {
        result = this.styled(this.pickStyle(withGlyph, bold, italic), bold, italic, primary.symbolMap);
        break;
      }
    }
    if (!result) {
      const any = this.faces.find((f) => f.font.hasGlyph(mapped));
      result = any ? this.styled(any, bold, italic, primary.symbolMap) : primary;
      if (!any) this.warnings.add(`No font has a glyph for U+${mapped.toString(16).toUpperCase().padStart(4, '0')}`);
    }
    this.glyphCache.set(key, result);
    return result;
  }

  get loadedFaces(): readonly FontFace[] {
    return this.faces;
  }
}
