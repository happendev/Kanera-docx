import type { FontRequest } from '../docx/model.js';
import { SfntFont, collectionSize, FontBudget, FontLimitError } from './sfnt.js';
import {
  ARABIC_SAMPLE,
  GENERIC_FALLBACKS,
  arabicWidth,
  cjkFallbackLine,
  cjkFallbacks,
  cjkLineMetrics,
  NOTO_SANS_CJK_LINE,
  genericFamily,
  glyphFallbacks,
  halfWidthLatin,
  normalizeFamily,
  ownArabicHebrew,
  substituteFamilies,
  unstyledFamily,
  type FontTableHints,
  type GenericFamily,
} from './substitutes.js';
import { singleSubstitutions } from './gsub.js';
import { mapSymbolChar, mathLetter, symbolFontKind, symbolLookalike, type SymbolFontKind } from './symbol-map.js';

export type FontOrigin = 'embedded' | 'user' | 'system' | 'bundled';

type FontBytes = Uint8Array | ArrayBuffer;

interface FontSourceMetadata {
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

type FontSourceOptions<LoadedFont> = FontSourceMetadata & (
  | { data: FontBytes; load?: () => Promise<LoadedFont> }
  | { data?: undefined; family: string | string[]; load: () => Promise<LoadedFont> }
);

/** A font supplied to the converter as bytes or as a loader with a family name. */
export type FontSource = FontSourceOptions<FontBytes>;

/** Parsed once for an immutable packaged font. Each conversion still owns its face and budget. */
export class BundledFontData {
  readonly cost = new FontBudget();
  readonly font: SfntFont;

  constructor(data: Uint8Array) {
    this.font = new SfntFont(data, 0, this.cost);
  }
}

type LoadedFont = FontBytes | BundledFontData;

let nextFaceId = 1;

/** Spaces and hyphens the shaper draws with the face's own space or hyphen when it lacks them, so the
 * text keeps its font (and line metrics) rather than falling back to another. */
const SYNTHESIZED_SPACES = new Set([0x2002, 0x2003, 0x2005, 0x2006, 0x2007, 0x2009, 0x202f]);
function synthesized(font: SfntFont, cp: number): boolean {
  if (cp === 0x2010 || cp === 0x2011) return font.hasGlyph(0x2d);
  return SYNTHESIZED_SPACES.has(cp) && font.hasGlyph(0x20);
}

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
    private sharedKerning?: { cost: FontBudget; target: FontBudget },
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
    if (font.useTypoMetrics) {
      // Fonts that set USE_TYPO_METRICS (Aptos, for one) are laid out with their typographic metrics.
      this.ascent = font.typoAscender / upm;
      this.descent = -font.typoDescender / upm;
      this.externalLeading = Math.max(0, font.typoLineGap) / upm;
    } else {
      this.ascent = font.winAscent / upm;
      this.descent = font.winDescent / upm;
      const hheaHeight = font.hheaAscender - font.hheaDescender;
      this.externalLeading = Math.max(0, font.hheaLineGap - (font.winAscent + font.winDescent - hheaHeight)) / upm;
    }
  }

  kerning(left: number, right: number): number {
    const value = this.font.kerning(left, right);
    if (this.sharedKerning) {
      this.sharedKerning.cost.chargeKerningTo(this.sharedKerning.target);
      this.sharedKerning = undefined;
    }
    return value;
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
  load: () => Promise<LoadedFont>;
  collectionIndex?: number;
  covers?: (cp: number) => boolean;
  promise?: Promise<void>;
  loaded?: boolean;
}

/**
 * Word draws symbols the requested font lacks (check boxes, stars, dingbats) from Segoe UI Symbol,
 * on Windows and Mac alike, and their line takes its height: 2210 + 514 of 2048 units.
 */
const SEGOE_SYMBOL: [number, number] = [2210 / 2048, 514 / 2048];
function symbolFallback(cp: number): boolean {
  // Miscellaneous Symbols and Dingbats. Other symbol blocks are in some Office text fonts and not
  // others, so they keep the line metrics of whichever font draws them.
  return cp >= 0x2600 && cp <= 0x27bf;
}

export interface ResolvedFace {
  face: FontFace;
  /** Line metrics in em: the face's own, or the requested font's when a substitute stands in. */
  ascent: number;
  descent: number;
  externalLeading: number;
  syntheticBold: boolean;
  syntheticItalic: boolean;
  /** Set when characters must be remapped from a legacy symbol font encoding to Unicode. */
  symbolMap?: SymbolFontKind;
  /** The character drawn in place of the requested one (the plain letter of a mathematical letter). */
  char?: number;
  /**
   * Set when a stand-in draws a fixed-pitch East Asian font: ASCII is then half an em wide, drawn
   * with the stand-in's half-width forms (its OpenType hwid substitutions, by glyph id).
   */
  halfWidth?: Map<number, number>;
  /** With halfWidth: the stand-in's full-width forms (fwid), for the punctuation such a font draws an em wide. */
  fullWidth?: Map<number, number>;
  /** Glyph size relative to the text's, for a fallback drawn to the width of the font it stands in for. */
  glyphScale?: number;
}

const ORIGIN_RANK: Record<FontOrigin, number> = { embedded: 1, user: 0, system: 2, bundled: 3 };

function toBytes(d: FontBytes): Uint8Array {
  return d instanceof Uint8Array ? d : new Uint8Array(d);
}

function fontErrorMessage(error: unknown): string {
  try {
    return error instanceof Error ? error.message : String(error);
  } catch {
    return 'Unknown font error';
  }
}

export class FontRegistry {
  private readonly budget = new FontBudget();
  private faces: FontFace[] = [];
  private pending: PendingFont[] = [];
  private byFamily = new Map<string, FontFace[]>();
  private resolveCache = new Map<string, ResolvedFace>();
  private glyphCache = new Map<string, ResolvedFace>();
  private fallbackLines = new WeakMap<ResolvedFace, ResolvedFace>();
  private sampleWidths = new WeakMap<FontFace, number>();
  private widthFormCache = new WeakMap<FontFace, { half: Map<number, number>; full: Map<number, number> }>();
  private symbolCache = new Map<string, number>();
  readonly hints = new Map<string, FontTableHints>();
  readonly warnings = new Set<string>();

  add(source: FontSourceOptions<LoadedFont>, origin: FontOrigin): void {
    const extra = source.family ? (Array.isArray(source.family) ? source.family : [source.family]) : [];
    if (source.data) {
      this.readFaces(toBytes(source.data), source.collectionIndex, origin, extra);
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

  private readFaces(data: Uint8Array, index: number | undefined, origin: FontOrigin, families: string[]): void {
    const warn = (e: unknown) => {
      if (e instanceof FontLimitError) throw e;
      this.warnings.add(`Could not read font ${families[0] ?? ''}: ${fontErrorMessage(e)}`);
    };
    let count: number;
    try {
      count = collectionSize(data);
    } catch (e) {
      warn(e);
      return;
    }
    const first = index ?? 0;
    const faces = index === undefined ? count : 1;
    this.budget.reserveFaces(faces);
    for (let i = 0; i < faces; i++) {
      try {
        this.addFace(new FontFace(new SfntFont(data, first + i, this.budget), origin, families));
      } catch (e) {
        warn(e);
      }
    }
  }

  private addFace(face: FontFace): void {
    const rights = face.font.fsType;
    if ((rights & 0x0e) === 2 || (rights & 0x200)) {
      this.warnings.add('Font "' + face.familyName + '" disallows outline embedding and was skipped');
      return;
    }
    this.faces.push(face);
    for (const fam of face.families) {
      let list = this.byFamily.get(fam);
      if (!list) this.byFamily.set(fam, (list = []));
      list.push(face);
      list.sort((a, b) => ORIGIN_RANK[a.origin] - ORIGIN_RANK[b.origin]);
    }
    this.resolveCache.clear();
    this.glyphCache.clear();
    this.symbolCache.clear();
  }

  setHints(family: string, hints: FontTableHints): void {
    this.hints.set(normalizeFamily(family), hints);
  }

  /**
   * Families tried for a requested font, in order: exact, alternate name, substitutes, generic.
   * `lang` is the text's East Asian language, which picks the region of East Asian fallbacks.
   */
  candidateFamilies(family: string, lang?: string): string[] {
    const norm = normalizeFamily(family);
    const out = [norm];
    const alt = this.altName(norm);
    if (alt) out.push(alt);
    const styleless = unstyledFamily(norm);
    if (styleless !== norm) out.push(styleless);
    const sub = substituteFamilies(norm);
    if (sub) out.push(...sub.families);
    // Word falls back to the font table's alternate name, so that font's substitutes stand in too.
    const altSub = alt ? substituteFamilies(alt) : null;
    if (altSub) out.push(...altSub.families);
    const generic = this.generic(family);
    out.push(...(generic === 'cjk' ? this.cjk(family, lang) : GENERIC_FALLBACKS[generic]));
    if (generic !== 'sans') out.push(...GENERIC_FALLBACKS.sans);
    return [...new Set(out)];
  }

  private altName(norm: string): string | undefined {
    const raw = this.hints.get(norm)?.altName;
    const alt = raw ? normalizeFamily(raw) : undefined;
    if (!alt || alt === norm) return undefined;
    // Word for Mac fetches Segoe UI from Office's cloud fonts only for a document that sets text in
    // it, which lists it in the font table. Otherwise the alternate name finds nothing, and a Myriad
    // Pro whose alternate is Segoe UI comes out in Arial, the sans fallback. (Segoe UI Black, a
    // separate font, is followed: Word's PDFs draw it for "Signifier Black".)
    if (alt === 'segoe ui' && !this.hints.has(alt)) return undefined;
    return alt;
  }

  /** Line metrics of the substitute family `cand` for `norm` (or for its alternate name). */
  private substituteMetrics(norm: string, cand: string): [number, number, number] | undefined {
    for (const name of [norm, this.altName(norm)]) {
      const sub = name && cand !== name ? substituteFamilies(name) : null;
      if (!sub?.vmetrics || !sub.families.includes(cand)) continue;
      // A font reached by its alternate name has no line gap: Word for Mac sets text it draws in
      // Helvetica for "Helvetica, Arial, sans-serif" 1.0 em apart, and text asking for Helvetica 1.2.
      return name === norm ? sub.vmetrics : [sub.vmetrics[0], sub.vmetrics[1], 0];
    }
    if (cand === norm) return undefined;
    return cjkLineMetrics(norm) ?? (this.generic(norm) === 'cjk' && /^(noto|source han) serif/.test(cand) ? NOTO_SANS_CJK_LINE : undefined);
  }

  private cjk(family: string, lang?: string): string[] {
    return cjkFallbacks(family, lang, this.hints.get(normalizeFamily(family)));
  }

  /** Families searched for characters a font lacks, East Asian ones by region and style. */
  private glyphFallbacks(family: string, lang?: string): string[] {
    const generic = this.generic(family);
    const cjk = this.cjk(family, lang);
    return [...new Set([...(generic === 'cjk' ? cjk : GENERIC_FALLBACKS[generic]), ...glyphFallbacks(cjk)])];
  }

  generic(family: string): GenericFamily {
    return genericFamily(family, this.hints.get(normalizeFamily(family)));
  }

  /**
   * Load the closest face of a candidate family, retrying failed or restricted sources. With
   * `missing`, a loaded face that lacks those characters doesn't end the search among faces of the
   * same style: two sources of one family can cover different characters.
   */
  private async prepareFamily(family: string, bold: boolean, italic: boolean, missing?: number[]): Promise<boolean> {
    let candidates = this.pending.filter((p) => !p.loaded && p.families.includes(family) && (!missing || !p.covers || missing.some(p.covers)));
    // As in resolve(), a font passed in `fonts` under this name shuts out the other sources.
    const callerOnly = !missing && this.callerWins([...(this.byFamily.get(family) ?? []), ...candidates], bold, italic);
    if (callerOnly) candidates = candidates.filter((p) => p.origin === 'user');
    // The legacy family name wins next: "Aptos" loads Aptos, not Aptos Display, which shares the
    // typographic family and the same weight.
    const legacy = candidates.filter((p) => p.families[0] === family);
    if (legacy.length) candidates = legacy;
    candidates.sort((a, b) => this.styleScore(a, bold, italic) - this.styleScore(b, bold, italic));
    for (const p of candidates) {
      const loaded = (this.byFamily.get(family) ?? []).filter((face) => !callerOnly || face.origin === 'user');
      const styled = loaded.filter((face) => !legacy.length || normalizeFamily(face.font.names.family) === family);
      if (missing) {
        const found = missing.every((cp) => loaded.some((face) => face.font.hasGlyph(cp)));
        const best = Math.min(...styled.map((face) => this.styleDistance(face, bold, italic)));
        const distance = this.styleDistance(p, bold, italic);
        if (best < distance || (found && best <= distance)) break;
      } else if (styled.some((face) => this.styleScore(face, bold, italic) <= this.styleScore(p, bold, italic))) break;
      await this.loadPending(p);
    }
    return this.byFamily.has(family);
  }

  /**
   * Whether a font passed in `fonts` replaces the other sources' faces of a family (whatever its own
   * name table says), as it does unless it can't render the style: only an italic face for upright
   * text, say.
   */
  private callerWins(faces: Array<FontFace | PendingFont>, bold: boolean, italic: boolean): boolean {
    return faces.some((f) => f.origin === 'user' && (!f.italic || italic) && (!f.bold || bold));
  }

  async prepare(requests: FontRequest[]): Promise<void> {
    for (const { family, bold, italic, lang } of requests) {
      const b = bold || /\bbold\b/i.test(family);
      const i = italic || /\b(italic|oblique)\b/i.test(family);
      for (const candidate of this.candidateFamilies(family, lang)) {
        if (!(await this.prepareFamily(candidate, b, i))) continue;
        if (this.lacksStyle(family, candidate, b, i)) continue;
        break;
      }
    }

    for (const { family, bold, italic, codePoints, lang } of requests) {
      const primary = this.resolve(family, bold, italic, lang);
      const kind = primary.symbolMap;
      const missing = [...codePoints].map((cp) => kind ? mapSymbolChar(kind, cp) || cp : cp)
        .filter((cp) => cp >= 32 && !primary.face.font.hasGlyph(cp) && !synthesized(primary.face.font, cp));
      const unfound = await this.loadFallbacks(family, bold, italic, missing, lang);
      // A dingbat no font has can still be drawn with an older look-alike.
      const lookalikes = kind ? unfound.map(symbolLookalike).filter((cp): cp is number => cp !== undefined) : [];
      if (lookalikes.length) await this.loadFallbacks(family, bold, italic, lookalikes, lang);
      // So can a mathematical letter, as its plain letter in the style it stands for.
      const letters = kind ? [] : unfound.map(mathLetter).filter((m) => m !== undefined);
      for (const m of letters) await this.prepare([{ family, bold: bold || m.bold, italic: italic || m.italic, codePoints: new Set([m.cp]), lang }]);
    }
  }

  /** Load fallback fonts for characters no loaded face has; returns those still missing. */
  private async loadFallbacks(family: string, bold: boolean, italic: boolean, missing: number[], lang?: string): Promise<number[]> {
    missing = missing.filter((cp) => !this.faces.some((face) => face.font.hasGlyph(cp)));
    for (const fallback of this.glyphFallbacks(family, lang)) {
      if (!missing.length) break;
      if (await this.prepareFamily(fallback, bold, italic, missing)) {
        missing = missing.filter((cp) => !this.byFamily.get(fallback)!.some((face) => face.font.hasGlyph(cp)));
      }
    }
    for (const p of this.pending) {
      if (!missing.length) break;
      if (!p.loaded && p.covers && missing.some(p.covers)) {
        await this.loadPending(p);
        missing = missing.filter((cp) => !this.faces.some((face) => face.font.hasGlyph(cp)));
      }
    }
    return missing;
  }

  private loadPending(p: PendingFont): Promise<void> {
    if (!p.promise) {
      p.promise = (async () => {
        try {
          const data = await p.load();
          if (data instanceof BundledFontData) {
            this.budget.reserveFaces(1);
            data.cost.chargeMappingsTo(this.budget);
            this.addFace(new FontFace(data.font, p.origin, p.families, { cost: data.cost, target: this.budget }));
          } else {
            this.readFaces(toBytes(data), p.collectionIndex, p.origin, p.families);
          }
        } catch (e) {
          if (e instanceof FontLimitError) throw e;
          this.warnings.add(`Could not load font ${p.families[0]}: ${fontErrorMessage(e)}`);
        }
        p.loaded = true;
      })();
    }
    return p.promise;
  }

  private styleDistance(face: FontFace | PendingFont, bold: boolean, italic: boolean): number {
    return Math.abs(face.weight - (bold ? 700 : 400)) / 100 + (face.bold !== bold ? 3 : 0) + (face.italic !== italic ? 5 : 0);
  }

  private styleScore(face: FontFace | PendingFont, bold: boolean, italic: boolean): number {
    return this.styleDistance(face, bold, italic) + ORIGIN_RANK[face.origin] * 0.01;
  }

  private pickStyle(list: FontFace[], bold: boolean, italic: boolean): FontFace {
    return list.reduce((best, face) => this.styleScore(face, bold, italic) < this.styleScore(best, bold, italic) ? face : best);
  }

  private styled(face: FontFace, bold: boolean, italic: boolean, symbolMap?: SymbolFontKind, vmetrics?: [number, number, number]): ResolvedFace {
    return {
      face,
      ascent: vmetrics ? vmetrics[0] : face.ascent,
      descent: vmetrics ? vmetrics[1] : face.descent,
      externalLeading: vmetrics ? vmetrics[2] : face.externalLeading,
      syntheticBold: bold && !face.bold && face.weight < 600,
      syntheticItalic: italic && !face.italic,
      symbolMap,
    };
  }

  /** Resolve the face Word would use for a family/style, substituting when it is unavailable. */
  resolve(family: string, bold: boolean, italic: boolean, lang?: string): ResolvedFace {
    const key = `${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}\u0000${lang ?? ''}`;
    const cached = this.resolveCache.get(key);
    if (cached) return cached;
    const norm = normalizeFamily(family);
    let result: ResolvedFace | undefined;
    // A family name that carries the style ("Arial Bold") selects that style.
    const styleBold = bold || /\bbold\b/.test(norm);
    const styleItalic = italic || /\b(italic|oblique)\b/.test(norm);
    const symKind = symbolFontKind(norm);
    for (const cand of this.candidateFamilies(family, lang)) {
      const list = this.byFamily.get(cand);
      if (!list || list.length === 0) continue;
      if (this.lacksStyle(family, cand, styleBold, styleItalic, true)) continue;
      // A font passed in `fonts` under this name replaces the other sources' faces, whatever its own
      // name says. Among the rest, Word matches the legacy family name: "Aptos" is Aptos and Aptos
      // Bold, not Aptos Display or Aptos Light, which only share the typographic family.
      const source = this.callerWins(list, styleBold, styleItalic) ? list.filter((f) => f.origin === 'user') : list;
      const legacy = source.filter((f) => normalizeFamily(f.font.names.family) === cand);
      const face = this.pickStyle(legacy.length ? legacy : source, styleBold, styleItalic);
      // A real symbol font keeps its encoding; any other face needs characters remapped.
      const isRealSymbol = symKind !== null && face.font.symbolCmap;
      result = this.styled(face, styleBold, styleItalic, symKind && !isRealSymbol ? symKind : undefined, this.substituteMetrics(norm, cand));
      if (cand !== norm && halfWidthLatin(norm)) {
        const { half, full } = this.widthForms(face);
        if (half.size) result.halfWidth = half;
        if (full.size) result.fullWidth = full;
      }
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

  /**
   * True when `cand` is the requested family itself but only has faces that can't render the style
   * (an italic or bold face can't be made upright or lighter), while a metric-compatible substitute
   * can: only Cambria Bold and Italic installed, say, renders regular Cambria text with Caladea.
   * With `loaded`, the substitute must already have such a face.
   */
  private lacksStyle(family: string, cand: string, bold: boolean, italic: boolean, loaded = false): boolean {
    const norm = normalizeFamily(family);
    const fits = (f: FontFace | PendingFont) => (!f.italic || italic) && (!f.bold || bold);
    if (cand !== norm || (this.byFamily.get(cand) ?? []).some(fits)) return false;
    const sub = substituteFamilies(norm);
    if (!sub?.compatible.length) return false;
    return sub.compatible.some((f) => (this.byFamily.get(f) ?? []).some(fits) || (!loaded && this.pending.some((p) => !p.loaded && p.families.includes(f) && fits(p))));
  }

  private isMetricSubstitute(norm: string, cand: string): boolean {
    const sub = substituteFamilies(norm);
    return !!sub && sub.compatible.includes(cand);
  }

  /**
   * The face that renders `cp` for the requested font: the resolved face when it has the glyph,
   * otherwise a fallback font that does (like Word's font linking).
   */
  faceForChar(family: string, bold: boolean, italic: boolean, cp: number, lang?: string, drawn = false, primary = this.resolve(family, bold, italic, lang)): ResolvedFace {
    const mapped = primary.symbolMap ? this.symbolChar(primary.symbolMap, cp) : cp;
    if (primary.face.font.hasGlyph(mapped) || cp < 0x20 || cp === 0xa0 || cp === 0x2028) return this.eastAsianFallbackLine(family, cp, primary);
    if (synthesized(primary.face.font, cp)) return primary;
    const key = `${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}\u0000${cp}\u0000${lang ?? ''}\u0000${drawn ? 1 : 0}`;
    const cached = this.glyphCache.get(key);
    if (cached) return cached;
    let result: ResolvedFace | undefined;
    let found = false;
    for (const fam of this.glyphFallbacks(family, lang)) {
      const list = this.byFamily.get(fam);
      if (!list) continue;
      const withGlyph = list.filter((f) => f.font.hasGlyph(mapped));
      if (withGlyph.length) {
        result = this.fallback(primary, this.pickStyle(withGlyph, bold, italic), bold, italic);
        // A Latin, Greek or Cyrillic letter that a stand-in lacks is one the real font most likely
        // has (Arial Black has Cyrillic, Archivo Black hasn't), so its line keeps the real font's
        // metrics.
        if (cp < 0x530 && (primary.ascent !== primary.face.ascent || primary.descent !== primary.face.descent)) {
          result = { ...result, ascent: primary.ascent, descent: primary.descent, externalLeading: primary.externalLeading };
        }
        found = true;
        break;
      }
    }
    if (!result) {
      const any = this.anyFaceWith(family, mapped, bold, italic);
      const math = any || primary.symbolMap ? undefined : mathLetter(cp);
      // Word draws a mathematical letter from Cambria Math; without it, the plain letter in the
      // style it stands for looks the same.
      if (math) result = { ...this.faceForChar(family, bold || math.bold, italic || math.italic, math.cp, lang), char: math.cp };
      else result = any ? this.fallback(primary, any, bold, italic) : primary;
      found = !!any;
      if (!any && !math) this.warnings.add(this.missingGlyph(family, primary.symbolMap, cp, mapped));
    }
    if (found && !primary.symbolMap && (drawn || symbolFallback(cp))) result = { ...result, ...this.symbolLine(family, primary, drawn) };
    if (found && (ownArabicHebrew(family, cp) || ownArabicHebrew(primary.face.familyName, cp))) {
      result = { ...result, ascent: primary.ascent, descent: primary.descent, externalLeading: primary.externalLeading };
      const target = arabicWidth(family, bold) ?? arabicWidth(primary.face.familyName, bold);
      const width = target && !(cp >= 0x590 && cp <= 0x5ff) && !(cp >= 0xfb1d && cp <= 0xfb4f) ? this.arabicSampleWidth(result.face) : 0;
      if (target && width) result.glyphScale = target / width;
    }
    result = this.eastAsianFallbackLine(family, cp, result);
    this.glyphCache.set(key, result);
    return result;
  }

  /** Width (em) of ARABIC_SAMPLE in `face`, or 0 when it lacks those forms. */
  private arabicSampleWidth(face: FontFace): number {
    let w = this.sampleWidths.get(face);
    if (w === undefined) {
      w = 0;
      for (const ch of ARABIC_SAMPLE) {
        const gid = face.font.glyphId(ch.codePointAt(0)!);
        if (gid <= 0) {
          w = 0;
          break;
        }
        w += face.font.advance(gid) / face.font.unitsPerEm;
      }
      this.sampleWidths.set(face, w);
    }
    return w;
  }

  /**
   * The loaded face, outside the fallback families, closest to `family` that has `cp`: of the same
   * kind (serif, sans) where one is, then in the nearest style. A font passed in to cover a script
   * comes in several kinds (Noto Naskh and Noto Sans Arabic).
   */
  private anyFaceWith(family: string, cp: number, bold: boolean, italic: boolean): FontFace | undefined {
    const faces = this.faces.filter((f) => f.font.hasGlyph(cp));
    if (!faces.length) return undefined;
    const generic = this.generic(family);
    const kin = faces.filter((f) => this.generic(f.families[0] ?? '') === generic);
    return this.pickStyle(kin.length ? kin : faces, bold, italic);
  }

  /** A face's half- and full-width forms of proportional glyphs (empty when it has none). */
  private widthForms(face: FontFace): { half: Map<number, number>; full: Map<number, number> } {
    let out = this.widthFormCache.get(face);
    if (!out) {
      const gsub = face.font.tableBytes('GSUB');
      out = gsub ? { half: singleSubstitutions(gsub, 'hwid'), full: singleSubstitutions(gsub, 'fwid') } : { half: new Map(), full: new Map() };
      this.widthFormCache.set(face, out);
    }
    return out;
  }

  /** `face` with Word's line metrics for an East Asian character the requested font lacks (see cjkFallbackLine). */
  private eastAsianFallbackLine(family: string, cp: number, face: ResolvedFace): ResolvedFace {
    const line = cjkFallbackLine(family, cp);
    if (!line) return face;
    let out = this.fallbackLines.get(face);
    if (!out) this.fallbackLines.set(face, (out = { ...face, ascent: line[0], descent: line[1], externalLeading: line[2] }));
    return out;
  }

  /**
   * Line metrics for a symbol (a check box, star or dingbat) that the requested font lacks. Word
   * draws it from Segoe UI Symbol and the line takes that font's height, still with the requested
   * font's external leading above it. A missing East Asian font has these symbols itself (MS
   * Gothic's ☐ is Word's content-control check box), and Word draws a legacy form check box as a
   * shape, so in both the requested font's height stands.
   */
  private symbolLine(family: string, primary: ResolvedFace, drawn: boolean): Pick<ResolvedFace, 'ascent' | 'descent' | 'externalLeading'> {
    if (drawn || cjkLineMetrics(family)) return { ascent: primary.ascent, descent: primary.descent, externalLeading: primary.externalLeading };
    return { ascent: SEGOE_SYMBOL[0], descent: SEGOE_SYMBOL[1], externalLeading: primary.externalLeading };
  }

  /**
   * A fallback face for a character the resolved face lacks. A symbol font's characters keep the
   * symbol font's line metrics whichever font draws them, as Word lays the line out with the
   * original: Segoe UI Symbol drawing a Wingdings arrowhead mustn't make its line taller.
   */
  private fallback(primary: ResolvedFace, face: FontFace, bold: boolean, italic: boolean): ResolvedFace {
    const styled = this.styled(face, bold, italic, primary.symbolMap);
    return primary.symbolMap ? { ...styled, ascent: primary.ascent, descent: primary.descent, externalLeading: primary.externalLeading } : styled;
  }

  /**
   * The character drawn for `cp` in a symbol font that has to be remapped: its Unicode equivalent,
   * or an older look-alike when no loaded font has that one. A code with no Unicode equivalent stays
   * as it is, so it draws as a missing glyph (with a warning) rather than as some other symbol.
   */
  symbolChar(kind: SymbolFontKind, cp: number): number {
    const key = `${kind}\u0000${cp}`;
    let out = this.symbolCache.get(key);
    if (out === undefined) {
      const exact = mapSymbolChar(kind, cp) || cp;
      const alt = symbolLookalike(exact);
      const drawable = (c: number) => this.faces.some((f) => f.font.hasGlyph(c));
      out = alt !== undefined && !drawable(exact) && drawable(alt) ? alt : exact;
      this.symbolCache.set(key, out);
    }
    return out;
  }

  private missingGlyph(family: string, kind: SymbolFontKind | undefined, cp: number, mapped: number): string {
    const hex = (c: number) => 'U+' + c.toString(16).toUpperCase().padStart(4, '0');
    if (!kind) return `No font has a glyph for ${hex(mapped)}`;
    const name = family.trim();
    if (!mapSymbolChar(kind, cp)) return `${name} character ${hex(cp)} has no Unicode equivalent; it needs the ${name} font`;
    return `No font has a glyph for ${name} character ${hex(cp)} (${hex(mapped)}); it needs the ${name} font or one with ${hex(mapped)}`;
  }

  get loadedFaces(): readonly FontFace[] {
    return this.faces;
  }
}
