import { child, children, type XmlElement } from '../util/xml.js';
import {
  combineStyleRuns,
  mergePara,
  mergeRun,
  mergeTbl,
  mergeTc,
  mergeTr,
  parseParaDelta,
  parseRunDelta,
  parseTblDelta,
  parseTcDelta,
  parseTrDelta,
  type ParaDelta,
  type PropContext,
  type RunDelta,
  type TblDelta,
  type TcDelta,
  type TrDelta,
} from './props.js';

export interface CondFormat {
  pPr: ParaDelta;
  rPr: RunDelta;
  tbl: TblDelta;
  tr: TrDelta;
  tc: TcDelta;
}

export interface StyleDef {
  id: string;
  type: string;
  name: string;
  basedOn?: string;
  isDefault: boolean;
  pPr: ParaDelta;
  rPr: RunDelta;
  tbl: TblDelta;
  tr: TrDelta;
  tc: TcDelta;
  cond: Map<string, CondFormat>;
}

export interface TableStyle {
  pPr: ParaDelta;
  rPr: RunDelta;
  tbl: TblDelta;
  tr: TrDelta;
  tc: TcDelta;
  cond: Map<string, CondFormat>;
}

export class StyleSheet {
  private readonly styles = new Map<string, { id: string; type: string; name: string; element: XmlElement; parsed?: StyleDef }>();
  defaultsRun: RunDelta = {};
  defaultsPara: ParaDelta = {};
  defaultParagraphStyle?: string;
  defaultCharacterStyle?: string;
  defaultTableStyle?: string;
  private paraCache = new Map<string, ParaDelta>();
  private paraRunCache = new Map<string, RunDelta>();
  private charCache = new Map<string, RunDelta>();
  private tableCache = new Map<string, TableStyle>();

  constructor(root: XmlElement | undefined, private readonly ctx: PropContext) {
    if (!root) return;
    const dd = child(root, 'w:docDefaults');
    this.defaultsRun = parseRunDelta(child(child(dd, 'w:rPrDefault'), 'w:rPr'), ctx);
    const paraDefault = child(dd, 'w:pPrDefault');
    this.defaultsPara = parseParaDelta(child(paraDefault, 'w:pPr'), ctx);
    for (const s of children(root, 'w:style')) {
      const id = s.attrs['w:styleId'] ?? '';
      const type = s.attrs['w:type'] ?? 'paragraph';
      const isDefault = s.attrs['w:default'] === '1' || s.attrs['w:default'] === 'true';
      const name = child(s, 'w:name')?.attrs['w:val'] ?? id;
      this.styles.set(id, { id, type, name, element: s });
      if (isDefault && type === 'paragraph' && !this.defaultParagraphStyle) this.defaultParagraphStyle = id;
      if (isDefault && type === 'character' && !this.defaultCharacterStyle) this.defaultCharacterStyle = id;
      if (isDefault && type === 'table' && !this.defaultTableStyle) this.defaultTableStyle = id;
    }
    // Without a style marked as the default, Word still applies the paragraph style named Normal.
    if (!this.defaultParagraphStyle) {
      const normal = [...this.styles.values()].find((d) => d.type === 'paragraph' && (d.id === 'Normal' || d.name.toLowerCase() === 'normal'));
      if (normal) this.defaultParagraphStyle = normal.id;
    }
    // Word supplies Normal spacing for a named default paragraph style when the
    // document specifies run defaults alone, or no defaults at all. Style and direct overrides still win.
    if (this.defaultParagraphStyle && !paraDefault) this.defaultsPara = { line: 1.15, spaceAfter: 8 };
  }

  /** Style names by style id. */
  names(): Map<string, string> {
    return new Map([...this.styles.values()].map((s) => [s.id, s.name]));
  }

  /** The style with this name: built-in styles keep their English names whatever their ids. */
  byName(name: string): StyleDef | undefined {
    const n = name.toLowerCase();
    for (const d of this.styles.values()) if (d.name.toLowerCase() === n) return this.get(d.id);
    return undefined;
  }

  get(id: string | undefined): StyleDef | undefined {
    const entry = id === undefined ? undefined : this.styles.get(id);
    if (!entry || entry.parsed) return entry?.parsed;
    const { element: s, name, type } = entry;
    const ctx = this.ctx;
    const def: StyleDef = {
      id: entry.id,
      type,
      name,
      basedOn: child(s, 'w:basedOn')?.attrs['w:val'],
      isDefault: s.attrs['w:default'] === '1' || s.attrs['w:default'] === 'true',
      pPr: parseParaDelta(child(s, 'w:pPr'), ctx),
      rPr: parseRunDelta(child(s, 'w:rPr'), ctx),
      tbl: parseTblDelta(child(s, 'w:tblPr'), ctx),
      tr: parseTrDelta(child(s, 'w:trPr')),
      tc: parseTcDelta(child(s, 'w:tcPr'), ctx),
      cond: new Map(),
    };
    // A style's own pStyle/rStyle references are meaningless; drop them.
    delete def.pPr.styleId;
    delete def.rPr.styleId;
    for (const c of children(s, 'w:tblStylePr')) {
      def.cond.set(c.attrs['w:type'] ?? '', {
        pPr: parseParaDelta(child(c, 'w:pPr'), ctx),
        rPr: parseRunDelta(child(c, 'w:rPr'), ctx),
        tbl: parseTblDelta(child(c, 'w:tblPr'), ctx),
        tr: parseTrDelta(child(c, 'w:trPr')),
        tc: parseTcDelta(child(c, 'w:tcPr'), ctx),
      });
    }
    entry.parsed = def;
    return def;
  }

  /** Style chain from the root ancestor down to `id` (cycles and missing parents are cut). */
  chain(id: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let cur = this.get(id);
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = this.get(cur.basedOn);
    }
    return out;
  }

  /** Paragraph style properties (resolved through basedOn), not including doc defaults. */
  paraStyle(id: string | undefined): ParaDelta {
    const key = id ?? '';
    let d = this.paraCache.get(key);
    if (!d) {
      d = {};
      for (const s of this.chain(id)) d = mergePara(d, s.pPr);
      this.paraCache.set(key, d);
    }
    return d;
  }

  /** Run properties contributed by a paragraph style. */
  paraStyleRun(id: string | undefined): RunDelta {
    const key = id ?? '';
    let d = this.paraRunCache.get(key);
    if (!d) {
      d = {};
      for (const s of this.chain(id)) d = mergeRun(d, s.rPr);
      this.paraRunCache.set(key, d);
    }
    return d;
  }

  /** Run properties of a character style. */
  charStyle(id: string | undefined): RunDelta {
    const key = id ?? '';
    let d = this.charCache.get(key);
    if (!d) {
      d = {};
      for (const s of this.chain(id)) if (s.type === 'character' || s.type === 'paragraph') d = mergeRun(d, s.rPr);
      this.charCache.set(key, d);
    }
    return d;
  }

  tableStyle(id: string | undefined): TableStyle {
    const key = id ?? '';
    let t = this.tableCache.get(key);
    if (!t) {
      t = { pPr: {}, rPr: {}, tbl: {}, tr: {}, tc: {}, cond: new Map() };
      for (const s of this.chain(id)) {
        t.pPr = mergePara(t.pPr, s.pPr);
        t.rPr = mergeRun(t.rPr, s.rPr);
        t.tbl = mergeTbl(t.tbl, s.tbl);
        t.tr = mergeTr(t.tr, s.tr);
        t.tc = mergeTc(t.tc, s.tc);
        for (const [type, c] of s.cond) {
          const prev = t.cond.get(type);
          t.cond.set(
            type,
            prev
              ? {
                  pPr: mergePara(prev.pPr, c.pPr),
                  rPr: mergeRun(prev.rPr, c.rPr),
                  tbl: mergeTbl(prev.tbl, c.tbl),
                  tr: mergeTr(prev.tr, c.tr),
                  tc: mergeTc(prev.tc, c.tc),
                }
              : c,
          );
        }
      }
      this.tableCache.set(key, t);
    }
    return t;
  }

  /**
   * Full run properties for text in a paragraph: defaults, table style, paragraph style and
   * character style (toggle properties XOR across style levels), then direct formatting.
   */
  runDelta(opts: { paraStyle?: string; charStyle?: string; tableRun?: RunDelta; direct?: RunDelta; tableOverridesSize?: boolean }): RunDelta {
    const styleLevels: RunDelta[] = [];
    if (opts.tableRun) styleLevels.push(opts.tableRun);
    styleLevels.push(this.paraStyleRun(opts.paraStyle));
    if (opts.charStyle) styleLevels.push(this.charStyle(opts.charStyle));
    let styles = combineStyleRuns(styleLevels);
    // Word 2013+: a table style's font size beats the default paragraph style's.
    if (opts.tableOverridesSize && opts.tableRun?.size !== undefined) {
      styles = { ...styles, size: opts.tableRun.size };
      if (opts.tableRun.sizeCs !== undefined) styles.sizeCs = opts.tableRun.sizeCs;
    }
    let out = mergeRun(this.defaultsRun, styles);
    if (opts.direct) {
      const direct = { ...opts.direct };
      delete direct.styleId;
      out = mergeRun(out, direct);
    }
    return out;
  }
}
