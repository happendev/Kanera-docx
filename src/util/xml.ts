/**
 * Minimal namespace-aware XML parser for OOXML parts.
 *
 * Element and attribute names are normalized to canonical prefixes (`w:p`, `a:blip`, ...) by
 * namespace URI, so documents that use unusual prefixes, default namespaces or Strict OOXML URIs
 * read the same as Transitional documents written by Word.
 */

export interface XmlElement {
  name: string;
  namespaceURI?: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  parent?: XmlElement;
}
export type XmlNode = XmlElement | string;

const NS: Record<string, string> = {
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main': 'w',
  'http://purl.oclc.org/ooxml/wordprocessingml/main': 'w',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing': 'wp',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing': 'wp',
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://purl.oclc.org/ooxml/drawingml/main': 'a',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://purl.oclc.org/ooxml/drawingml/picture': 'pic',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://purl.oclc.org/ooxml/drawingml/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://purl.oclc.org/ooxml/drawingml/diagram': 'dgm',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://purl.oclc.org/ooxml/officeDocument/math': 'm',
  'urn:schemas-microsoft-com:vml': 'v',
  'urn:schemas-microsoft-com:office:office': 'o',
  'urn:schemas-microsoft-com:office:word': 'w10',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape': 'wps',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup': 'wpg',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas': 'wpc',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingInk': 'wpi',
  'http://schemas.microsoft.com/office/word/2010/wordml': 'w14',
  'http://schemas.microsoft.com/office/word/2012/wordml': 'w15',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing': 'wp14',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14',
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main': 'asvg',
  'http://www.w3.org/2000/svg': 'svg',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'pr',
  'http://schemas.openxmlformats.org/package/2006/content-types': 'ct',
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties': 'cp',
  'http://purl.org/dc/elements/1.1/': 'dc',
  'http://purl.org/dc/terms/': 'dcterms',
  'http://www.w3.org/XML/1998/namespace': 'xml',
  'http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas': 'lc',
  'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
};

export class XmlDepthError extends Error {}

/** Namespaces whose mc:Choice branches we can render (others fall back to mc:Fallback). */
export const SUPPORTED_CHOICE_PREFIXES = new Set(['wps', 'wpg', 'wp14', 'w14', 'w15', 'a14', 'asvg', 'wpc', 'w', 'a', 'wp', 'pic', 'r', 'm', 'v', 'o']);

const ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, ent: string) => {
    if (ent[0] === '#') {
      const cp = ent[1] === 'x' || ent[1] === 'X' ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
      if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff) return m;
      return String.fromCodePoint(cp);
    }
    return Object.hasOwn(ENTITIES, ent) ? ENTITIES[ent] : m;
  });
}

/** Elements whose whitespace-only text content is significant. */
function keepsWhitespace(name: string): boolean {
  const local = name.slice(name.indexOf(':') + 1);
  return local === 't' || local === 'instrText' || local === 'delText' || local === 'delInstrText';
}

// Scratch space for a start tag's attributes (parsing is synchronous, so one pair serves all calls).
const attrNames: string[] = [];
const attrValues: string[] = [];

// Attribute maps inherit nothing, as with Object.create(null), but V8 keeps them as compact
// fast-mode objects; an object from Object.create(null) is a hash table from the start.
function NoPrototype(): void {}
NoPrototype.prototype = Object.create(null);
const AttrMap = NoPrototype as unknown as new () => Record<string, string>;

export function parseXml(src: string): XmlElement {
  return new XmlReader(src).parse();
}

/** Canonical prefix for a document prefix in this scope; null when unbound or not an OOXML namespace. */
function canonicalPrefix(prefix: string, scope: Record<string, string>): string | null {
  const uri = scope[prefix];
  return uri !== undefined && Object.hasOwn(NS, uri) ? NS[uri] : null;
}

function resolve(qname: string, scope: Record<string, string>, isAttr: boolean): string {
  const colon = qname.indexOf(':');
  if (colon < 0) {
    if (isAttr) return qname;
    const canon = canonicalPrefix('', scope);
    return canon ? `${canon}:${qname}` : qname;
  }
  const canon = canonicalPrefix(qname.slice(0, colon), scope);
  if (canon === null) return qname;
  // Most documents already use the canonical prefixes: keep the name as it is.
  if (canon.length === colon && qname.startsWith(canon)) return qname;
  return canon + qname.slice(colon);
}

/** mc:Choice's Requires lists document prefixes; rewrite to canonical ones so callers can test support. */
function canonicalRequires(requires: string, scope: Record<string, string>): string {
  return requires
    .split(/\s+/)
    .filter(Boolean)
    .map((p) => {
      const uri = scope[p];
      return uri === undefined ? p : Object.hasOwn(NS, uri) ? NS[uri] : '#unknown';
    })
    .join(' ');
}

// Resolved names by name as written, per namespace scope. Parts repeat a few dozen names
// thousands of times, so each is resolved once and every element shares the resulting string.
interface ScopeNames {
  elements: Map<string, { name: string; namespaceURI: string | undefined }>;
  attrs: Map<string, string>;
}

/**
 * One parse, with each kind of markup read by its own small method. V8 optimises a function that
 * has run a few thousand times on a background thread, and the memory that takes grows with the
 * function: as one large function, the parser added more than 10 MiB to a conversion's peak memory.
 */
class XmlReader {
  private readonly root: XmlElement = { name: '#document', attrs: {}, children: [] };
  private readonly stack: XmlElement[] = [this.root];
  private readonly nsStack: Array<Record<string, string>> = [Object.assign(Object.create(null), { xml: 'http://www.w3.org/XML/1998/namespace' })];
  private readonly scopeNames = new Map<Record<string, string>, ScopeNames>();
  private namesScope: Record<string, string> | undefined;
  private names: ScopeNames = { elements: new Map(), attrs: new Map() };

  constructor(private readonly src: string) {}

  parse(): XmlElement {
    const src = this.src;
    const n = src.length;
    let i = 0;
    while (i < n) {
      const lt = src.indexOf('<', i);
      if (lt < 0) {
        appendText(this.stack[this.stack.length - 1], src.slice(i));
        break;
      }
      if (lt > i) appendText(this.stack[this.stack.length - 1], src.slice(i, lt));
      const c1 = src.charCodeAt(lt + 1);
      i = c1 === 47 /* / */ ? this.endTag(lt) : c1 === 33 /* ! */ || c1 === 63 /* ? */ ? this.markup(lt) : this.startTag(lt);
    }
    const docEl = this.root.children.find((c): c is XmlElement => typeof c !== 'string');
    if (!docEl) throw new Error('XML: no document element');
    return docEl;
  }

  /** A comment, CDATA section, processing instruction or declaration. Returns the index after it. */
  private markup(i: number): number {
    const src = this.src;
    const n = src.length;
    if (src.charCodeAt(i + 1) === 63 /* ? */) {
      const end = src.indexOf('?>', i + 2);
      return end < 0 ? n : end + 2;
    }
    if (src.startsWith('<!--', i)) {
      const end = src.indexOf('-->', i + 4);
      return end < 0 ? n : end + 3;
    }
    if (src.startsWith('<![CDATA[', i)) {
      const end = src.indexOf(']]>', i + 9);
      this.stack[this.stack.length - 1].children.push(src.slice(i + 9, end < 0 ? n : end));
      return end < 0 ? n : end + 3;
    }
    // DOCTYPE or other declaration; skip including a possible internal subset.
    let depth = 0;
    let j = i + 2;
    for (; j < n; j++) {
      const ch = src[j];
      if (ch === '[') depth++;
      else if (ch === ']') depth--;
      else if (ch === '>' && depth <= 0) break;
    }
    return j + 1;
  }

  private endTag(i: number): number {
    const end = this.src.indexOf('>', i + 2);
    if (this.stack.length > 1) {
      this.nsStack.pop();
      finishElement(this.stack.pop()!);
    }
    return end < 0 ? this.src.length : end + 1;
  }

  /** Read a start tag's name and attributes, then add its element. Returns the index after the tag. */
  private startTag(i: number): number {
    if (this.stack.length > 128) throw new XmlDepthError('XML nesting limit exceeded');
    const src = this.src;
    const n = src.length;
    // Character tests are written out in the loops: they run for every character of every name.
    let j = i + 1;
    for (; j < n; j++) {
      const c = src.charCodeAt(j);
      if (c === 32 || c === 9 || c === 10 || c === 13 || c === 62 /* > */ || c === 47 /* / */) break;
    }
    const rawName = src.slice(i + 1, j);
    // Attribute names and values, gathered before namespace declarations among them are applied.
    let count = 0;
    let declares = false;
    let selfClose = false;
    for (;;) {
      for (let c = src.charCodeAt(j); c === 32 || c === 9 || c === 10 || c === 13; c = src.charCodeAt(++j));
      const ch = src.charCodeAt(j);
      if (ch === 62 /* > */) {
        j++;
        break;
      }
      if (ch === 47 /* / */) {
        selfClose = true;
        j = src.indexOf('>', j) + 1;
        if (j === 0) j = n;
        break;
      }
      if (j >= n) break;
      const ns = j;
      for (; j < n; j++) {
        const c = src.charCodeAt(j);
        if (c === 32 || c === 9 || c === 10 || c === 13 || c === 62 || c === 47 || c === 61 /* = */) break;
      }
      const an = src.slice(ns, j);
      for (let c = src.charCodeAt(j); c === 32 || c === 9 || c === 10 || c === 13; c = src.charCodeAt(++j));
      if (an.charCodeAt(0) === 120 /* x */ && (an === 'xmlns' || an.startsWith('xmlns:'))) declares = true;
      if (src.charCodeAt(j) !== 61 /* = */) {
        attrNames[count] = an;
        attrValues[count++] = '';
        continue;
      }
      j++;
      for (let c = src.charCodeAt(j); c === 32 || c === 9 || c === 10 || c === 13; c = src.charCodeAt(++j));
      const q = src[j];
      let value: string;
      if (q === '"' || q === "'") {
        const end = src.indexOf(q, j + 1);
        value = src.slice(j + 1, end < 0 ? n : end);
        j = end < 0 ? n : end + 1;
      } else {
        const vs = j;
        for (; j < n; j++) {
          const c = src.charCodeAt(j);
          if (c === 32 || c === 9 || c === 10 || c === 13 || c === 62) break;
        }
        value = src.slice(vs, j);
      }
      attrNames[count] = an;
      attrValues[count++] = value.indexOf('&') < 0 ? value : decodeEntities(value);
    }
    this.open(rawName, count, declares, selfClose);
    return j;
  }

  /** Add the element whose name and attributes `startTag` read. */
  private open(rawName: string, count: number, declares: boolean, selfClose: boolean): void {
    const parentScope = this.nsStack[this.nsStack.length - 1];
    let scope = parentScope;
    if (declares) {
      for (let k = 0; k < count; k++) {
        const an = attrNames[k];
        if (an === 'xmlns' || an.startsWith('xmlns:')) {
          if (scope === parentScope) scope = Object.create(parentScope) as Record<string, string>;
          scope[an === 'xmlns' ? '' : an.slice(6)] = attrValues[k];
        }
      }
    }
    if (scope !== this.namesScope) {
      let found = this.scopeNames.get(scope);
      if (!found) this.scopeNames.set(scope, (found = { elements: new Map(), attrs: new Map() }));
      this.names = found;
      this.namesScope = scope;
    }
    const names = this.names;
    const attrs = new AttrMap();
    for (let k = 0; k < count; k++) {
      const an = attrNames[k];
      if (declares && (an === 'xmlns' || an.startsWith('xmlns:'))) continue;
      let key = names.attrs.get(an);
      if (key === undefined) names.attrs.set(an, (key = resolve(an, scope, true)));
      attrs[key] = attrValues[k];
    }
    let resolved = names.elements.get(rawName);
    if (!resolved) {
      const prefixEnd = rawName.indexOf(':');
      resolved = { name: resolve(rawName, scope, false), namespaceURI: scope[prefixEnd < 0 ? '' : rawName.slice(0, prefixEnd)] };
      names.elements.set(rawName, resolved);
    }
    const { name, namespaceURI } = resolved;
    if (name === 'mc:Choice' && attrs.Requires) attrs.Requires = canonicalRequires(attrs.Requires, scope);
    const el: XmlElement = { name, namespaceURI, attrs, children: [] };
    this.stack[this.stack.length - 1].children.push(el);
    if (selfClose) {
      finishElement(el);
    } else {
      this.stack.push(el);
      this.nsStack.push(scope);
    }
  }
}

function appendText(el: XmlElement, raw: string): void {
  if (!raw) return;
  el.children.push(decodeEntities(raw));
}

/** Drop insignificant whitespace-only text nodes and merge adjacent text. */
function finishElement(el: XmlElement): void {
  const kids = el.children;
  if (kids.length === 0) return;
  const keep = keepsWhitespace(el.name) || el.attrs['xml:space'] === 'preserve';
  let hasElement = false;
  for (const k of kids) if (typeof k !== 'string') hasElement = true;
  if (!hasElement) {
    if (kids.length > 1) el.children = [kids.join('')];
    if (!keep && typeof el.children[0] === 'string' && isBlank(el.children[0] as string)) el.children = [];
    return;
  }
  if (keep) return;
  let blank = false;
  for (const k of kids) if (typeof k === 'string' && isBlank(k)) blank = true;
  if (blank) el.children = kids.filter((k) => typeof k !== 'string' || !isBlank(k));
}

/** Whitespace only (as /^\s*$/ tests it, for the characters XML allows). */
function isBlank(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (!(c === 32 || c === 9 || c === 10 || c === 13) && !/\s/.test(s[i])) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------------------------
// Query helpers

export function child(el: XmlElement | undefined | null, name: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) if (typeof c !== 'string' && c.name === name) return c;
  return undefined;
}

export function children(el: XmlElement | undefined | null, name?: string): XmlElement[] {
  if (!el) return [];
  const out: XmlElement[] = [];
  for (const c of el.children) if (typeof c !== 'string' && (name === undefined || c.name === name)) out.push(c);
  return out;
}

export function elements(el: XmlElement | undefined | null): XmlElement[] {
  return children(el);
}

/** Depth-first search for the first descendant with the given name. */
export function descendant(el: XmlElement | undefined | null, name: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const c of el.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) return c;
    const d = descendant(c, name);
    if (d) return d;
  }
  return undefined;
}

export function descendants(el: XmlElement | undefined | null, name: string, out: XmlElement[] = []): XmlElement[] {
  if (!el) return out;
  for (const c of el.children) {
    if (typeof c === 'string') continue;
    if (c.name === name) out.push(c);
    descendants(c, name, out);
  }
  return out;
}

export function attr(el: XmlElement | undefined | null, name: string): string | undefined {
  return el ? el.attrs[name] : undefined;
}

export function textContent(el: XmlElement | undefined | null): string {
  if (!el) return '';
  let s = '';
  for (const c of el.children) s += typeof c === 'string' ? c : textContent(c);
  return s;
}

/** Path lookup: `path(el, 'a:graphic', 'a:graphicData', 'pic:pic')`. */
export function path(el: XmlElement | undefined | null, ...names: string[]): XmlElement | undefined {
  let cur: XmlElement | undefined = el ?? undefined;
  for (const n of names) {
    cur = child(cur, n);
    if (!cur) return undefined;
  }
  return cur;
}

/**
 * Replace mc:AlternateContent with the children of the branch we can render: the first
 * mc:Choice whose required namespaces are all supported, otherwise mc:Fallback.
 */
export function resolveAlternateContent(el: XmlElement): XmlElement[] {
  for (const c of children(el, 'mc:Choice')) {
    const req = (c.attrs.Requires ?? '').split(/\s+/).filter(Boolean);
    if (req.every((p) => SUPPORTED_CHOICE_PREFIXES.has(p))) return children(c);
  }
  const fb = child(el, 'mc:Fallback');
  return fb ? children(fb) : [];
}
