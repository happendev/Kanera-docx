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

export function parseXml(src: string): XmlElement {
  const root: XmlElement = { name: '#document', attrs: {}, children: [] };
  const stack: XmlElement[] = [root];
  const nsStack: Array<Record<string, string>> = [Object.assign(Object.create(null), { xml: 'http://www.w3.org/XML/1998/namespace' })];
  let i = 0;
  const n = src.length;

  const resolve = (qname: string, scope: Record<string, string>, isAttr: boolean): string => {
    const colon = qname.indexOf(':');
    if (colon < 0) {
      if (isAttr) return qname;
      const uri = scope[''];
      if (uri === undefined) return qname;
      const canon = Object.hasOwn(NS, uri) ? NS[uri] : undefined;
      return canon ? `${canon}:${qname}` : qname;
    }
    const prefix = qname.slice(0, colon);
    const local = qname.slice(colon + 1);
    const uri = scope[prefix];
    if (uri === undefined) return qname;
    const canon = Object.hasOwn(NS, uri) ? NS[uri] : undefined;
    return canon ? `${canon}:${local}` : qname;
  };

  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) {
      appendText(stack[stack.length - 1], src.slice(i));
      break;
    }
    if (lt > i) appendText(stack[stack.length - 1], src.slice(i, lt));
    i = lt;
    const c1 = src.charCodeAt(i + 1);
    if (c1 === 33 /* ! */) {
      if (src.startsWith('<!--', i)) {
        const end = src.indexOf('-->', i + 4);
        i = end < 0 ? n : end + 3;
      } else if (src.startsWith('<![CDATA[', i)) {
        const end = src.indexOf(']]>', i + 9);
        const text = src.slice(i + 9, end < 0 ? n : end);
        const parent = stack[stack.length - 1];
        parent.children.push(text);
        i = end < 0 ? n : end + 3;
      } else {
        // DOCTYPE or other declaration; skip including a possible internal subset.
        let depth = 0;
        let j = i + 2;
        for (; j < n; j++) {
          const ch = src[j];
          if (ch === '[') depth++;
          else if (ch === ']') depth--;
          else if (ch === '>' && depth <= 0) break;
        }
        i = j + 1;
      }
      continue;
    }
    if (c1 === 63 /* ? */) {
      const end = src.indexOf('?>', i + 2);
      i = end < 0 ? n : end + 2;
      continue;
    }
    if (c1 === 47 /* / */) {
      const end = src.indexOf('>', i + 2);
      i = end < 0 ? n : end + 1;
      if (stack.length > 1) {
        const el = stack.pop()!;
        nsStack.pop();
        finishElement(el);
      }
      continue;
    }
    if (stack.length > 128) throw new XmlDepthError('XML nesting limit exceeded');
    // Start tag
    let j = i + 1;
    while (j < n && !isNameEnd(src.charCodeAt(j))) j++;
    const rawName = src.slice(i + 1, j);
    const rawAttrs: Array<[string, string]> = [];
    let selfClose = false;
    for (;;) {
      while (j < n && isSpace(src.charCodeAt(j))) j++;
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
      while (j < n && !isNameEnd(src.charCodeAt(j)) && src.charCodeAt(j) !== 61) j++;
      const an = src.slice(ns, j);
      while (j < n && isSpace(src.charCodeAt(j))) j++;
      if (src.charCodeAt(j) !== 61 /* = */) {
        rawAttrs.push([an, '']);
        continue;
      }
      j++;
      while (j < n && isSpace(src.charCodeAt(j))) j++;
      const q = src[j];
      if (q === '"' || q === "'") {
        const end = src.indexOf(q, j + 1);
        rawAttrs.push([an, decodeEntities(src.slice(j + 1, end < 0 ? n : end))]);
        j = end < 0 ? n : end + 1;
      } else {
        const vs = j;
        while (j < n && !isSpace(src.charCodeAt(j)) && src.charCodeAt(j) !== 62) j++;
        rawAttrs.push([an, decodeEntities(src.slice(vs, j))]);
      }
    }
    i = j;

    const parentScope = nsStack[nsStack.length - 1];
    let scope = parentScope;
    for (const [an, av] of rawAttrs) {
      if (an === 'xmlns' || an.startsWith('xmlns:')) {
        if (scope === parentScope) scope = Object.create(parentScope);
        scope[an === 'xmlns' ? '' : an.slice(6)] = av;
      }
    }
    const attrs: Record<string, string> = Object.create(null);
    for (const [an, av] of rawAttrs) {
      if (an === 'xmlns' || an.startsWith('xmlns:')) continue;
      attrs[resolve(an, scope, true)] = av;
    }
    const name = resolve(rawName, scope, false);
    if (name === 'mc:Choice' && attrs.Requires) {
      // Requires lists document prefixes; rewrite to canonical ones so callers can test support.
      attrs.Requires = attrs.Requires.split(/\s+/)
        .filter(Boolean)
        .map((p) => {
          const uri = scope[p];
          return uri === undefined ? p : Object.hasOwn(NS, uri) ? NS[uri] : '#unknown';
        })
        .join(' ');
    }
    const parent = stack[stack.length - 1];
    const prefixEnd = rawName.indexOf(':');
    const namespaceURI = scope[prefixEnd < 0 ? '' : rawName.slice(0, prefixEnd)];
    const el: XmlElement = { name, namespaceURI, attrs, children: [], parent };
    parent.children.push(el);
    if (!selfClose) {
      stack.push(el);
      nsStack.push(scope);
    } else {
      finishElement(el);
    }
  }
  const docEl = root.children.find((c): c is XmlElement => typeof c !== 'string');
  if (!docEl) throw new Error('XML: no document element');
  docEl.parent = undefined;
  return docEl;
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
    if (!keep && typeof el.children[0] === 'string' && /^\s*$/.test(el.children[0] as string)) el.children = [];
    return;
  }
  el.children = kids.filter((k) => typeof k !== 'string' || keep || !/^\s*$/.test(k));
}

function isSpace(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13;
}
function isNameEnd(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13 || c === 62 || c === 47 || Number.isNaN(c);
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
