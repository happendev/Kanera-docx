import type { InflateBytes } from '../util/compression.js';
import { ZipArchive, type ArchiveLimits } from '../util/zip.js';
import { children, parseXml, XmlDepthError, type XmlElement } from '../util/xml.js';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/** Retain a recurring template and the latest variant. Compare the complete XML before sharing a tree. */
export class XmlTemplate {
  static readonly theme = new XmlTemplate('a:theme');
  private recurring?: { text: string; root: XmlElement };
  private recent?: { text: string; root: XmlElement };

  constructor(private readonly rootName: string, private readonly maxLength = 64 * 1024) {}

  parse(text: string): XmlElement {
    if (text.length > this.maxLength) return parseXml(text);
    if (this.recurring?.text === text) return this.recurring.root;
    if (this.recent?.text === text) {
      const hit = this.recent;
      this.recent = this.recurring;
      this.recurring = hit;
      return hit.root;
    }
    const root = parseXml(text);
    if (root.name === this.rootName) this.recent = { text, root };
    return root;
  }
}

const packageRels = new XmlTemplate('pr:Relationships', 8192);
const partRels = new XmlTemplate('pr:Relationships', 8192);

/** OPC package access: parts, relationships and XML parsing with caching. */
export class OpcPackage {
  readonly zip: ZipArchive;
  private xmlCache = new Map<string, XmlElement | null>();
  private relsCache = new Map<string, Map<string, Relationship>>();

  private contentTypes?: { overrides: Map<string, string | undefined>; defaults: Map<string, string | undefined> };

  constructor(data: Uint8Array, inflateRaw: InflateBytes, limits?: ArchiveLimits) {
    this.zip = new ZipArchive(data, inflateRaw, limits);
  }

  private partName(part: string): string {
    return this.zip.canonicalName(part) ?? part;
  }

  xml(part: string, template?: XmlTemplate): XmlElement | undefined {
    part = this.partName(part);
    if (this.xmlCache.has(part)) return this.xmlCache.get(part) ?? undefined;
    const text = this.zip.readText(part);
    let el: XmlElement | null = null;
    if (text) {
      try {
        el = template ? template.parse(text) : parseXml(text);
      } catch (err) {
        if (err instanceof XmlDepthError) throw err;
        el = null;
      }
    }
    this.xmlCache.set(part, el);
    return el ?? undefined;
  }

  contentType(part: string): string | undefined {
    if (!this.contentTypes) {
      const overrides = new Map<string, string | undefined>();
      const defaults = new Map<string, string | undefined>();
      for (const el of children(this.xml('[Content_Types].xml'))) {
        const local = el.name.split(':').pop();
        const type = el.attrs.ContentType?.split(';')[0].trim().toLowerCase();
        if (local === 'Override' && el.attrs.PartName?.startsWith('/')) {
          const key = this.partName(el.attrs.PartName.slice(1));
          if (!overrides.has(key)) overrides.set(key, type);
        } else if (local === 'Default' && el.attrs.Extension !== undefined) {
          const ext = el.attrs.Extension.toLowerCase();
          if (!defaults.has(ext)) defaults.set(ext, type);
        }
      }
      this.contentTypes = { overrides, defaults };
    }
    const { overrides, defaults } = this.contentTypes;
    part = this.partName(part);
    return overrides.has(part) ? overrides.get(part) : defaults.get(part.split('.').pop()!.toLowerCase());
  }

  bytes(part: string): Uint8Array | null {
    return this.zip.read(part);
  }

  /** Relationships of a part, with targets resolved to absolute part names. */
  rels(part: string): Map<string, Relationship> {
    part = this.partName(part);
    let map = this.relsCache.get(part);
    if (map) return map;
    map = new Map();
    const slash = part.lastIndexOf('/');
    const dir = slash >= 0 ? part.slice(0, slash) : '';
    const name = slash >= 0 ? part.slice(slash + 1) : part;
    const relsPart = `${dir ? dir + '/' : ''}_rels/${name}.rels`;
    const root = this.xml(relsPart, part ? partRels : packageRels);
    for (const r of children(root, 'pr:Relationship').concat(children(root, 'Relationship'))) {
      const external = r.attrs.TargetMode === 'External';
      const target = r.attrs.Target ?? '';
      map.set(r.attrs.Id ?? '', {
        id: r.attrs.Id ?? '',
        type: r.attrs.Type ?? '',
        target: external ? target : this.partName(resolvePartName(dir, target)),
        external,
      });
    }
    this.relsCache.set(part, map);
    return map;
  }

  relByType(part: string, typeSuffix: string): Relationship | undefined {
    for (const r of this.rels(part).values()) if (r.type.endsWith(typeSuffix)) return r;
    return undefined;
  }

  mainDocumentPart(): string {
    const rel = this.relByType('', '/officeDocument');
    if (rel) return rel.target;
    if (this.zip.has('word/document.xml')) return 'word/document.xml';
    throw new Error('Not a Word document: no main document part found');
  }
}

export function resolvePartName(baseDir: string, target: string): string {
  let t = target.replace(/\\/g, '/');
  if (t.startsWith('/')) return normalizeParts(t.slice(1).split('/'));
  const parts = baseDir ? baseDir.split('/') : [];
  return normalizeParts([...parts, ...t.split('/')]);
}

function normalizeParts(parts: string[]): string {
  const out: string[] = [];
  for (const p of parts) {
    if (p === '' || p === '.') continue;
    if (p === '..') out.pop();
    else out.push(decodeURIComponentSafe(p));
  }
  return out.join('/');
}

function decodeURIComponentSafe(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}
