import { ZipArchive } from '../util/zip.js';
import { children, parseXml, type XmlElement } from '../util/xml.js';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/** OPC package access: parts, relationships and XML parsing with caching. */
export class OpcPackage {
  readonly zip: ZipArchive;
  private xmlCache = new Map<string, XmlElement | null>();
  private relsCache = new Map<string, Map<string, Relationship>>();

  constructor(data: Uint8Array) {
    this.zip = new ZipArchive(data);
  }

  xml(part: string): XmlElement | undefined {
    if (this.xmlCache.has(part)) return this.xmlCache.get(part) ?? undefined;
    const text = this.zip.readText(part);
    let el: XmlElement | null = null;
    if (text) {
      try {
        el = parseXml(text);
      } catch {
        el = null;
      }
    }
    this.xmlCache.set(part, el);
    return el ?? undefined;
  }

  bytes(part: string): Uint8Array | null {
    return this.zip.read(part);
  }

  /** Relationships of a part, with targets resolved to absolute part names. */
  rels(part: string): Map<string, Relationship> {
    let map = this.relsCache.get(part);
    if (map) return map;
    map = new Map();
    const slash = part.lastIndexOf('/');
    const dir = slash >= 0 ? part.slice(0, slash) : '';
    const name = slash >= 0 ? part.slice(slash + 1) : part;
    const relsPart = `${dir ? dir + '/' : ''}_rels/${name}.rels`;
    const root = this.xml(relsPart);
    for (const r of children(root, 'pr:Relationship').concat(children(root, 'Relationship'))) {
      const external = r.attrs.TargetMode === 'External';
      const target = r.attrs.Target ?? '';
      map.set(r.attrs.Id ?? '', {
        id: r.attrs.Id ?? '',
        type: r.attrs.Type ?? '',
        target: external ? target : resolvePartName(dir, target),
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
