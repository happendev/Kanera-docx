import type { InflateBytes } from '../util/compression.js';
import { ZipArchive, type ArchiveLimits } from '../util/zip.js';
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

  constructor(data: Uint8Array, inflateRaw: InflateBytes, limits?: ArchiveLimits) {
    this.zip = new ZipArchive(data, inflateRaw, limits);
  }

  private partName(part: string): string {
    return this.zip.canonicalName(part) ?? part;
  }

  xml(part: string): XmlElement | undefined {
    part = this.partName(part);
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

  contentType(part: string): string | undefined {
    const root = this.xml('[Content_Types].xml');
    part = this.partName(part);
    const ext = part.split('.').pop()?.toLowerCase();
    // The XML reader preserves unprefixed names in the OPC content-types namespace.
    const declarations = children(root);
    const local = (name: string) => name.split(':').pop();
    const override = declarations.find((e) => local(e.name) === 'Override' && e.attrs.PartName?.startsWith('/') && this.partName(e.attrs.PartName.slice(1)) === part);
    const fallback = declarations.find((e) => local(e.name) === 'Default' && e.attrs.Extension?.toLowerCase() === ext);
    return (override ?? fallback)?.attrs.ContentType?.split(';')[0].trim().toLowerCase();
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
    const root = this.xml(relsPart);
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
