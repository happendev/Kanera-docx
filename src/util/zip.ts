import { inflateRaw } from './inflate.js';

export interface ZipEntry {
  name: string;
  method: number;
  compressedSize: number;
  size: number;
  localHeaderOffset: number;
}

const textDecoder = new TextDecoder('utf-8');

function u16(d: Uint8Array, o: number): number {
  return d[o] | (d[o + 1] << 8);
}
function u32(d: Uint8Array, o: number): number {
  return (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0;
}
function u64(d: Uint8Array, o: number): number {
  return u32(d, o) + u32(d, o + 4) * 0x100000000;
}

/** Read-only ZIP archive (stored + deflated entries, ZIP64 sizes). */
export class ZipArchive {
  readonly entries = new Map<string, ZipEntry>();
  /** Lower-cased name lookup: OPC part names are case-insensitive. */
  private readonly lower = new Map<string, ZipEntry>();
  private readonly cache = new Map<string, Uint8Array>();

  constructor(private readonly data: Uint8Array) {
    let eocd = -1;
    for (let i = data.length - 22; i >= Math.max(0, data.length - 22 - 65535); i--) {
      if (u32(data, i) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error('Not a ZIP archive (no end of central directory record)');
    let count = u16(data, eocd + 10);
    let cdOffset = u32(data, eocd + 16);
    // ZIP64 end of central directory locator
    if (eocd >= 20 && u32(data, eocd - 20) === 0x07064b50) {
      const z64 = u64(data, eocd - 20 + 8);
      if (u32(data, z64) === 0x06064b50) {
        count = u64(data, z64 + 32);
        cdOffset = u64(data, z64 + 48);
      }
    }
    let p = cdOffset;
    for (let i = 0; i < count; i++) {
      if (u32(data, p) !== 0x02014b50) break;
      const flags = u16(data, p + 8);
      const method = u16(data, p + 10);
      let compressedSize = u32(data, p + 20);
      let size = u32(data, p + 24);
      const nameLen = u16(data, p + 28);
      const extraLen = u16(data, p + 30);
      const commentLen = u16(data, p + 32);
      let localHeaderOffset = u32(data, p + 42);
      const nameBytes = data.subarray(p + 46, p + 46 + nameLen);
      const name = flags & 0x800 ? textDecoder.decode(nameBytes) : decodeCp437OrUtf8(nameBytes);
      // ZIP64 extended information
      let e = p + 46 + nameLen;
      const eEnd = e + extraLen;
      while (e + 4 <= eEnd) {
        const id = u16(data, e);
        const len = u16(data, e + 2);
        if (id === 0x0001) {
          let q = e + 4;
          if (size === 0xffffffff) {
            size = u64(data, q);
            q += 8;
          }
          if (compressedSize === 0xffffffff) {
            compressedSize = u64(data, q);
            q += 8;
          }
          if (localHeaderOffset === 0xffffffff) {
            localHeaderOffset = u64(data, q);
          }
        }
        e += 4 + len;
      }
      const entry: ZipEntry = { name, method, compressedSize, size, localHeaderOffset };
      this.entries.set(name, entry);
      this.lower.set(name.toLowerCase(), entry);
      p += 46 + nameLen + extraLen + commentLen;
    }
  }

  private find(name: string): ZipEntry | undefined {
    const clean = name.replace(/^\/+/, '');
    return this.entries.get(clean) ?? this.lower.get(clean.toLowerCase());
  }

  has(name: string): boolean {
    return this.find(name) !== undefined;
  }

  read(name: string): Uint8Array | null {
    const entry = this.find(name);
    if (!entry) return null;
    const cached = this.cache.get(entry.name);
    if (cached) return cached;
    const d = this.data;
    const lh = entry.localHeaderOffset;
    if (u32(d, lh) !== 0x04034b50) throw new Error(`ZIP: bad local header for ${entry.name}`);
    const start = lh + 30 + u16(d, lh + 26) + u16(d, lh + 28);
    const raw = d.subarray(start, start + entry.compressedSize);
    let out: Uint8Array;
    if (entry.method === 0) out = raw;
    else if (entry.method === 8) out = inflateRaw(raw, entry.size);
    else throw new Error(`ZIP: unsupported compression method ${entry.method} for ${entry.name}`);
    this.cache.set(entry.name, out);
    return out;
  }

  readText(name: string): string | null {
    const bytes = this.read(name);
    if (!bytes) return null;
    return decodeXmlText(bytes);
  }
}

function decodeCp437OrUtf8(bytes: Uint8Array): string {
  // OOXML part names are ASCII in practice; UTF-8 decoding is a safe superset for that case.
  return textDecoder.decode(bytes);
}

/** Decode an XML part, honoring a UTF-16 byte-order mark. */
export function decodeXmlText(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return textDecoder.decode(bytes.subarray(3));
  return textDecoder.decode(bytes);
}
