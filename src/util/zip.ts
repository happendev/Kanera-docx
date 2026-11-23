import type { InflateBytes } from './compression.js';

export interface ArchiveLimits {
  maxPartBytes?: number;
  maxTotalBytes?: number;
}

function byteLimit(value: number | undefined, fallback: number): number {
  if (value === undefined) return fallback;
  if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('ZIP byte limits must be non-negative safe integers');
  return value;
}

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

  private readonly maxPartBytes: number;
  private remainingBytes: number;

  constructor(private readonly data: Uint8Array, private readonly inflateRaw: InflateBytes, limits: ArchiveLimits = {}) {
    this.maxPartBytes = byteLimit(limits.maxPartBytes, 64 * 1024 * 1024);
    this.remainingBytes = byteLimit(limits.maxTotalBytes, 256 * 1024 * 1024);
    let eocd = -1;
    // Bytes after the record and its comment (often a newline added in transit) are ignored, as
    // Office and other ZIP readers ignore them; the directory bounds check rejects a false match.
    for (let i = data.length - 22; i >= Math.max(0, data.length - 22 - 65535); i--) {
      if (u32(data, i) === 0x06054b50 && i + 22 + u16(data, i + 20) <= data.length) {
        const hasLocator = i >= 20 && u32(data, i - 20) === 0x07064b50;
        if (!hasLocator && u32(data, i + 16) + u32(data, i + 12) !== i) continue;
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error('Not a ZIP archive (no end of central directory record)');
    let count = u16(data, eocd + 10);
    let cdOffset = u32(data, eocd + 16);
    let cdSize = u32(data, eocd + 12);
    let directoryLimit = eocd;
    // ZIP64 end of central directory locator
    if (eocd >= 20 && u32(data, eocd - 20) === 0x07064b50) {
      const z64 = u64(data, eocd - 20 + 8);
      if (z64 + 56 > eocd - 20 || u32(data, z64) !== 0x06064b50) throw new Error('Invalid ZIP64 directory');
      const length = u64(data, z64 + 4);
      if (length < 44 || z64 + 12 + length > eocd - 20) throw new Error('Invalid ZIP64 directory length');
      count = u64(data, z64 + 32);
      cdOffset = u64(data, z64 + 48);
      cdSize = u64(data, z64 + 40);
      directoryLimit = z64;
    }
    const cdEnd = cdOffset + cdSize;
    if (cdEnd > directoryLimit || count * 46 > cdSize) throw new Error('Invalid ZIP directory bounds');
    let p = cdOffset;
    for (let i = 0; i < count; i++) {
      if (p + 46 > cdEnd || u32(data, p) !== 0x02014b50) throw new Error('Truncated ZIP directory entry');
      const flags = u16(data, p + 8);
      const method = u16(data, p + 10);
      let compressedSize = u32(data, p + 20);
      let size = u32(data, p + 24);
      const nameLen = u16(data, p + 28);
      const extraLen = u16(data, p + 30);
      const commentLen = u16(data, p + 32);
      let localHeaderOffset = u32(data, p + 42);
      if (p + 46 + nameLen + extraLen + commentLen > cdEnd) throw new Error('Truncated ZIP directory entry');
      const nameBytes = data.subarray(p + 46, p + 46 + nameLen);
      const name = (flags & 0x800 ? textDecoder.decode(nameBytes) : decodeCp437(nameBytes)).replace(/\\/g, '/');
      // ZIP64 extended information
      let e = p + 46 + nameLen;
      const eEnd = e + extraLen;
      while (e + 4 <= eEnd) {
        const id = u16(data, e);
        const len = u16(data, e + 2);
        if (e + 4 + len > eEnd) throw new Error('Truncated ZIP extra field');
        if (id === 0x0001) {
          const fields = [size, compressedSize, localHeaderOffset].filter((n) => n === 0xffffffff);
          if (len < fields.length * 8) throw new Error('Truncated ZIP64 sizes');
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
    const clean = name.replace(/\\/g, '/').replace(/^\/+/, '');
    return this.entries.get(clean) ?? this.lower.get(clean.toLowerCase());
  }

  canonicalName(name: string): string | undefined {
    return this.find(name)?.name;
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
    if (lh + 30 > d.length || u32(d, lh) !== 0x04034b50) throw new Error(`ZIP: bad local header for ${entry.name}`);
    const start = lh + 30 + u16(d, lh + 26) + u16(d, lh + 28);
    if (entry.size > this.maxPartBytes || entry.size > this.remainingBytes) throw new Error('ZIP: decoded byte limit exceeded for ' + entry.name);
    if (start + entry.compressedSize > d.length) throw new Error('ZIP: truncated part ' + entry.name);
    const raw = d.subarray(start, start + entry.compressedSize);
    let out: Uint8Array;
    if (entry.method === 0) out = raw;
    else if (entry.method === 8) out = this.inflateRaw(raw, entry.size);
    else throw new Error(`ZIP: unsupported compression method ${entry.method} for ${entry.name}`);
    if (out.length !== entry.size) throw new Error('ZIP: incorrect decoded size for ' + entry.name);
    this.remainingBytes -= out.length;
    this.cache.set(entry.name, out);
    return out;
  }

  readText(name: string): string | null {
    const bytes = this.read(name);
    if (!bytes) return null;
    return decodeXmlText(bytes);
  }
}

const CP437_HIGH = 'ÇüéâäàåçêëèïîìÄÅÉæÆôöòûùÿÖÜ¢£¥₧ƒáíóúñÑªº¿⌐¬½¼¡«»░▒▓│┤╡╢╖╕╣║╗╝╜╛┐└┴┬├─┼╞╟╚╔╩╦╠═╬╧╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀αßΓπΣσµτΦΘΩδ∞φε∩≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ';

function decodeCp437(bytes: Uint8Array): string {
  let name = '';
  for (const b of bytes) name += b < 128 ? String.fromCharCode(b) : CP437_HIGH[b - 128];
  return name;
}

/** Decode an XML part, honoring a UTF-16 byte-order mark. */
export function decodeXmlText(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return textDecoder.decode(bytes.subarray(3));
  return textDecoder.decode(bytes);
}
