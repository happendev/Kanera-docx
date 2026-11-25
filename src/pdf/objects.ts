import type { Compression } from '../util/compression.js';

/** Minimal PDF file assembler: numbered objects, streams, cross-reference table. */
export class PdfObjects {
  private bodies: Array<string | { head: string; body: Uint8Array } | null> = [];

  private readonly pending = new Set<Promise<void>>();
  private failure?: { cause: unknown };

  constructor(private readonly compression: Compression) {}

  alloc(): number {
    this.bodies.push(null);
    return this.bodies.length;
  }

  set(num: number, content: string): void {
    this.bodies[num - 1] = content;
  }

  add(content: string): number {
    const n = this.alloc();
    this.set(n, content);
    return n;
  }

  stream(num: number, dict: string, data: Uint8Array, compress = true): void {
    const filter = compress ? ' /Filter /FlateDecode' : '';
    const prefix = `<<${dict ? ' ' + dict : ''}${filter} /Length `;
    const stream = { head: '', body: data };
    this.bodies[num - 1] = stream;
    const store = (body: Uint8Array) => {
      stream.head = prefix + body.length + ' >>\nstream\n';
      stream.body = body;
    };
    // Let native compression overlap the following objects. Bound both the number
    // of outstanding jobs and the uncompressed bytes they can keep alive.
    if (compress && this.compression.deflateZlibExactAsync && data.length >= 512 && data.length <= 512 * 1024 && this.pending.size < 8) {
      const job = this.compression.deflateZlibExactAsync(data)
        .then(store)
        .catch(cause => { this.failure = { cause }; })
        .finally(() => { this.pending.delete(job); });
      this.pending.add(job);
    } else store(compress ? this.compression.deflateZlib(data) : data);
  }

  addStream(dict: string, data: Uint8Array, compress = true): number {
    const n = this.alloc();
    this.stream(n, dict, data, compress);
    return n;
  }

  /** Serialise all objects. `root` is the catalog, `info` the document information dictionary. */
  async finish(root: number, info?: number): Promise<Uint8Array<ArrayBuffer>> {
    if (this.pending.size) await Promise.all(this.pending);
    if (this.failure) throw this.failure.cause;
    const parts: Array<string | Uint8Array> = [];
    let offset = 0;
    const push = (part: string | Uint8Array) => {
      parts.push(part);
      offset += part.length;
    };
    push('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n');
    const offsets: number[] = [];
    this.bodies.forEach((body, i) => {
      offsets.push(offset);
      push(`${i + 1} 0 obj\n`);
      if (body !== null && typeof body !== 'string') {
        push(body.head);
        push(body.body);
        push('\nendstream');
      } else push(body ?? 'null');
      push('\nendobj\n');
    });
    const xref = offset;
    let table = `xref\n0 ${this.bodies.length + 1}\n0000000000 65535 f \n`;
    for (const o of offsets) table += `${String(o).padStart(10, '0')} 00000 n \n`;
    table += `trailer\n<< /Size ${this.bodies.length + 1} /Root ${root} 0 R${info ? ` /Info ${info} 0 R` : ''} >>\nstartxref\n${xref}\n%%EOF\n`;
    push(table);
    const out = new Uint8Array(offset);
    let at = 0;
    for (const part of parts) {
      if (typeof part === 'string') {
        for (let i = 0; i < part.length; i++) out[at++] = part.charCodeAt(i);
      } else {
        out.set(part, at);
        at += part.length;
      }
    }
    return out;
  }
}

export function latin1(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

export function concat(parts: Uint8Array[]): Uint8Array<ArrayBuffer> {
  let n = 0;
  for (const p of parts) n += p.length;
  const out = new Uint8Array(n);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

// Glyph spacing and coordinates often repeat. Keep one formatted value per slot.
const roundedNumbers = new Float64Array(4096).fill(NaN);
const numberStrings: string[] = [];

/** A number in PDF syntax (no exponent, trimmed). */
export function num(v: number): string {
  if (!Number.isFinite(v)) return '0';
  const rounded = Math.round(v * 1000);
  const slot = rounded & (roundedNumbers.length - 1);
  if (roundedNumbers[slot] !== rounded) {
    roundedNumbers[slot] = rounded;
    numberStrings[slot] = String(rounded / 1000);
  }
  return numberStrings[slot];
}

/** Text string: PDFDocEncoding-safe ASCII as a literal, anything else as UTF-16BE hex. */
export function textString(s: string): string {
  if (/^[\x20-\x7e]*$/.test(s)) return `(${s.replace(/[\\()]/g, (c) => '\\' + c)})`;
  let hex = 'FEFF';
  for (let i = 0; i < s.length; i++) hex += s.charCodeAt(i).toString(16).padStart(4, '0');
  return `<${hex.toUpperCase()}>`;
}

/** Name object with #xx escapes for delimiters and non-ASCII. */
export function pdfName(s: string): string {
  let out = '/';
  for (const b of new TextEncoder().encode(s)) {
    if (b < 0x21 || b > 0x7e || '#()<>[]{}/%'.includes(String.fromCharCode(b))) out += '#' + b.toString(16).padStart(2, '0');
    else out += String.fromCharCode(b);
  }
  return out;
}
