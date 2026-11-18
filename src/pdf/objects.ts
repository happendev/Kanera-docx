import { deflateZlib } from '../util/deflate.js';

/** Minimal PDF file assembler: numbered objects, streams, cross-reference table. */
export class PdfObjects {
  private bodies: Array<Uint8Array | null> = [];

  alloc(): number {
    this.bodies.push(null);
    return this.bodies.length;
  }

  set(num: number, content: string): void {
    this.bodies[num - 1] = latin1(content);
  }

  add(content: string): number {
    const n = this.alloc();
    this.set(n, content);
    return n;
  }

  stream(num: number, dict: string, data: Uint8Array, compress = true): void {
    const body = compress ? deflateZlib(data) : data;
    const filter = compress ? ' /Filter /FlateDecode' : '';
    const head = latin1(`<<${dict ? ' ' + dict : ''}${filter} /Length ${body.length} >>\nstream\n`);
    this.bodies[num - 1] = concat([head, body, latin1('\nendstream')]);
  }

  addStream(dict: string, data: Uint8Array, compress = true): number {
    const n = this.alloc();
    this.stream(n, dict, data, compress);
    return n;
  }

  /** Serialise all objects. `root` is the catalog, `info` the document information dictionary. */
  finish(root: number, info?: number): Uint8Array {
    const parts: Uint8Array[] = [];
    let offset = 0;
    const push = (b: Uint8Array) => {
      parts.push(b);
      offset += b.length;
    };
    push(latin1('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n'));
    const offsets: number[] = [];
    this.bodies.forEach((body, i) => {
      offsets.push(offset);
      push(latin1(`${i + 1} 0 obj\n`));
      push(body ?? latin1('null'));
      push(latin1('\nendobj\n'));
    });
    const xref = offset;
    let table = `xref\n0 ${this.bodies.length + 1}\n0000000000 65535 f \n`;
    for (const o of offsets) table += `${String(o).padStart(10, '0')} 00000 n \n`;
    table += `trailer\n<< /Size ${this.bodies.length + 1} /Root ${root} 0 R${info ? ` /Info ${info} 0 R` : ''} >>\nstartxref\n${xref}\n%%EOF\n`;
    push(latin1(table));
    return concat(parts);
  }
}

export function latin1(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}

export function concat(parts: Uint8Array[]): Uint8Array {
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

/** A number in PDF syntax (no exponent, trimmed). */
export function num(v: number): string {
  if (!Number.isFinite(v)) return '0';
  const r = Math.round(v * 1000) / 1000;
  if (Object.is(r, -0) || r === 0) return '0';
  return String(r);
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
