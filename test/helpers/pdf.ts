import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { inflateSync } from 'node:zlib';

function hasGhostscript(): boolean {
  try {
    execFileSync('gs', ['--version'], { stdio: 'ignore' });
    return true;
  } catch {
    return false;
  }
}

/** Ghostscript is used as an independent PDF reader when it is installed. */
export const ghostscript = hasGhostscript();

/**
 * Text of each page as Ghostscript extracts it (through our ToUnicode maps), so a test can check
 * that the PDF is readable by a third-party parser and says what the document says.
 */
export function pdfPagesText(pdf: Uint8Array): string[] {
  const dir = mkdtempSync(join(tmpdir(), 'docx2pdf-test-'));
  try {
    const file = join(dir, 'in.pdf');
    writeFileSync(file, pdf);
    const pages: string[] = [];
    for (let i = 1; i <= pdfPageCount(pdf); i++) {
      pages.push(
        execFileSync('gs', ['-q', '-dNOPAUSE', '-dBATCH', '-dSAFER', '-sDEVICE=txtwrite', `-dFirstPage=${i}`, `-dLastPage=${i}`, '-sOutputFile=-', file], {
          encoding: 'utf8',
          // txtwrite prints a harmless device-refcount message on some gs versions.
          stdio: ['ignore', 'pipe', 'ignore'],
        }),
      );
    }
    return pages;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** Render page 1 with Ghostscript; throws if Ghostscript reports an error. */
export function ghostscriptRenders(pdf: Uint8Array): void {
  const dir = mkdtempSync(join(tmpdir(), 'docx2pdf-test-'));
  try {
    const file = join(dir, 'in.pdf');
    writeFileSync(file, pdf);
    const out = execFileSync('gs', ['-dNOPAUSE', '-dBATCH', '-dSAFER', '-sDEVICE=png16m', '-r20', `-sOutputFile=${join(dir, 'p%d.png')}`, file], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    if (/error|warning/i.test(out)) throw new Error(out);
    readFileSync(join(dir, 'p1.png'));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

export function pdfText(pdf: Uint8Array): string {
  let s = '';
  for (const b of pdf) s += String.fromCharCode(b);
  return s;
}

export function pdfPageCount(pdf: Uint8Array): number {
  return (pdfText(pdf).match(/\/Type \/Page\b/g) ?? []).length;
}

/** Decompressed content of every FlateDecode stream, for checking drawing operators. */
export function pdfStreams(pdf: Uint8Array): string[] {
  const src = pdfText(pdf);
  const out: string[] = [];
  const re = /<<([^]*?)>>\s*stream\n/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) {
    const len = /\/Length (\d+)/.exec(m[1]);
    if (!len) continue;
    const start = m.index + m[0].length;
    const data = pdf.subarray(start, start + Number(len[1]));
    const body = m[1].includes('/FlateDecode') ? inflateSync(data) : data;
    let s = '';
    for (const b of body) s += String.fromCharCode(b);
    out.push(s);
  }
  return out;
}
