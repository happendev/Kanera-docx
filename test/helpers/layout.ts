import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseDocx } from '../../src/docx/parser.js';
import { BUNDLED_FONTS } from '../../src/fonts/bundled.js';
import { FontRegistry } from '../../src/fonts/registry.js';
import { layoutDocument, type LayoutResult } from '../../src/layout/engine.js';
import type { DrawOp, GlyphOp, Page } from '../../src/layout/types.js';

const FONT_DIR = fileURLToPath(new URL('../../fonts/', import.meta.url));

/** Lay out a .docx with only the bundled fonts, as the converter does by default. */
export async function layout(docx: Uint8Array): Promise<LayoutResult> {
  const doc = parseDocx(docx);
  const fonts = new FontRegistry();
  for (const b of BUNDLED_FONTS) fonts.add({ family: b.family, bold: b.bold, italic: b.italic, weight: b.weight, covers: b.covers, load: async () => readFileSync(FONT_DIR + b.file) }, 'bundled');
  for (const [family, hints] of doc.fontHints) fonts.setHints(family, hints);
  await fonts.prepare(doc.fontRequests);
  return layoutDocument(doc, fonts);
}

export function glyphOps(ops: DrawOp[]): GlyphOp[] {
  return ops.filter((o): o is GlyphOp => o.kind === 'glyphs');
}

export interface TextLine {
  y: number;
  x: number;
  text: string;
  ops: GlyphOp[];
}

/** Body text of a page grouped into lines by baseline, top to bottom. */
export function lines(ops: DrawOp[]): TextLine[] {
  const byY = new Map<number, GlyphOp[]>();
  for (const op of glyphOps(ops)) {
    const key = Math.round(op.y * 100) / 100;
    byY.set(key, [...(byY.get(key) ?? []), op]);
  }
  return [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([y, list]) => {
      list.sort((a, b) => a.x - b.x);
      return { y, x: list[0].x, text: list.map((o) => o.texts.join('')).join(''), ops: list };
    });
}

export function pageText(page: Page): string {
  return lines(page.body)
    .map((l) => l.text)
    .join('\n');
}
