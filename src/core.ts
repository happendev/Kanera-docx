import type { DocumentInfo } from './docx/model.js';
import { parseDocx } from './docx/parser.js';
import { FontRegistry, type FontSource } from './fonts/registry.js';
import { layoutDocument, type LayoutStats } from './layout/engine.js';
import { renderPdf } from './pdf/render.js';

export type { FontSource } from './fonts/registry.js';
export type { LayoutStats } from './layout/engine.js';

export interface ConvertOptions {
  /** Fonts to use, in addition to fonts embedded in the document. They win over bundled fonts. */
  fonts?: FontSource[];
  /** Use fonts embedded in the .docx (default true). */
  embeddedFonts?: boolean;
  /** Override document information written to the PDF. */
  info?: Partial<DocumentInfo>;
  /** Compress PDF content streams (default true). */
  compress?: boolean;
}

export interface ConvertResult {
  pdf: Uint8Array;
  pageCount: number;
  /** Things that could not be reproduced exactly (missing fonts, unsupported images, ...). */
  warnings: string[];
  stats: LayoutStats;
}

/** Fonts registered by an entry point (bundled or system fonts), below user and embedded fonts. */
export interface FontProvider {
  register(registry: FontRegistry): void | Promise<void>;
}

function toBytes(input: Uint8Array | ArrayBuffer): Uint8Array {
  return input instanceof Uint8Array ? input : new Uint8Array(input);
}

/** Convert a .docx file to PDF. */
export async function convertDocx(input: Uint8Array | ArrayBuffer, options: ConvertOptions = {}, providers: FontProvider[] = []): Promise<ConvertResult> {
  const doc = parseDocx(toBytes(input));
  const fonts = new FontRegistry();
  for (const f of options.fonts ?? []) fonts.add(f, 'user');
  if (options.embeddedFonts !== false) {
    for (const e of doc.embeddedFonts) {
      try {
        fonts.add({ data: e.data, family: e.family, bold: e.bold, italic: e.italic }, 'embedded');
      } catch {
        doc.warnings.push(`Embedded font ${e.family} could not be read`);
      }
    }
  }
  for (const p of providers) await p.register(fonts);
  for (const [family, hints] of doc.fontHints) fonts.setHints(family, hints);
  await fonts.prepare(doc.fontFamilies, doc.codePoints);
  const layout = layoutDocument(doc, fonts);
  const warnings = [...doc.warnings, ...layout.warnings];
  const pdf = renderPdf(layout.pages, { info: { ...doc.info, ...options.info }, compress: options.compress }, warnings);
  warnings.push(...fonts.warnings);
  return { pdf, pageCount: layout.pages.length, warnings: [...new Set(warnings)], stats: layout.stats };
}
