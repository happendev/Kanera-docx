import type { DocumentInfo } from './docx/model.js';
import { parseDocx } from './docx/parser.js';
import { FontRegistry, type FontSource } from './fonts/registry.js';
import { layoutDocument, type LayoutStats } from './layout/engine.js';
import type { Compression } from './util/compression.js';
import { renderPdf } from './pdf/render.js';

/** Word for Mac renders through Quartz, which places text on a 0.24pt (1/300 inch) grid. */
const MAC_PAINT_GRID = 0.24;

export type { FontSource } from './fonts/registry.js';
export type { LayoutStats } from './layout/engine.js';

export interface ConvertOptions {
  /** Maximum uncompressed bytes per ZIP part and across all parts read. */
  archiveLimits?: import('./util/zip.js').ArchiveLimits;
  /** Maximum pixels per raster image and across distinct images rendered. */
  imageLimits?: import('./image/budget.js').ImageLimits;
  /** JPEG quality from 1 to 100. Omit to select JPEGs by quantization table and encode them at quality 80, or use false to disable lossy recompression. */
  imageQuality?: number | false;
  /** Additional URI schemes to enable for PDF links, such as ['tel']. */
  additionalLinkSchemes?: readonly string[];
  /** Fonts to use, in addition to fonts embedded in the document. They win over bundled fonts. */
  fonts?: FontSource[];
  /** Use fonts embedded in the .docx (default true). */
  embeddedFonts?: boolean;
  /** Override document information written to the PDF. */
  info?: Partial<DocumentInfo>;
  /** Compress PDF content streams (default true). */
  compress?: boolean;
  /**
   * Which Word's PDF output to match (default 'windows'). Word for Mac draws text at sizes, baselines
   * and line starts snapped to a 0.24pt grid; layout is the same on both.
   */
  wordPlatform?: 'windows' | 'mac';
}

export interface ConvertResult {
  pdf: Uint8Array<ArrayBuffer>;
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
export async function convertWithCompression(input: Uint8Array | ArrayBuffer, options: ConvertOptions, providers: FontProvider[], compression: Compression): Promise<ConvertResult> {
  const doc = parseDocx(toBytes(input), compression.inflateRaw, options.archiveLimits, options.embeddedFonts);
  const fonts = new FontRegistry();
  for (const f of options.fonts ?? []) fonts.add(f, 'user');
  for (const e of doc.embeddedFonts) fonts.add({ data: e.data, family: e.family, bold: e.bold, italic: e.italic }, 'embedded');
  for (const p of providers) await p.register(fonts);
  for (const [family, hints] of doc.fontHints) fonts.setHints(family, hints);
  await fonts.prepare(doc.fontRequests);
  const layout = layoutDocument(doc, fonts);
  const warnings = [...doc.warnings, ...layout.warnings];
  const pdf = await renderPdf(layout.pages, compression, {
    info: { ...doc.info, ...options.info },
    compress: options.compress,
    imageLimits: options.imageLimits,
    imageQuality: options.imageQuality,
    additionalLinkSchemes: options.additionalLinkSchemes,
    paintGrid: options.wordPlatform === 'mac' ? MAC_PAINT_GRID : undefined,
  }, warnings);
  warnings.push(...fonts.warnings);
  return { pdf, pageCount: layout.pages.length, warnings: [...new Set(warnings)], stats: layout.stats };
}
