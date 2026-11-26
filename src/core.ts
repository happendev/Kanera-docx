import { inputBytes, type InputOptions } from './util/input.js';
import type { DocumentInfo } from './ooxml/model.js';
import { parseDocx } from './docx/parser.js';
import { FontRegistry, type FontSource } from './fonts/registry.js';
import { layoutDocument, type LayoutStats } from './layout/engine.js';
import type { Compression } from './util/compression.js';
import { renderPdf } from './pdf/render.js';
import type { ImageFormat, PageImage } from './raster/render.js';

export { extractTextWithCompression, type ExtractTextOptions, type ExtractTextResult } from './text/index.js';

/** Word for Mac renders through Quartz, which places text on a 0.24pt (1/300 inch) grid. */
const MAC_PAINT_GRID = 0.24;

export type { FontSource } from './fonts/registry.js';
export type { LayoutStats } from './layout/engine.js';
export type { ImageFormat, PageImage } from './raster/render.js';

export interface ConvertOptions extends InputOptions {
  /** Maximum uncompressed bytes per ZIP part and across all parts read. */
  archiveLimits?: import('./util/zip.js').ArchiveLimits;
  /** Maximum pixels per raster image and across distinct images rendered. */
  imageLimits?: import('./image/budget.js').ImageLimits;
  /** JPEG quality from 1 to 100. Omit to select JPEGs up to 4 megapixels by quantization table and encode them at quality 80, or use false to disable lossy recompression. */
  imageQuality?: number | false;
  /** Additional URI schemes to enable for PDF links, such as ['tel']. */
  additionalLinkSchemes?: readonly string[];
  /** Fonts to use, in addition to fonts embedded in the document. They win over bundled fonts. */
  fonts?: FontSource[];
  /** Use fonts embedded in the .docx (default true). */
  useEmbeddedFonts?: boolean;
  /** Override document information written to the PDF. */
  info?: Partial<DocumentInfo>;
  /** Compress PDF content streams (default true). */
  compress?: boolean;
  /**
   * Stops the conversion once aborted, rejecting with the signal's reason. It is checked between
   * stages, font loads and pages, so a long layout finishes before it takes effect.
   */
  signal?: AbortSignal;
  /**
   * Which Word's PDF output to match (default 'windows'). Word for Mac draws text at sizes, baselines
   * and line starts snapped to a 0.24pt grid, and has macOS's fonts: on Windows a document's macOS
   * fonts give way to the alternates its font table names for them, or to the Windows fonts Word
   * matches them with (Avenir Next to Calibri).
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

export interface ImageOptions extends Omit<ConvertOptions, 'info' | 'compress' | 'imageQuality' | 'additionalLinkSchemes'> {
  /** 'png' (default) or 'jpeg'. */
  format?: ImageFormat;
  /** Resolution in dots per inch (default 150), unless `width` or `height` is set. */
  dpi?: number;
  /**
   * The image size in pixels, in place of `dpi`, from 1 to 65535. With both, each page fits within
   * them, keeping its shape unless `keepAspectRatio` is false; with one, the other follows the
   * page's shape.
   */
  width?: number;
  height?: number;
  /** Keep each page's shape when both `width` and `height` are set (default true); false stretches it to exactly that size. */
  keepAspectRatio?: boolean;
  /** JPEG quality from 1 to 100 (default 90). */
  quality?: number;
  /** Page numbers to render, from 1 (default: every page). */
  pages?: readonly number[];
  /**
   * Receives each image as soon as it is encoded, in page order, with the page count, so only the
   * pages in flight stay in memory: the result's images then have empty `data`. Drawing waits for
   * a returned promise, so writing a file here holds back the next page.
   */
  onPage?: (image: PageImage, pageCount: number) => void | Promise<void>;
}

export interface ImagesResult {
  /** One image per rendered page, in the order requested. */
  images: PageImage[];
  /** Pages in the whole document. */
  pageCount: number;
  warnings: string[];
  stats: LayoutStats;
}

/** Fonts registered by an entry point (bundled or system fonts), below user and embedded fonts. */
export interface FontProvider {
  register(registry: FontRegistry): void | Promise<void>;
}

export async function prepareDocxWithCompression(input: Uint8Array | ArrayBuffer, options: ConvertOptions, providers: FontProvider[], compression: Compression) {
  const { signal } = options;
  signal?.throwIfAborted();
  const doc = parseDocx(inputBytes(input, options.maxInputBytes), compression.inflateRaw, options.archiveLimits, options.useEmbeddedFonts);
  const fonts = new FontRegistry(doc.hasMath);
  fonts.macFonts = options.wordPlatform === 'mac';
  for (const f of options.fonts ?? []) fonts.add(f, 'user');
  for (const e of doc.embeddedFonts) fonts.add({ data: e.data, family: e.family, bold: e.bold, italic: e.italic }, 'embedded');
  for (const p of providers) await p.register(fonts);
  for (const [family, hints] of doc.fontHints) fonts.setHints(family, hints);
  await fonts.prepare(doc.fontRequests, signal);
  signal?.throwIfAborted();
  return { doc, fonts };
}

async function layOut(input: Uint8Array | ArrayBuffer, options: ConvertOptions, providers: FontProvider[], compression: Compression) {
  const { doc, fonts } = await prepareDocxWithCompression(input, options, providers, compression);
  options.signal?.throwIfAborted();
  const layout = layoutDocument(doc, fonts);
  options.signal?.throwIfAborted();
  const warnings = [...doc.warnings, ...layout.warnings];
  const paintGrid = options.wordPlatform === 'mac' ? MAC_PAINT_GRID : undefined;
  return { doc, fonts, layout, warnings, paintGrid };
}

/** Convert a .docx file to PDF. */
export async function convertWithCompression(input: Uint8Array | ArrayBuffer, options: ConvertOptions, providers: FontProvider[], compression: Compression): Promise<ConvertResult> {
  const { doc, fonts, layout, warnings, paintGrid } = await layOut(input, options, providers, compression);
  const pdf = await renderPdf(layout.pages, compression, {
    info: { ...doc.info, ...options.info },
    compress: options.compress,
    imageLimits: options.imageLimits,
    imageQuality: options.imageQuality,
    additionalLinkSchemes: options.additionalLinkSchemes,
    signal: options.signal,
    paintGrid,
  }, warnings);
  warnings.push(...fonts.warnings);
  return { pdf, pageCount: layout.pages.length, warnings: [...new Set(warnings)], stats: layout.stats };
}

/** Render a .docx file's pages as PNG or JPEG images. */
export async function convertToImagesWithCompression(input: Uint8Array | ArrayBuffer, options: ImageOptions, providers: FontProvider[], compression: Compression): Promise<ImagesResult> {
  // Load the rasteriser only for image output, and check its options before any work.
  const { checkRasterOptions, renderImages } = await import('./raster/render.js');
  checkRasterOptions(options);
  const { fonts, layout, warnings, paintGrid } = await layOut(input, options, providers, compression);
  const pageCount = layout.pages.length;
  const pages = options.pages?.map((n) => {
    if (!Number.isInteger(n) || n < 1 || n > pageCount) throw new RangeError(`Page ${n} is out of range (1-${pageCount})`);
    return n - 1;
  });
  const images = await renderImages(layout.pages, compression, { format: options.format, dpi: options.dpi, width: options.width, height: options.height, keepAspectRatio: options.keepAspectRatio, quality: options.quality, pages, imageLimits: options.imageLimits, onPage: options.onPage, signal: options.signal, paintGrid }, warnings);
  warnings.push(...fonts.warnings);
  return { images, pageCount, warnings: [...new Set(warnings)], stats: layout.stats };
}
