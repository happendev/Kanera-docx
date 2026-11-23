// Raster images decoded to RGBA pixels, for drawing page images: the same formats and colour
// adjustments as the PDF writer embeds, but decoded rather than passed through.

import type { Compression } from '../util/compression.js';
import { ImageBudget } from './budget.js';
import { decodeAdam7, decodeDib, decodeGif, le32s, readPng, sniffImage, unfilter, writePixel, type PngInfo } from './image.js';
import { decodeJpegPixels } from './jpeg-decoder.js';
import { decodeTiff } from './tiff.js';

/** Straight (not premultiplied) RGBA pixels. */
export interface Pixmap {
  width: number;
  height: number;
  data: Uint8Array;
  /** Whether every pixel's alpha is 255. */
  opaque: boolean;
}

/** Each colour component c (0–1) drawn as slope·c + offset, clipped to 0–1. */
export interface ColorAdjust {
  slope: number;
  offset: number;
}

/**
 * Decode an image file (PNG, JPEG, GIF, BMP or TIFF), or a DIB's header and `dibBits`, to RGBA.
 * Null when the format isn't supported; throws on malformed data or when the image exceeds the
 * pixel budget.
 */
export async function decodePixels(data: Uint8Array, compression: Compression, budget: ImageBudget, dibBits?: Uint8Array, adjust?: ColorAdjust): Promise<Pixmap | null> {
  let px: Pixmap | null = null;
  let adjustable = true;
  if (dibBits) px = fromRgb(decodeDib(data, dibBits, budget));
  else {
    switch (sniffImage(data)) {
      case 'png':
        px = decodePng(readPng(data, budget), compression);
        break;
      case 'jpeg': {
        const jpeg = await decodeJpeg(data, budget);
        px = jpeg?.pixels ?? null;
        // The PDF writer leaves CMYK JPEGs' colours as they are.
        adjustable = jpeg?.components !== 4;
        break;
      }
      case 'gif':
        px = fromRgba(decodeGif(data, budget));
        break;
      case 'bmp': {
        const offset = le32s(data, 10);
        if (offset < 14 || offset > data.length) throw new Error('BMP: invalid pixel offset');
        px = fromRgb(decodeDib(data.subarray(14, offset), data.subarray(offset), budget));
        break;
      }
      case 'tiff':
        px = fromRgba(decodeTiff(data, compression.inflateZlib, budget));
        break;
    }
  }
  if (px && adjust && adjustable) applyAdjust(px, adjust);
  return px;
}

function fromRgba(img: { width: number; height: number; data: Uint8Array } | null): Pixmap | null {
  if (!img) return null;
  let opaque = true;
  for (let i = 3; i < img.data.length; i += 4) {
    if (img.data[i] !== 255) {
      opaque = false;
      break;
    }
  }
  return { width: img.width, height: img.height, data: img.data, opaque };
}

function fromRgb(img: { width: number; height: number; rgb: Uint8Array; alpha?: Uint8Array } | null): Pixmap | null {
  if (!img) return null;
  const n = img.width * img.height;
  const data = new Uint8Array(n * 4);
  const { rgb, alpha } = img;
  for (let i = 0, j = 0, k = 0; i < n; i++, j += 4, k += 3) {
    data[j] = rgb[k];
    data[j + 1] = rgb[k + 1];
    data[j + 2] = rgb[k + 2];
    data[j + 3] = alpha ? alpha[i] : 255;
  }
  return { width: img.width, height: img.height, data, opaque: !alpha };
}

/** Whether this machine stores the low byte of a word first, so RGBA pixels write as one word. */
const LITTLE_ENDIAN = new Uint8Array(new Uint32Array([1]).buffer)[0] === 1;

/**
 * RGBA words for every sample value of a palette or grey PNG of up to 8 bits, as `writePixel`
 * makes them, and whether any of them is not opaque.
 */
function pixelTable(info: PngInfo): { rgba: Uint32Array; clear: boolean } {
  const { bitDepth, colorType, palette, trns } = info;
  const rgba = new Uint32Array(1 << bitDepth);
  const line = new Uint8Array(1);
  const px = new Uint8Array(4);
  const view = new Uint32Array(px.buffer);
  let clear = false;
  for (let v = 0; v < rgba.length; v++) {
    // The value as the line's first sample.
    line[0] = v << (8 - bitDepth);
    px.fill(0);
    writePixel(px, 0, line, 0, bitDepth, colorType, palette, trns);
    rgba[v] = view[0];
    if (px[3] !== 255) clear = true;
  }
  return { rgba, clear };
}

function applyAdjust(px: Pixmap, { slope, offset }: ColorAdjust): void {
  const table = new Uint8Array(256);
  for (let v = 0; v < 256; v++) table[v] = Math.round(Math.min(1, Math.max(0, slope * (v / 255) + offset)) * 255);
  const d = px.data;
  for (let i = 0; i < d.length; i += 4) {
    d[i] = table[d[i]];
    d[i + 1] = table[d[i + 1]];
    d[i + 2] = table[d[i + 2]];
  }
}

function decodePng(info: PngInfo, compression: Compression): Pixmap {
  const { width, height, bitDepth, colorType, palette, trns } = info;
  if (info.interlace) return fromRgba({ width, height, data: decodeAdam7(info, compression) })!;
  const channels = colorType === 2 ? 3 : colorType === 0 ? 1 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4;
  const bitsPP = channels * bitDepth;
  const bpp = Math.max(1, bitsPP >> 3);
  const stride = Math.ceil((width * bitsPP) / 8);
  const raw = compression.inflateZlib(info.idat, height * (stride + 1));
  if (raw.length !== height * (stride + 1)) throw new Error('PNG: truncated scanlines');
  const out = new Uint8Array(width * height * 4);
  const words = LITTLE_ENDIAN ? new Uint32Array(out.buffer) : null;
  // Palette and grey pixels (up to 8 bits) through a table of whole RGBA words.
  const table = words && bitDepth <= 8 && (colorType === 3 || colorType === 0) ? pixelTable(info) : null;
  let prev: Uint8Array = new Uint8Array(stride);
  let opaque = true;
  for (let y = 0; y < height; y++) {
    const at = y * (stride + 1);
    const line = raw.subarray(at + 1, at + 1 + stride);
    unfilter(raw[at], line, prev, bpp);
    prev = line;
    let o = y * width * 4;
    if (table && words) {
      const w0 = y * width;
      if (bitDepth === 8) for (let x = 0; x < width; x++) words[w0 + x] = table.rgba[line[x]];
      else {
        const perByte = 8 / bitDepth;
        const mask = (1 << bitDepth) - 1;
        for (let x = 0; x < width; x++) {
          const shift = 8 - bitDepth * ((x % perByte) + 1);
          words[w0 + x] = table.rgba[(line[(x / perByte) | 0] >> shift) & mask];
        }
      }
      if (opaque && table.clear) for (let x = 0; x < width; x++) if (out[o + x * 4 + 3] !== 255) opaque = false;
    } else if (bitDepth === 8 && colorType === 2 && !trns && words) {
      for (let x = 0, s = 0, w = y * width; x < width; x++, s += 3, w++) words[w] = (line[s] | (line[s + 1] << 8) | (line[s + 2] << 16) | 0xff000000) >>> 0;
    } else if (bitDepth === 8 && colorType === 4 && words) {
      for (let x = 0, s = 0, w = y * width; x < width; x++, s += 2, w++) {
        const a = line[s + 1];
        words[w] = (line[s] * 0x010101 + a * 0x1000000) >>> 0;
        if (a !== 255) opaque = false;
      }
    } else if (bitDepth === 8 && colorType === 6) {
      out.set(line, o);
      if (opaque) for (let i = 3; i < stride; i += 4) if (line[i] !== 255) opaque = false;
    } else if (bitDepth === 8 && colorType === 2 && !trns) {
      for (let x = 0, s = 0; x < width; x++, s += 3, o += 4) {
        out[o] = line[s];
        out[o + 1] = line[s + 1];
        out[o + 2] = line[s + 2];
        out[o + 3] = 255;
      }
    } else if (bitDepth === 8 && colorType === 0 && !trns) {
      for (let x = 0; x < width; x++, o += 4) {
        out[o] = out[o + 1] = out[o + 2] = line[x];
        out[o + 3] = 255;
      }
    } else {
      for (let x = 0; x < width; x++, o += 4) {
        writePixel(out, o, line, x, bitDepth, colorType, palette, trns);
        if (out[o + 3] !== 255) opaque = false;
      }
    }
  }
  return { width, height, data: out, opaque };
}

async function decodeJpeg(data: Uint8Array, budget: ImageBudget): Promise<{ pixels: Pixmap; components: number } | null> {
  const size = jpegFrame(data);
  if (!size || !size.width || !size.height) return null;
  if (!budget.tryReserve(size.width, size.height)) throw new Error('Image pixel limit exceeded');
  try {
    const img = decodeJpegPixels(data);
    return { pixels: { width: img.width, height: img.height, data: img.data, opaque: true }, components: size.components };
  } catch {
    // Kinds the decoder leaves out, such as unusual sampling factors, fall back to jpeg-js.
  }
  const { decode } = await import('./jpeg-codec.js');
  const img = decode(data, {
    useTArray: true,
    formatAsRGBA: true,
    tolerantDecoding: true,
    maxResolutionInMP: Math.ceil((size.width * size.height) / 1_000_000),
    // jpeg-js keeps 32-bit coefficients for each of up to four components besides its output.
    maxMemoryUsageInMB: 64 + Math.ceil((size.width * size.height * 24) / (1024 * 1024)),
  });
  return { pixels: { width: img.width, height: img.height, data: img.data, opaque: true }, components: size.components };
}

/** Size and component count from a JPEG's frame header. */
function jpegFrame(d: Uint8Array): { width: number; height: number; components: number } | null {
  let p = 2;
  while (p + 4 <= d.length) {
    if (d[p] !== 0xff) {
      p++;
      continue;
    }
    const marker = d[p + 1];
    if (marker === 0xff) {
      p++;
      continue;
    }
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
      p += 2;
      continue;
    }
    const len = (d[p + 2] << 8) | d[p + 3];
    if (len < 2 || p + 2 + len > d.length) return null;
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (len < 8) return null;
      return { height: (d[p + 5] << 8) | d[p + 6], width: (d[p + 7] << 8) | d[p + 8], components: d[p + 9] };
    }
    if (marker === 0xda) return null;
    p += 2 + len;
  }
  return null;
}
