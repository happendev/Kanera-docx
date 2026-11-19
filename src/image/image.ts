import type { Compression } from '../util/compression.js';
import { ImageBudget } from './budget.js';
import { decodeTiff } from './tiff.js';

/** An image ready to be written as a PDF image XObject. */
export interface PdfImage {
  width: number;
  height: number;
  /** PDF colour space: a name like `/DeviceRGB`, or an inline array for indexed images. */
  colorSpace: string;
  bitsPerComponent: number;
  filter: 'DCTDecode' | 'FlateDecode';
  data: Uint8Array;
  decodeParms?: string;
  decode?: number[];
  /** 8-bit alpha channel, already zlib-compressed. */
  smask?: Uint8Array;
}

export type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'tiff' | 'emf' | 'wmf' | 'svg' | 'unknown';

export function sniffImage(d: Uint8Array): ImageFormat {
  if (d.length < 8) return 'unknown';
  if (d[0] === 0x89 && d[1] === 0x50 && d[2] === 0x4e && d[3] === 0x47) return 'png';
  if (d[0] === 0xff && d[1] === 0xd8) return 'jpeg';
  if (d[0] === 0x47 && d[1] === 0x49 && d[2] === 0x46) return 'gif';
  if (d[0] === 0x42 && d[1] === 0x4d) return 'bmp';
  if ((d[0] === 0x49 && d[1] === 0x49 && d[2] === 42) || (d[0] === 0x4d && d[1] === 0x4d && d[3] === 42)) return 'tiff';
  if (d[0] === 1 && d[1] === 0 && d[2] === 0 && d[3] === 0 && d[40] === 0x20 && d[41] === 0x45 && d[42] === 0x4d && d[43] === 0x46) return 'emf';
  if ((d[0] === 0xd7 && d[1] === 0xcd && d[2] === 0xc6 && d[3] === 0x9a) || (d[0] === 1 && d[1] === 0 && d[2] === 9 && d[3] === 0)) return 'wmf';
  const head = new TextDecoder().decode(d.subarray(0, Math.min(d.length, 512)));
  if (/<(?:[\w.-]+:)?svg[\s/>]/.test(head) || /^\s*<\?xml/.test(head)) return 'svg';
  return 'unknown';
}

/** Pixel size of an image without fully decoding it (null if unknown). */
export function imageSize(d: Uint8Array): { width: number; height: number } | null {
  const fmt = sniffImage(d);
  if (fmt === 'png') return { width: be32(d, 16), height: be32(d, 20) };
  if (fmt === 'jpeg') {
    const info = jpegInfo(d);
    return info ? { width: info.width, height: info.height } : null;
  }
  if (fmt === 'gif') return { width: d[6] | (d[7] << 8), height: d[8] | (d[9] << 8) };
  if (fmt === 'bmp') return { width: le32s(d, 18), height: Math.abs(le32s(d, 22)) };
  return null;
}

export async function toPdfImage(
  d: Uint8Array, compression: Compression, budget = new ImageBudget(), quality?: number | false,
): Promise<PdfImage | null> {
  switch (sniffImage(d)) {
    case 'png':
      return pngToPdf(d, compression, budget);
    case 'jpeg':
      return jpegToPdf(d, budget, quality);
    case 'gif': {
      const rgba = decodeGif(d, budget);
      return rgba ? rgbaToPdf(rgba.width, rgba.height, rgba.data, compression) : null;
    }
    case 'bmp': {
      const rgba = decodeBmp(d, budget);
      return rgba ? rgbaToPdf(rgba.width, rgba.height, rgba.data, compression) : null;
    }
    case 'tiff': {
      const rgba = decodeTiff(d, compression.inflateZlib, budget);
      return rgba ? rgbaToPdf(rgba.width, rgba.height, rgba.data, compression) : null;
    }
    default:
      return null;
  }
}

function be32(d: Uint8Array, o: number): number {
  return ((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]) >>> 0;
}
function le32s(d: Uint8Array, o: number): number {
  return d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24);
}
function le16(d: Uint8Array, o: number): number {
  return d[o] | (d[o + 1] << 8);
}

function rgbaToPdf(width: number, height: number, rgba: Uint8Array, compression: Compression): PdfImage {
  const px = width * height;
  const rgb = new Uint8Array(px * 3);
  const alpha = new Uint8Array(px);
  let hasAlpha = false;
  for (let i = 0; i < px; i++) {
    rgb[i * 3] = rgba[i * 4];
    rgb[i * 3 + 1] = rgba[i * 4 + 1];
    rgb[i * 3 + 2] = rgba[i * 4 + 2];
    const a = rgba[i * 4 + 3];
    alpha[i] = a;
    if (a !== 255) hasAlpha = true;
  }
  const plain = compression.deflateZlib(rgb);
  // Smooth images and scans often compress much better as differences between pixels.
  // Compare against the raw stream since a predictor can increase the size of some images.
  const decodeParms = `<< /Predictor 15 /Colors 3 /BitsPerComponent 8 /Columns ${width} >>`;
  const predicted = plain.length > 1024 ? predictedRgb(width, height, rgb, compression) : undefined;
  const usePredictor = predicted !== undefined && predicted.length + decodeParms.length + 14 < plain.length;
  return {
    width,
    height,
    colorSpace: '/DeviceRGB',
    bitsPerComponent: 8,
    filter: 'FlateDecode',
    data: usePredictor ? predicted : plain,
    decodeParms: usePredictor ? decodeParms : undefined,
    smask: hasAlpha ? compression.deflateZlib(alpha) : undefined,
  };
}

function predictedRgb(width: number, height: number, rgb: Uint8Array, compression: Compression): Uint8Array {
  const stride = width * 3;
  const rows = new Uint8Array((stride + 1) * height);
  for (let y = 0; y < height; y++) {
    const src = y * stride;
    const dst = y * (stride + 1);
    let subScore = 0;
    let upScore = 0;
    for (let x = 0; x < stride; x++) {
      const sub = (rgb[src + x] - (x < 3 ? 0 : rgb[src + x - 3])) & 255;
      const up = (rgb[src + x] - (y === 0 ? 0 : rgb[src + x - stride])) & 255;
      subScore += Math.min(sub, 256 - sub);
      upScore += Math.min(up, 256 - up);
    }
    const up = y > 0 && upScore < subScore;
    rows[dst] = up ? 2 : 1;
    for (let x = 0; x < stride; x++) {
      const prev = up ? rgb[src + x - stride] : x < 3 ? 0 : rgb[src + x - 3];
      rows[dst + x + 1] = rgb[src + x] - prev;
    }
  }
  return compression.deflateZlib(rows);
}

// ---------------------------------------------------------------------------------------------
// PNG

interface PngInfo {
  width: number;
  height: number;
  bitDepth: number;
  colorType: number;
  interlace: number;
  palette?: Uint8Array;
  trns?: Uint8Array;
  idat: Uint8Array;
}

function readPng(d: Uint8Array, budget: ImageBudget): PngInfo {
  if (d.length < 33 || be32(d, 4) !== 0x0d0a1a0a || be32(d, 8) !== 13 || be32(d, 12) !== 0x49484452)
    throw new Error('PNG: invalid header');
  const width = be32(d, 16);
  const height = be32(d, 20);
  const bitDepth = d[24];
  const colorType = d[25];
  const interlace = d[28];
  const depths = colorType === 0 ? [1, 2, 4, 8, 16] : colorType === 3 ? [1, 2, 4, 8] : [8, 16];
  if (![0, 2, 3, 4, 6].includes(colorType) || !depths.includes(bitDepth) || d[26] || d[27] || interlace > 1)
    throw new Error('PNG: invalid image parameters');
  budget.reserve(width, height);

  let palette: Uint8Array | undefined;
  let trns: Uint8Array | undefined;
  let firstIdat = 0;
  let total = 0;
  let p = 33;
  for (;;) {
    if (p + 12 > d.length) throw new Error('PNG: missing end chunk');
    const len = be32(d, p);
    if (p + 12 + len > d.length) throw new Error('PNG: truncated chunk');
    const type = String.fromCharCode(...d.subarray(p + 4, p + 8));
    if (type === 'IHDR') throw new Error('PNG: duplicate header');
    if (type === 'PLTE') {
      if (!len || len > 768 || len % 3) throw new Error('PNG: invalid palette');
      palette = d.subarray(p + 8, p + 8 + len);
    } else if (type === 'tRNS') trns = d.subarray(p + 8, p + 8 + len);
    else if (type === 'IDAT') {
      if (!firstIdat) firstIdat = p;
      total += len;
    } else if (type === 'IEND') {
      if (len || !total || (colorType === 3 && !palette)) throw new Error('PNG: incomplete image');
      break;
    }
    p += 12 + len;
  }
  let idat = d.subarray(firstIdat + 8, firstIdat + 8 + be32(d, firstIdat));
  if (idat.length !== total) {
    idat = new Uint8Array(total);
    let pos = 0;
    for (let i = firstIdat; i < p; ) {
      const len = be32(d, i);
      if (be32(d, i + 4) === 0x49444154) {
        idat.set(d.subarray(i + 8, i + 8 + len), pos);
        pos += len;
      }
      i += 12 + len;
    }
  }
  return { width, height, bitDepth, colorType, interlace, palette, trns, idat };
}

async function pngToPdf(d: Uint8Array, compression: Compression, budget: ImageBudget): Promise<PdfImage | null> {
  const info = readPng(d, budget);
  const { width, height, bitDepth, colorType, interlace, palette, trns, idat } = info;
  const channels = colorType === 2 ? 3 : colorType === 0 ? 1 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4;
  // Fast path: pass the compressed scanlines straight through with a PNG predictor.
  if (!interlace && !trns && (colorType === 0 || colorType === 2 || colorType === 3)) {
    const rowBytes = 1 + Math.ceil(width * channels * bitDepth / 8);
    let bytes = 0;
    await compression.inflateZlibStream(idat, rowBytes * height, (chunk) => {
      for (let i = (rowBytes - bytes % rowBytes) % rowBytes; i < chunk.length; i += rowBytes)
        if (chunk[i] > 4) throw new Error('PNG: invalid scanline filter');
      bytes += chunk.length;
    });
    if (bytes !== rowBytes * height) throw new Error('PNG: truncated scanlines');
    let colorSpace = colorType === 2 ? '/DeviceRGB' : '/DeviceGray';
    if (colorType === 3) {
      if (!palette) return null;
      let hex = '';
      for (const b of palette) hex += b.toString(16).padStart(2, '0');
      colorSpace = `[/Indexed /DeviceRGB ${palette.length / 3 - 1} <${hex}>]`;
    }
    return {
      width,
      height,
      colorSpace,
      bitsPerComponent: bitDepth,
      filter: 'FlateDecode',
      data: idat,
      decodeParms: `<< /Predictor 15 /Colors ${channels} /BitsPerComponent ${bitDepth} /Columns ${width} >>`,
    };
  }
  if (!interlace && width * height > MAX_DECODED_PIXELS) {
    const small = await decodePngReduced(info, compression);
    return rgbaToPdf(small.width, small.height, small.data, compression);
  }
  const rgba = decodePngRgba(info, compression);
  return rgbaToPdf(width, height, rgba, compression);
}

/** Larger pictures are reduced while decoding, to keep memory and file size bounded. */
const MAX_DECODED_PIXELS = 1 << 23;

/**
 * Decode a non-interlaced PNG row by row, averaging k×k blocks (weighted by alpha) so the result
 * stays within MAX_DECODED_PIXELS; the full image is never held in memory.
 */
async function decodePngReduced(info: PngInfo, compression: Compression): Promise<{ width: number; height: number; data: Uint8Array }> {
  const { width, height, bitDepth, colorType, palette, trns } = info;
  const channels = colorType === 2 ? 3 : colorType === 0 ? 1 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4;
  const bitsPP = channels * bitDepth;
  const bpp = Math.max(1, bitsPP >> 3);
  const stride = Math.ceil((width * bitsPP) / 8);
  const k = Math.ceil(Math.sqrt((width * height) / MAX_DECODED_PIXELS));
  const ow = Math.ceil(width / k);
  const oh = Math.ceil(height / k);
  const out = new Uint8Array(ow * oh * 4);
  const sums = new Float64Array(ow * 5);
  const rowRgba = new Uint8Array(width * 4);
  let line = new Uint8Array(stride);
  let prev = new Uint8Array(stride);
  let filter = -1;
  let fill = 0;
  let y = 0;
  const flushBand = (oy: number) => {
    for (let i = 0; i < ow; i++) {
      const s = i * 5;
      const n = sums[s + 4];
      const a = sums[s + 3];
      const o = (oy * ow + i) * 4;
      if (a > 0) for (let c = 0; c < 3; c++) out[o + c] = Math.round(sums[s + c] / a);
      out[o + 3] = n > 0 ? Math.round(a / n) : 0;
    }
    sums.fill(0);
  };
  const finishRow = () => {
    unfilter(filter, line, prev, bpp);
    for (let x = 0; x < width; x++) writePixel(rowRgba, x * 4, line, x, bitDepth, colorType, palette, trns);
    for (let x = 0; x < width; x++) {
      const s = Math.floor(x / k) * 5;
      const a = rowRgba[x * 4 + 3];
      sums[s] += rowRgba[x * 4] * a;
      sums[s + 1] += rowRgba[x * 4 + 1] * a;
      sums[s + 2] += rowRgba[x * 4 + 2] * a;
      sums[s + 3] += a;
      sums[s + 4]++;
    }
    if (y % k === k - 1 || y === height - 1) flushBand(Math.floor(y / k));
    [prev, line] = [line, prev];
    y++;
    filter = -1;
    fill = 0;
  };
  await compression.inflateZlibStream(
    info.idat,
    height * (stride + 1),
    (chunk) => {
      for (let i = 0; i < chunk.length && y < height; ) {
        if (filter < 0) {
          filter = chunk[i++];
          continue;
        }
        const take = Math.min(stride - fill, chunk.length - i);
        line.set(chunk.subarray(i, i + take), fill);
        fill += take;
        i += take;
        if (fill === stride) finishRow();
      }
    },
  );
  if (y !== height) throw new Error('PNG: truncated scanlines');
  return { width: ow, height: oh, data: out };
}

const ADAM7 = [
  [0, 0, 8, 8],
  [4, 0, 8, 8],
  [0, 4, 4, 8],
  [2, 0, 4, 4],
  [0, 2, 2, 4],
  [1, 0, 2, 2],
  [0, 1, 1, 2],
];

function decodePngRgba(info: PngInfo, compression: Compression): Uint8Array {
  const { width, height, bitDepth, colorType, interlace, palette, trns } = info;
  const channels = colorType === 2 ? 3 : colorType === 0 ? 1 : colorType === 3 ? 1 : colorType === 4 ? 2 : 4;
  const bitsPP = channels * bitDepth;
  const bpp = Math.max(1, bitsPP >> 3);
  let pos = 0;

  const passes = interlace ? ADAM7 : [[0, 0, 1, 1]];
  let expected = 0;
  for (const [x0, y0, dx, dy] of passes) {
    const pw = Math.ceil((width - x0) / dx);
    const ph = Math.ceil((height - y0) / dy);
    if (pw > 0 && ph > 0) expected += (1 + Math.ceil(pw * bitsPP / 8)) * ph;
  }
  const raw = compression.inflateZlib(info.idat, expected);
  if (raw.length !== expected) throw new Error('PNG: truncated scanlines');
  const out = new Uint8Array(width * height * 4);
  for (const [x0, y0, dx, dy] of passes) {
    const pw = Math.ceil((width - x0) / dx);
    const ph = Math.ceil((height - y0) / dy);
    if (pw <= 0 || ph <= 0) continue;
    const stride = Math.ceil((pw * bitsPP) / 8);
    let prev: Uint8Array = new Uint8Array(stride);
    for (let row = 0; row < ph; row++) {
      const filter = raw[pos++];
      const line = raw.subarray(pos, pos + stride);
      pos += stride;
      unfilter(filter, line, prev, bpp);
      prev = line;
      const y = y0 + row * dy;
      for (let col = 0; col < pw; col++) {
        const x = x0 + col * dx;
        const o = (y * width + x) * 4;
        writePixel(out, o, line, col, bitDepth, colorType, palette, trns);
      }
    }
  }
  return out;
}

function unfilter(f: number, line: Uint8Array, prev: Uint8Array, bpp: number): void {
  const n = line.length;
  switch (f) {
    case 0:
      return;
    case 1:
      for (let i = bpp; i < n; i++) line[i] = (line[i] + line[i - bpp]) & 255;
      return;
    case 2:
      for (let i = 0; i < n; i++) line[i] = (line[i] + prev[i]) & 255;
      return;
    case 3:
      for (let i = 0; i < n; i++) line[i] = (line[i] + (((i >= bpp ? line[i - bpp] : 0) + prev[i]) >> 1)) & 255;
      return;
    case 4:
      for (let i = 0; i < n; i++) {
        const a = i >= bpp ? line[i - bpp] : 0;
        const b = prev[i];
        const c = i >= bpp ? prev[i - bpp] : 0;
        const p = a + b - c;
        const pa = Math.abs(p - a);
        const pb = Math.abs(p - b);
        const pc = Math.abs(p - c);
        line[i] = (line[i] + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c)) & 255;
      }
      return;
    default:
      throw new Error(`PNG: bad filter ${f}`);
  }
}

function sample(line: Uint8Array, index: number, bitDepth: number): number {
  if (bitDepth === 8) return line[index];
  if (bitDepth === 16) return line[index * 2];
  const perByte = 8 / bitDepth;
  const byte = line[Math.floor(index / perByte)];
  const shift = 8 - bitDepth * ((index % perByte) + 1);
  const v = (byte >> shift) & ((1 << bitDepth) - 1);
  return v;
}

function writePixel(
  out: Uint8Array,
  o: number,
  line: Uint8Array,
  col: number,
  bitDepth: number,
  colorType: number,
  palette?: Uint8Array,
  trns?: Uint8Array,
): void {
  const scale = bitDepth < 8 ? 255 / ((1 << bitDepth) - 1) : 1;
  switch (colorType) {
    case 0: {
      const raw = sample(line, col, bitDepth);
      const g = Math.round(raw * scale);
      out[o] = out[o + 1] = out[o + 2] = g;
      let a = 255;
      if (trns && trns.length >= 2) {
        const t = (trns[0] << 8) | trns[1];
        const rawFull = bitDepth === 16 ? (line[col * 2] << 8) | line[col * 2 + 1] : raw;
        if (rawFull === t) a = 0;
      }
      out[o + 3] = a;
      return;
    }
    case 2: {
      out[o] = sample(line, col * 3, bitDepth);
      out[o + 1] = sample(line, col * 3 + 1, bitDepth);
      out[o + 2] = sample(line, col * 3 + 2, bitDepth);
      let a = 255;
      if (trns && trns.length >= 6) {
        if (bitDepth === 16) {
          const r = (line[col * 6] << 8) | line[col * 6 + 1];
          const g = (line[col * 6 + 2] << 8) | line[col * 6 + 3];
          const b = (line[col * 6 + 4] << 8) | line[col * 6 + 5];
          if (r === ((trns[0] << 8) | trns[1]) && g === ((trns[2] << 8) | trns[3]) && b === ((trns[4] << 8) | trns[5])) a = 0;
        } else if (out[o] === trns[1] && out[o + 1] === trns[3] && out[o + 2] === trns[5]) a = 0;
      }
      out[o + 3] = a;
      return;
    }
    case 3: {
      const idx = sample(line, col, bitDepth);
      if (palette && idx * 3 + 2 < palette.length) {
        out[o] = palette[idx * 3];
        out[o + 1] = palette[idx * 3 + 1];
        out[o + 2] = palette[idx * 3 + 2];
      }
      out[o + 3] = trns && idx < trns.length ? trns[idx] : 255;
      return;
    }
    case 4: {
      const g = sample(line, col * 2, bitDepth);
      out[o] = out[o + 1] = out[o + 2] = g;
      out[o + 3] = sample(line, col * 2 + 1, bitDepth);
      return;
    }
    case 6: {
      out[o] = sample(line, col * 4, bitDepth);
      out[o + 1] = sample(line, col * 4 + 1, bitDepth);
      out[o + 2] = sample(line, col * 4 + 2, bitDepth);
      out[o + 3] = sample(line, col * 4 + 3, bitDepth);
      return;
    }
  }
}

// ---------------------------------------------------------------------------------------------
// JPEG

interface JpegInfo {
  width: number;
  height: number;
  components: number;
  adobeTransform: number | null;
  canReencode: boolean;
}

function jpegInfo(d: Uint8Array): JpegInfo | null {
  let p = 2;
  let adobe: number | null = null;
  let hasColorProfile = false;
  let frame: Omit<JpegInfo, 'adobeTransform' | 'canReencode'> | undefined;
  let supportedFrame = false;
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
    if (len < 2 || p + 2 + len > d.length) break;
    if (marker === 0xee && len >= 14 && d[p + 4] === 0x41 && d[p + 5] === 0x64 && d[p + 6] === 0x6f && d[p + 7] === 0x62 && d[p + 8] === 0x65) {
      adobe = d[p + 15];
    }
    if (marker === 0xe2 && len >= 14 && new TextDecoder().decode(d.subarray(p + 4, p + 16)) === 'ICC_PROFILE\0')
      hasColorProfile = true;
    const isSof = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      if (len < 8 || len < 8 + 3 * d[p + 9]) return null;
      frame = {
        height: (d[p + 5] << 8) | d[p + 6],
        width: (d[p + 7] << 8) | d[p + 8],
        components: d[p + 9],
      };
      supportedFrame = (marker === 0xc0 || marker === 0xc1 || marker === 0xc2) && d[p + 4] === 8;
    }
    if (marker === 0xda) break;
    p += 2 + len;
  }
  return frame ? { ...frame, adobeTransform: adobe, canReencode: supportedFrame && frame.components === 3 && !hasColorProfile } : null;
}

async function jpegToPdf(d: Uint8Array, budget: ImageBudget, quality?: number | false): Promise<PdfImage | null> {
  const info = jpegInfo(d);
  if (!info) return null;
  budget.reserve(info.width, info.height);
  const colorSpace = info.components === 1 ? '/DeviceGray' : info.components === 4 ? '/DeviceCMYK' : '/DeviceRGB';
  let data = d;
  if (quality !== false && info.canReencode && (quality !== undefined || d.length > 1024)) {
    try {
      data = (await smallerJpeg(d, info.width, info.height, quality)) ?? d;
    } catch {
      // The PDF reader can still handle JPEGs the optional recompressor cannot decode.
    }
  }
  return {
    width: info.width,
    height: info.height,
    colorSpace,
    bitsPerComponent: 8,
    filter: 'DCTDecode',
    data,
    // Adobe-flavoured CMYK JPEGs store inverted values.
    decode: info.components === 4 && info.adobeTransform !== null ? [1, 0, 1, 0, 1, 0, 1, 0] : undefined,
  };
}

async function smallerJpeg(data: Uint8Array, width: number, height: number, quality?: number): Promise<Uint8Array | undefined> {
  const { encode, decode } = await import('./jpeg-codec.js');
  const decodeOptions = {
    useTArray: true as const,
    tolerantDecoding: false,
    maxResolutionInMP: Math.ceil(width * height / 1_000_000),
    maxMemoryUsageInMB: 256,
  };
  const pixels = decode(data, decodeOptions);
  if (pixels.width !== width || pixels.height !== height) return undefined;
  const jpeg = encode(pixels, quality ?? 80).data;
  if (quality !== undefined) return jpeg.length < data.length ? jpeg : undefined;
  if (jpeg.length > data.length * 0.9) return undefined;
  const { data: candidate } = decode(jpeg, decodeOptions);
  // Guard the default against visible changes even when re-encoding yields a much smaller stream.
  let squaredError = 0;
  for (let i = 0; i < candidate.length; i += 4) {
    for (let c = 0; c < 3; c++) {
      const error = candidate[i + c] - pixels.data[i + c];
      if (Math.abs(error) > 32) return undefined;
      squaredError += error * error;
    }
  }
  return squaredError <= width * height * 3 * 9 ? jpeg : undefined;
}

// ---------------------------------------------------------------------------------------------
// GIF (first frame)

function decodeGif(d: Uint8Array, budget: ImageBudget): { width: number; height: number; data: Uint8Array } | null {
  const width = le16(d, 6);
  const height = le16(d, 8);
  if (d.length < 13) throw new Error('GIF: truncated header');
  budget.reserve(width, height);
  const flags = d[10];
  let p = 13;
  let gct: Uint8Array | null = null;
  if (flags & 0x80) {
    const size = 3 * (1 << ((flags & 7) + 1));
    gct = d.subarray(p, p + size);
    p += size;
  }
  let transparent = -1;
  while (p < d.length) {
    const b = d[p++];
    if (b === 0x21) {
      const label = d[p++];
      if (label === 0xf9) {
        const pf = d[p + 1];
        if (pf & 1) transparent = d[p + 4];
      }
      while (p < d.length) {
        const sz = d[p++];
        if (sz === 0) break;
        p += sz;
      }
    } else if (b === 0x2c) {
      const ix = le16(d, p);
      const iy = le16(d, p + 2);
      const iw = le16(d, p + 4);
      const ih = le16(d, p + 6);
      if (p + 9 > d.length || !iw || !ih || ix + iw > width || iy + ih > height) throw new Error('GIF: invalid frame dimensions');
      const iflags = d[p + 8];
      p += 9;
      let ct = gct;
      if (iflags & 0x80) {
        const size = 3 * (1 << ((iflags & 7) + 1));
        ct = d.subarray(p, p + size);
        p += size;
      }
      const interlaced = (iflags & 0x40) !== 0;
      const minCode = d[p++];
      if (!ct || minCode < 2 || minCode > 8 || p >= d.length) throw new Error('GIF: invalid image data');
      const start = p;
      let length = 0;
      while (p < d.length && d[p]) {
        const size = d[p++];
        length += size;
        p += size;
      }
      if (p >= d.length) throw new Error('GIF: truncated image data');
      const codes = new Uint8Array(length);
      let dest = 0;
      for (let i = start; i < p; ) {
        const size = d[i++];
        codes.set(d.subarray(i, i + size), dest);
        i += size;
        dest += size;
      }
      const indices = lzwDecode(codes, minCode, iw * ih);
      const out = new Uint8Array(width * height * 4);
      const rows: number[] = [];
      if (interlaced) {
        for (let y = 0; y < ih; y += 8) rows.push(y);
        for (let y = 4; y < ih; y += 8) rows.push(y);
        for (let y = 2; y < ih; y += 4) rows.push(y);
        for (let y = 1; y < ih; y += 2) rows.push(y);
      } else for (let y = 0; y < ih; y++) rows.push(y);
      for (let r = 0; r < ih; r++) {
        const y = iy + rows[r];
        for (let x = 0; x < iw; x++) {
          const px = ix + x;
          const idx = indices[r * iw + x];
          const o = (y * width + px) * 4;
          if (idx === transparent) {
            out[o + 3] = 0;
            continue;
          }
          if (idx * 3 + 2 >= ct.length) throw new Error('GIF: invalid palette index');
          out[o] = ct[idx * 3];
          out[o + 1] = ct[idx * 3 + 1];
          out[o + 2] = ct[idx * 3 + 2];
          out[o + 3] = 255;
        }
      }
      return { width, height, data: out };
    } else if (b === 0x3b) break;
  }
  return null;
}

function lzwDecode(data: Uint8Array, minCode: number, pixels: number): Uint8Array {
  const out = new Uint8Array(pixels);
  const clear = 1 << minCode;
  const eoi = clear + 1;
  let codeSize = minCode + 1;
  let dictSize = eoi + 1;
  const prefix = new Int32Array(4096);
  const suffix = new Uint8Array(4096);
  const first = new Uint8Array(4096);
  for (let i = 0; i < clear; i++) {
    prefix[i] = -1;
    suffix[i] = i;
    first[i] = i;
  }
  let bitPos = 0;
  let prev = -1;
  let op = 0;
  const stack = new Uint8Array(4097);
  while (op < pixels) {
    let code = 0;
    for (let i = 0; i < codeSize; i++) {
      const byte = data[(bitPos + i) >> 3];
      if (byte === undefined) throw new Error('GIF: truncated pixel data');
      code |= ((byte >> ((bitPos + i) & 7)) & 1) << i;
    }
    bitPos += codeSize;
    if (code === clear) {
      codeSize = minCode + 1;
      dictSize = eoi + 1;
      prev = -1;
      continue;
    }
    if (code === eoi) break;
    let cur = code;
    let sp = 0;
    if (code >= dictSize) {
      if (code !== dictSize || prev < 0) throw new Error('GIF: invalid LZW code');
      stack[sp++] = first[prev];
      cur = prev;
    }
    while (cur >= 0 && sp < 4097) {
      stack[sp++] = suffix[cur];
      cur = prefix[cur];
    }
    const firstChar = stack[sp - 1];
    while (sp > 0 && op < pixels) out[op++] = stack[--sp];
    if (prev >= 0 && dictSize < 4096) {
      prefix[dictSize] = prev;
      suffix[dictSize] = firstChar;
      first[dictSize] = first[prev];
      dictSize++;
      if (dictSize === 1 << codeSize && codeSize < 12) codeSize++;
    }
    prev = code;
  }
  if (op !== pixels) throw new Error('GIF: truncated pixel data');
  return out;
}

// ---------------------------------------------------------------------------------------------
// BMP

function decodeBmp(d: Uint8Array, budget: ImageBudget): { width: number; height: number; data: Uint8Array } | null {
  const offset = le32s(d, 10);
  const headerSize = le32s(d, 14);
  const width = le32s(d, 18);
  const rawH = le32s(d, 22);
  const height = Math.abs(rawH);
  const bpp = le16(d, 28);
  const compression = le32s(d, 30);
  if (d.length < 54 || headerSize < 40 || ![1, 4, 8, 24, 32].includes(bpp) || (compression !== 0 && compression !== 3)) return null;
  budget.reserve(width, height);
  const topDown = rawH < 0;
  const stride = Math.ceil((width * bpp) / 32) * 4;
  const palOff = 14 + headerSize;
  if (offset < palOff || offset + height * stride > d.length) throw new Error('BMP: truncated pixel data');
  const out = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y++) {
    const row = offset + (topDown ? y : height - 1 - y) * stride;
    for (let x = 0; x < width; x++) {
      const o = (y * width + x) * 4;
      if (bpp === 24 || bpp === 32) {
        const p = row + x * (bpp / 8);
        out[o] = d[p + 2];
        out[o + 1] = d[p + 1];
        out[o + 2] = d[p];
        out[o + 3] = 255;
      } else {
        const perByte = 8 / bpp;
        const byte = d[row + Math.floor(x / perByte)];
        const idx = (byte >> (8 - bpp * ((x % perByte) + 1))) & ((1 << bpp) - 1);
        const pp = palOff + idx * 4;
        if (pp + 4 > offset) throw new Error('BMP: truncated palette');
        out[o] = d[pp + 2];
        out[o + 1] = d[pp + 1];
        out[o + 2] = d[pp];
        out[o + 3] = 255;
      }
    }
  }
  return { width, height, data: out };
}
