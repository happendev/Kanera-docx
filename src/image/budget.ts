export interface ImageLimits {
  maxImagePixels?: number;
  maxTotalPixels?: number;
}

function pixelLimit(value: number): number {
  if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('Image pixel limits must be non-negative safe integers');
  return value;
}

function supported(width: number, height: number): boolean {
  return Number.isInteger(width) && Number.isInteger(height) && width >= 1 && height >= 1 && Math.max(width, height) <= 32768;
}

/** One reservation per source raster, before allocating scanlines or pixel buffers. */
export class ImageBudget {
  private readonly maxImagePixels: number;
  private remaining: number;

  constructor(limits: ImageLimits = {}) {
    this.maxImagePixels = pixelLimit(limits.maxImagePixels ?? 16_777_216);
    this.remaining = pixelLimit(limits.maxTotalPixels ?? 67_108_864);
  }

  reserve(width: number, height: number): void {
    if (!supported(width, height)) throw new Error('Image dimensions outside supported range');
    if (!this.tryReserve(width, height)) throw new Error('Image pixel limit exceeded');
  }

  /**
   * The smallest of 1, 2, 4 and 8 by which an image shrunk each way fits, reserved; 0 when none
   * does. Only an image up to four times the per-image limit may be shrunk: its decoder still reads
   * the full-size image.
   */
  tryReserveReduced(width: number, height: number): number {
    if (this.tryReserve(width, height)) return 1;
    if (!supported(width, height) || width * height > 4 * this.maxImagePixels) return 0;
    for (const reduce of [2, 4, 8]) if (this.tryReserve(Math.ceil(width / reduce), Math.ceil(height / reduce))) return reduce;
    return 0;
  }

  /** Reserve when the image fits, or return false for callers that can skip the decode instead. */
  tryReserve(width: number, height: number): boolean {
    const pixels = width * height;
    if (!supported(width, height) || pixels > this.maxImagePixels || pixels > this.remaining) return false;
    this.remaining -= pixels;
    return true;
  }
}
