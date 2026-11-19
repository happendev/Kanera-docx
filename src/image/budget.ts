export interface ImageLimits {
  maxImagePixels?: number;
  maxTotalPixels?: number;
}

function pixelLimit(value: number): number {
  if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('Image pixel limits must be non-negative safe integers');
  return value;
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
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || Math.max(width, height) > 32768)
      throw new Error('Image dimensions outside supported range');
    const pixels = width * height;
    if (pixels > this.maxImagePixels || pixels > this.remaining) throw new Error('Image pixel limit exceeded');
    this.remaining -= pixels;
  }
}
