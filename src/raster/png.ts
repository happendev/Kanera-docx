// PNG encoding of RGB pages: 8-bit truecolour, zlib-compressed. Rows are left unfiltered, which
// on document pages compresses smaller than Sub or Up (deflate already matches a row that repeats
// the one above), except white rows, which the Sub filter turns to zeros after the first pixel.

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function tableCrc32(d: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < d.length; i++) c = CRC_TABLE[(c ^ d[i]) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * Filtered scanlines of an RGB image, ready for zlib, written to `out` when given (any contents,
 * `(width * 3 + 1) * height` bytes). Pixels outside each row's [lo, hi) are white.
 */
export function pngScanlines(width: number, height: number, rgb: Uint8Array, lo: Int32Array, hi: Int32Array, out = new Uint8Array((width * 3 + 1) * height)): Uint8Array<ArrayBuffer> {
  const stride = width * 3;
  for (let y = 0; y < height; y++) {
    const o = y * (stride + 1);
    if (lo[y] < hi[y]) {
      out[o] = 0;
      out.set(rgb.subarray(y * stride, y * stride + stride), o + 1);
    } else {
      out[o] = 1;
      out[o + 1] = out[o + 2] = out[o + 3] = 255;
      out.fill(0, o + 4, o + stride + 1);
    }
  }
  return out;
}

/**
 * A PNG file of an RGB image whose filtered scanlines `deflated` already compressed (zlib), with
 * the platform's CRC-32 where it has one.
 */
export function pngFile(width: number, height: number, deflated: Uint8Array, dpi: number, crc32: (data: Uint8Array) => number = tableCrc32): Uint8Array<ArrayBuffer> {
  const out = new Uint8Array(8 + 25 + 21 + 12 + deflated.length + 12);
  const view = new DataView(out.buffer);
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  let p = 8;
  const chunk = (type: string, length: number, write: (at: number) => void) => {
    view.setUint32(p, length);
    for (let i = 0; i < 4; i++) out[p + 4 + i] = type.charCodeAt(i);
    write(p + 8);
    view.setUint32(p + 8 + length, crc32(out.subarray(p + 4, p + 8 + length)) >>> 0);
    p += 12 + length;
  };
  chunk('IHDR', 13, (at) => {
    view.setUint32(at, width);
    view.setUint32(at + 4, height);
    out[at + 8] = 8;
    out[at + 9] = 2;
  });
  // Pixels per metre, so viewers and editors know the resolution.
  const ppm = Math.round(dpi / 0.0254);
  chunk('pHYs', 9, (at) => {
    view.setUint32(at, ppm);
    view.setUint32(at + 4, ppm);
    out[at + 8] = 1;
  });
  chunk('IDAT', deflated.length, (at) => out.set(deflated, at));
  chunk('IEND', 0, () => {});
  return out;
}
