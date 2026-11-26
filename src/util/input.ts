export interface InputOptions {
  /** Maximum size of the source file in bytes, before decompression (default 256 MiB). */
  maxInputBytes?: number;
}

export function inputLimit(value: number | undefined): number {
  if (value === undefined) return 256 * 1024 * 1024;
  if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('maxInputBytes must be a non-negative safe integer');
  return value;
}

export function checkInputSize(size: number, limit: number): void {
  if (size > limit) throw new RangeError('Input byte limit exceeded');
}

export function inputBytes(input: Uint8Array | ArrayBuffer, maxInputBytes?: number): Uint8Array {
  const limit = inputLimit(maxInputBytes);
  if (ArrayBuffer.isView(input) && (input instanceof Uint8Array || Object.prototype.toString.call(input) === '[object Uint8Array]')) {
    checkInputSize(input.byteLength, limit);
    return input;
  }
  // DataView accepts buffers from other realms and never treats a number as an allocation size.
  let view: DataView;
  try {
    view = new DataView(input as ArrayBuffer);
  } catch {
    throw new TypeError('Expected a Uint8Array or ArrayBuffer document');
  }
  checkInputSize(view.byteLength, limit);
  return new Uint8Array(view.buffer);
}

export async function readInputBlob(blob: Blob, options: InputOptions & { signal?: AbortSignal }): Promise<Uint8Array> {
  options.signal?.throwIfAborted();
  checkInputSize(blob.size, inputLimit(options.maxInputBytes));
  const data = await blob.arrayBuffer();
  options.signal?.throwIfAborted();
  return inputBytes(data, options.maxInputBytes);
}
