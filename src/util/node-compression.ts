import { createInflate, deflateSync, inflateRawSync, inflateSync } from 'node:zlib';
import { boundedSink, type Compression } from './compression.js';

// ZIP parts and PDF streams are retained. Avoid pinning a full zlib output
// slab for each small result, and preserve Uint8Array's copying slice().
function resultBytes(buffer: Buffer): Uint8Array {
  return buffer.byteLength === buffer.buffer.byteLength ? new Uint8Array(buffer.buffer) : new Uint8Array(buffer);
}

function inflate(data: Uint8Array, limit: number, raw: boolean): Uint8Array {
  const decode = raw ? inflateRawSync : inflateSync;
  let result: Buffer;
  try {
    // Node requires a positive maxOutputLength, including for empty ZIP entries.
    result = decode(data, { maxOutputLength: Math.max(1, limit) });
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ERR_BUFFER_TOO_LARGE') throw new Error('inflate: decoded size exceeds limit', { cause: err });
    if (code === 'Z_BUF_ERROR') throw new Error('inflate: truncated stream', { cause: err });
    throw err;
  }
  if (result.length > limit) throw new Error('inflate: decoded size exceeds limit');
  return resultBytes(result);
}

export const nodeCompression: Compression = {
  inflateRaw: (data, limit) => inflate(data, limit, true),
  inflateZlib: (data, limit) => inflate(data, limit, false),
  async inflateZlibStream(data, limit, sink) {
    const write = boundedSink(limit, sink);
    const decoder = createInflate({ chunkSize: Math.max(64, Math.min(65536, limit + 1)) });
    decoder.end(data);
    for await (const chunk of decoder) write(chunk);
  },
  deflateZlib: (data) => resultBytes(deflateSync(data)),
};
