import { deflate, Inflate, Z_BUF_ERROR } from 'pako';
import { boundedSink, type ChunkSink, type Compression } from './compression.js';

function inflate(data: Uint8Array, limit: number, raw: boolean, sink?: ChunkSink): Uint8Array {
  const decoder = new Inflate({ raw, windowBits: 15, chunkSize: Math.min(65536, limit + 1) });
  // Only the one-shot path retains chunks to assemble a result.
  decoder.onData = boundedSink(limit, sink ?? decoder.onData.bind(decoder));
  if (!decoder.push(data, true)) {
    throw new Error(decoder.err === Z_BUF_ERROR ? 'inflate: truncated stream' : 'inflate: ' + decoder.msg);
  }
  return decoder.result;
}

export const pakoCompression: Compression = {
  inflateRaw: (data, limit) => inflate(data, limit, true),
  inflateZlib: (data, limit) => inflate(data, limit, false),
  inflateZlibStream(data, limit, sink) {
    inflate(data, limit, false, sink);
  },
  deflateZlib: deflate,
};
