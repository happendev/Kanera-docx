export type ChunkSink = (chunk: Uint8Array<ArrayBuffer>) => void;
export type InflateBytes = (data: Uint8Array, limit: number) => Uint8Array;

export interface Compression {
  inflateRaw: InflateBytes;
  inflateZlib: InflateBytes;
  inflateZlibStream(data: Uint8Array, limit: number, sink: ChunkSink): void | Promise<void>;
  /** zlib-compress `data`, at `level` 1 (fastest) to 9 (smallest) or the library's default. */
  deflateZlib(data: Uint8Array, level?: number): Uint8Array;
  /** The same, off the calling thread or natively where the platform can (page images use it). */
  deflateZlibAsync?(data: Uint8Array, level?: number): Promise<Uint8Array>;
  /** Off-thread compression with the exact same encoding as the default deflateZlib. */
  deflateZlibExactAsync?(data: Uint8Array): Promise<Uint8Array>;
  /** The CRC-32 of `data`, where the platform computes it natively (PNG chunks use it). */
  crc32?(data: Uint8Array): number;
}

export function boundedSink(limit: number, sink: ChunkSink): ChunkSink {
  let remaining = limit;
  return (chunk) => {
    if (chunk.length > remaining) throw new Error('inflate: decoded size exceeds limit');
    remaining -= chunk.length;
    sink(chunk);
  };
}
