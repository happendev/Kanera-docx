export type ChunkSink = (chunk: Uint8Array<ArrayBuffer>) => void;
export type InflateBytes = (data: Uint8Array, limit: number) => Uint8Array;

export interface Compression {
  inflateRaw: InflateBytes;
  inflateZlib: InflateBytes;
  inflateZlibStream(data: Uint8Array, limit: number, sink: ChunkSink): void | Promise<void>;
  deflateZlib(data: Uint8Array): Uint8Array;
}

export function boundedSink(limit: number, sink: ChunkSink): ChunkSink {
  let remaining = limit;
  return (chunk) => {
    if (chunk.length > remaining) throw new Error('inflate: decoded size exceeds limit');
    remaining -= chunk.length;
    sink(chunk);
  };
}
