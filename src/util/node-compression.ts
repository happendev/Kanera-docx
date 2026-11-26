import * as zlib from 'node:zlib';
import { constants, createDeflate, createInflate, deflate, deflateSync, inflateRawSync, inflateSync, type ZlibOptions } from 'node:zlib';
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
    const maxOutputLength = Math.max(1, limit);
    const chunkSize = Math.max(64, Math.min(65536, limit + 1));
    result = decode(data, { maxOutputLength, chunkSize });
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    if (code === 'ERR_BUFFER_TOO_LARGE') throw new RangeError('inflate: decoded size exceeds limit', { cause: err });
    if (code === 'Z_BUF_ERROR') throw new Error('inflate: truncated stream', { cause: err });
    throw err;
  }
  if (result.length > limit) throw new RangeError('inflate: decoded size exceeds limit');
  return resultBytes(result);
}

function deflateAsync(data: Uint8Array, opts: ZlibOptions): Promise<Buffer> {
  return new Promise((resolve, reject) => deflate(data, opts, (err, out) => (err ? reject(err) : resolve(out))));
}

/** Inputs at least this large are compressed in parts at once, on several of libuv's threads. */
const PARALLEL_BYTES = 1 << 20;
const MAX_PARTS = 4;
const ADLER_BASE = 65521;

/** The Adler-32 checksum of a + b, from theirs and b's length (zlib's adler32_combine). */
function adlerCombine(a: number, b: number, bLength: number): number {
  const rem = bLength % ADLER_BASE;
  const a1 = a & 0xffff;
  const s2 = (rem * a1) % ADLER_BASE;
  const lo = (a1 + (b & 0xffff) + ADLER_BASE - 1) % ADLER_BASE;
  const hi = (s2 + (a >>> 16) + (b >>> 16) + ADLER_BASE - rem) % ADLER_BASE;
  return ((hi << 16) | lo) >>> 0;
}

/**
 * A part of a zlib stream: its deflate blocks, ended on a byte boundary, and its checksum. The
 * part goes through a zlib stream of its own (primed with the bytes before it), flushed and then
 * finished, so the output ends with the flush marker, an empty final block and the checksum.
 */
function deflatePart(part: Uint8Array, level: number, dictionary: Uint8Array | undefined): Promise<{ body: Uint8Array; adler: number } | null> {
  return new Promise((resolve, reject) => {
    const z = createDeflate({ level, chunkSize: Math.max(65536, part.length >> 1), ...(dictionary && { dictionary }) });
    const chunks: Buffer[] = [];
    z.on('data', (c: Buffer) => chunks.push(c));
    z.on('error', reject);
    z.on('end', () => {
      const out = Buffer.concat(chunks);
      const n = out.length;
      const head = dictionary ? 6 : 2;
      // Expect 00 00 ff ff (the flush), 03 00 (the empty final block), then the checksum.
      const ok = n >= head + 10 && out.readUInt32BE(n - 10) === 0xffff && out[n - 6] === 3 && out[n - 5] === 0;
      resolve(ok ? { body: out.subarray(head, n - 6), adler: out.readUInt32BE(n - 4) } : null);
    });
    z.write(part);
    // Ended only once flushed: Node folds a flush still queued at the end into the finish.
    z.flush(constants.Z_SYNC_FLUSH, () => z.end());
  });
}

/** One zlib stream compressed in parts side by side, each primed with the 32 KiB before it. */
async function deflateParts(data: Uint8Array, level: number): Promise<Uint8Array> {
  const n = Math.min(MAX_PARTS, Math.ceil(data.length / (PARALLEL_BYTES / 2)));
  const size = Math.ceil(data.length / n);
  const jobs: ReturnType<typeof deflatePart>[] = [];
  for (let start = 0; start < data.length; start += size) {
    jobs.push(deflatePart(data.subarray(start, Math.min(data.length, start + size)), level, start > 0 ? data.subarray(Math.max(0, start - 32768), start) : undefined));
  }
  const parts = await Promise.all(jobs);
  if (parts.some((p) => !p)) return resultBytes(await deflateAsync(data, { level }));
  const out = new Uint8Array(2 + parts.reduce((sum, p) => sum + p!.body.length, 0) + 6);
  // CMF: deflate with a 32 KiB window. FLG: the level's class, with the header's check bits.
  out[0] = 0x78;
  out[1] = level < 0 || level === 6 ? 0x9c : level < 2 ? 0x01 : level < 6 ? 0x5e : 0xda;
  let at = 2;
  let adler = 1;
  for (let i = 0; i < parts.length; i++) {
    const { body, adler: partAdler } = parts[i]!;
    out.set(body, at);
    at += body.length;
    adler = adlerCombine(adler, partAdler, Math.min(size, data.length - i * size));
  }
  // An empty final block, then the checksum of the whole.
  out[at] = 3;
  out[at + 1] = 0;
  new DataView(out.buffer).setUint32(at + 2, adler);
  return out;
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
  deflateZlib: (data, level) => resultBytes(deflateSync(data, level === undefined ? undefined : { level })),
  deflateZlibExactAsync: async (data) => resultBytes(await deflateAsync(data, {})),
  // Runs on libuv's thread pool, so the next page can be drawn meanwhile.
  // zlib.crc32 arrived in Node 20.15 and 22.2.
  crc32: typeof zlib.crc32 === 'function' ? (data) => zlib.crc32(data) : undefined,
  deflateZlibAsync: async (data, level) =>
    data.length >= PARALLEL_BYTES ? deflateParts(data, level ?? constants.Z_DEFAULT_COMPRESSION) : resultBytes(await deflateAsync(data, level === undefined ? {} : { level })),
};
