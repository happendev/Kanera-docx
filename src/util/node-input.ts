import { open, type FileHandle } from 'node:fs/promises';
import { checkInputSize, inputLimit, type InputOptions } from './input.js';

type ReadOptions = InputOptions & { signal?: AbortSignal };

export async function readInputPath(path: string, options: ReadOptions = {}): Promise<Uint8Array> {
  options.signal?.throwIfAborted();
  const handle = await open(path, 'r');
  try {
    return await readInputFile(handle, options);
  } finally {
    await handle.close();
  }
}

/** Keep the caller's handle open (the CLI also uses it to check the output file's identity). */
export async function readInputFile(handle: FileHandle, options: ReadOptions = {}): Promise<Uint8Array> {
  const limit = inputLimit(options.maxInputBytes);
  options.signal?.throwIfAborted();
  const stat = await handle.stat();
  const size = stat.size;
  checkInputSize(size, limit);

  // A regular file generally fits one allocation. Its size can change after stat, and pipes can
  // report zero, so keep counting bytes actually read and probe for EOF even at the exact limit.
  const chunks: Buffer[] = [];
  let total = 0;
  // Leave room for the EOF probe in the same buffer when the size is known.
  let capacity = stat.isFile() ? size + 1 : 64 * 1024;
  while (true) {
    options.signal?.throwIfAborted();
    const chunk = Buffer.allocUnsafe(Math.min(capacity, limit - total + 1));
    let filled = 0;
    while (filled < chunk.length) {
      options.signal?.throwIfAborted();
      const length = Math.min(chunk.length - filled, 512 * 1024);
      const { bytesRead } = await handle.read(chunk, filled, length, null);
      if (bytesRead === 0) {
        if (filled) chunks.push(chunk.subarray(0, filled));
        return chunks.length === 1 ? chunks[0] : Buffer.concat(chunks, total);
      }
      total += bytesRead;
      checkInputSize(total, limit);
      filled += bytesRead;
    }
    chunks.push(chunk);
    capacity = 64 * 1024;
  }
}
