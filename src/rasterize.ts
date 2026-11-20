import { execFile } from 'node:child_process';
import { copyFile, mkdir, mkdtemp, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

/** An installed command-line PDF rasteriser; tried in this order. */
const RASTERISERS: { name: string; args(pdf: string, dir: string, dpi: number): string[] }[] = [
  { name: 'pdftoppm', args: (pdf, dir, dpi) => ['-png', '-r', String(dpi), pdf, join(dir, 'p')] },
  { name: 'mutool', args: (pdf, dir, dpi) => ['draw', '-q', '-r', String(dpi), '-o', join(dir, 'p-%d.png'), pdf] },
  {
    name: process.platform === 'win32' ? 'gswin64c' : 'gs',
    args: (pdf, dir, dpi) => [
      '-q', '-dNOPAUSE', '-dBATCH', '-dSAFER', '-sDEVICE=png16m', `-r${dpi}`,
      '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4', `-sOutputFile=${join(dir, 'p-%d.png')}`, pdf,
    ],
  },
];

function run(command: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile(command, args, { maxBuffer: 16 * 1024 * 1024 }, (err, _stdout, stderr) => {
      if (!err) return resolve();
      reject(Object.assign(err, { stderr: String(stderr).trim() }));
    });
  });
}

/**
 * Render every page of `pdf` to `<dir>/<prefix>-<n>.png`, with page numbers zero-padded to the
 * width of the page count. Uses pdftoppm (Poppler), mutool (MuPDF) or Ghostscript, whichever is
 * installed first. Returns the written paths in page order.
 */
export async function rasterizePdf(
  pdf: Uint8Array,
  options: { dir: string; prefix: string; dpi: number; pageCount: number },
): Promise<string[]> {
  const { dir, prefix, dpi, pageCount } = options;
  const work = await mkdtemp(join(tmpdir(), 'docx2pdf-raster-'));
  try {
    const input = join(work, 'in.pdf');
    await writeFile(input, pdf);
    const pages = join(work, 'pages');
    let used: string | undefined;
    for (const tool of RASTERISERS) {
      await rm(pages, { recursive: true, force: true });
      await mkdir(pages);
      try {
        await run(tool.name, tool.args(input, pages, dpi));
        used = tool.name;
        break;
      } catch (err) {
        if ((err as NodeJS.ErrnoException).code === 'ENOENT') continue;
        const detail = (err as { stderr?: string }).stderr || (err as Error).message;
        throw new Error(`${tool.name} could not render the pages: ${detail}`);
      }
    }
    if (!used) {
      throw new Error('Rendering page images needs pdftoppm (Poppler), mutool (MuPDF) or Ghostscript on the PATH');
    }
    const rendered = (await readdir(pages))
      .map((file) => ({ file, page: Number(/^p-0*(\d+)\.png$/.exec(file)?.[1]) }))
      .filter(({ page }) => page > 0)
      .sort((a, b) => a.page - b.page);
    if (rendered.length !== pageCount) {
      throw new Error(`${used} rendered ${rendered.length} of ${pageCount} pages`);
    }
    await mkdir(dir, { recursive: true });
    const width = String(pageCount).length;
    const written: string[] = [];
    for (const { file, page } of rendered) {
      const target = join(dir, `${prefix}-${String(page).padStart(width, '0')}.png`);
      await copyFile(join(pages, file), target);
      written.push(target);
    }
    return written;
  } finally {
    await rm(work, { recursive: true, force: true });
  }
}
