/**
 * Builds the browser demo into out/demo (a static site for GitHub Pages).
 *
 *   pnpm demo            # build
 *   pnpm demo --serve    # build, then serve on http://localhost:8000 and rebuild on change
 */
import { context, type BuildOptions } from 'esbuild';
import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Sample } from '../demo/samples.js';
import { FULL_FONTS } from '../src/fonts/bundled.js';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const out = join(root, 'out', 'demo');
const serve = process.argv.includes('--serve');

/** Samples shown first, in this order. */
const FEATURED = [
  'docx_classic120_comprehensive_business_proposal',
  'Fabrikam',
  '13_IEEE_Style_Paper',
  'docx_classic90_comprehensive_annual_report',
  'Invoice',
  'docx_classic75_newsletter_layout',
  'Class News',
  'docx_classic40_resume',
  'docx_classic49_cjk_document',
  'docx_classic57_right_to_left_text',
  'CCU_article',
];

interface Fixture {
  id: string;
  suite: 'classic' | 'issue';
  docx: string;
  pdf?: string;
  reference_application: string;
  reference_images?: { path: string }[];
}

function title(id: string): string {
  const m = /^docx_classic(\d+)_(.*)$/.exec(id);
  const words = (m ? m[2] : id).replace(/[_-]+/g, ' ').trim();
  return words.charAt(0).toUpperCase() + words.slice(1);
}

async function copySamples(): Promise<void> {
  const upstream = join(root, 'benchmarks', 'word', 'upstream');
  const { fixtures } = JSON.parse(await readFile(join(upstream, 'fixtures.json'), 'utf8')) as { fixtures: Fixture[] };
  const samples: Sample[] = [];
  // The classic reference images have no redistribution license. Keep them in the local preview.
  for (const f of fixtures.filter((x) => /\bWord\b/.test(x.reference_application) && (serve || x.suite === 'issue'))) {
    const dir = `samples/${f.suite}`;
    const docx = `${dir}/${basename(f.docx)}`;
    await cp(join(upstream, f.docx), join(out, docx));
    const refs = f.pdf ? [f.pdf] : (f.reference_images ?? []).map((r) => r.path);
    const files: string[] = [];
    for (const r of refs) {
      const dest = `${dir}/reference/${basename(r)}`;
      await cp(join(upstream, r), join(out, dest));
      files.push(dest);
    }
    samples.push({
      id: f.id,
      title: title(f.id),
      suite: f.suite,
      featured: FEATURED.includes(f.id),
      docx,
      reference: { kind: f.pdf ? 'pdf' : 'png', app: f.reference_application, files },
    });
  }
  const rank = (s: Sample) => (s.featured ? FEATURED.indexOf(s.id) : FEATURED.length);
  const num = (s: Sample) => Number(/^docx_classic(\d+)/.exec(s.id)?.[1] ?? 0);
  samples.sort((a, b) => rank(a) - rank(b) || a.suite.localeCompare(b.suite) || num(a) - num(b) || a.title.localeCompare(b.title));
  await writeFile(join(out, 'samples', 'samples.json'), JSON.stringify(samples));
}

async function copyStatic(): Promise<void> {
  await cp(join(root, 'demo', 'index.html'), join(out, 'index.html'));
  await cp(join(root, 'demo', 'style.css'), join(out, 'style.css'));
  await cp(join(root, 'LICENSE'), join(out, 'LICENSE.txt'));
  await cp(join(root, 'NOTICE'), join(out, 'NOTICE.txt'));
  for (const f of ['kanera-logo.svg', 'kanera-icon.svg']) await cp(join(root, 'demo', f), join(out, f));
  // Kanera's typefaces (OFL), self-hosted so the page loads nothing from third parties.
  const faces = [
    ['geist-sans', 'geist-sans-latin-400-normal.woff2'],
    ['geist-sans', 'geist-sans-latin-500-normal.woff2'],
    ['geist-sans', 'geist-sans-latin-600-normal.woff2'],
    ['instrument-serif', 'instrument-serif-latin-400-normal.woff2'],
    ['instrument-serif', 'instrument-serif-latin-400-italic.woff2'],
    ['jetbrains-mono', 'jetbrains-mono-latin-500-normal.woff2'],
  ];
  for (const [pkg, file] of faces) await cp(join(root, 'node_modules', '@fontsource', pkg, 'files', file), join(out, 'ui-fonts', file));
  for (const pkg of new Set(faces.map(([pkg]) => pkg))) {
    await cp(join(root, 'node_modules', '@fontsource', pkg, 'LICENSE'), join(out, 'ui-fonts', `${pkg}-LICENSE.txt`));
  }
  await writeFile(join(out, '.nojekyll'), '');
}

async function main(): Promise<void> {
  await rm(out, { recursive: true, force: true });
  await mkdir(out, { recursive: true });

  // The library's full font pack and its licenses; the browser build fetches faces on demand.
  for (const { file } of FULL_FONTS) await cp(join(root, 'fonts', file), join(out, 'fonts', file));
  await cp(join(root, 'fonts', 'licenses'), join(out, 'fonts', 'licenses'), { recursive: true });
  // pdf.js renders the pages for the preview.
  const pdfjs = join(root, 'node_modules', 'pdfjs-dist');
  await cp(join(pdfjs, 'build', 'pdf.worker.min.mjs'), join(out, 'assets', 'pdf.worker.min.mjs'));
  await cp(join(pdfjs, 'cmaps'), join(out, 'assets', 'pdfjs', 'cmaps'), { recursive: true });
  await cp(join(pdfjs, 'standard_fonts'), join(out, 'assets', 'pdfjs', 'standard_fonts'), { recursive: true });
  await cp(join(pdfjs, 'wasm'), join(out, 'assets', 'pdfjs', 'wasm'), { recursive: true });
  await cp(join(pdfjs, 'LICENSE'), join(out, 'assets', 'pdfjs', 'LICENSE.txt'));
  await copySamples();
  await copyStatic();

  const options: BuildOptions = {
    entryPoints: { main: join(root, 'demo', 'main.ts'), 'convert.worker': join(root, 'demo', 'convert.worker.ts') },
    outdir: join(out, 'assets'),
    bundle: true,
    splitting: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    minify: !serve,
    sourcemap: true,
    // As in the package build: jpeg-js reaches for Buffer.from.
    define: { 'Buffer.from': 'Uint8Array.from' },
    logLevel: 'info',
  };
  const ctx = await context(options);
  if (serve) {
    await ctx.watch();
    const { port } = await ctx.serve({ servedir: out, port: 8000 });
    console.log(`Demo at http://localhost:${port}/ (HTML and CSS changes need a restart)`);
  } else {
    await ctx.rebuild();
    await ctx.dispose();
    console.log(`Demo built in ${out}`);
  }
}

await main();
