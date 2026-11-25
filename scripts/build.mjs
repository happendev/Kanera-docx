// Bundle browser dependencies into their individual modules. The Node entry bundles
// the conversion path while keeping dynamically imported renderers and codecs lazy.
import { rm } from 'node:fs/promises';
import { build } from 'esbuild';

const common = { bundle: true, format: 'esm', sourcemap: true, sourcesContent: false, outdir: 'dist', logLevel: 'info' };
await build({
  ...common,
  platform: 'browser',
  target: 'es2022',
  entryPoints: ['src/util/pako.ts', 'src/pdf/svg.ts', ...process.argv.slice(2)],
  outbase: 'src',
  minify: true,
});
await rm('dist/node-chunks', { recursive: true, force: true });
await build({
  ...common,
  platform: 'node',
  target: 'node20',
  entryPoints: ['src/node.ts'],
  splitting: true,
  chunkNames: 'node-chunks/[name]-[hash]',
});
