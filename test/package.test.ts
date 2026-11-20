import { execFileSync, execSync, spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, it } from 'vitest';
import { COMMON_FONTS, FULL_FONTS } from '../src/fonts/bundled.js';
import { makeDocx, p } from './helpers/docx.js';

it('ships the full font pack for the CLI and exposes the browser subset after an offline install', () => {
  const root = fileURLToPath(new URL('../', import.meta.url));
  const dir = mkdtempSync(join(tmpdir(), 'docx2pdf-package-'));
  try {
    const packed = execSync('npm pack --json', {
      cwd: root,
      encoding: 'utf8',
      env: { ...process.env, npm_config_pack_destination: dir },
    });
    const [tarball] = JSON.parse(packed) as Array<{ filename: string; files: { path: string }[] }>;
    const files = tarball.files.map((file) => file.path);
    expect(files.filter((path) => /\.(ttf|otf|ttc)$/i.test(path)).sort()).toEqual(
      FULL_FONTS.map((font) => `fonts/${font.file}`).sort(),
    );
    expect(files).toEqual(expect.arrayContaining([
      'LICENSE', 'NOTICE', 'fonts/licenses/NotoSansCJK-OFL.txt', 'fonts/licenses/NotoSerifCJK-OFL.txt', 'fonts/licenses/LiberationSansNarrow-EXCEPTION.txt',
    ]));
    renameSync(join(dir, tarball.filename), join(dir, 'library.tgz'));
    writeFileSync(join(dir, 'package.json'), JSON.stringify({ name: 'docx-browser-consumer', private: true, type: 'module' }));
    execSync('npm install --offline --ignore-scripts --no-audit --no-fund ./library.tgz', { cwd: dir });
    const browserFiles = execFileSync(process.execPath, ['--input-type=module', '-e',
      "import { commonFontFiles } from '@happendev/docx2pdf/browser'; console.log(JSON.stringify(commonFontFiles));",
    ], { cwd: dir, encoding: 'utf8' });
    expect(JSON.parse(browserFiles)).toEqual(COMMON_FONTS.map(({ file }) => file));

    const readme = readFileSync(join(root, 'README.md'), 'utf8');
    const example = readme.match(/### Browser\s+```ts\n([\s\S]+?)\n```/);
    expect(example).toBeTruthy();
    writeFileSync(join(dir, 'example.ts'), 'declare const input: { files: FileList };\n' + example![1]);
    writeFileSync(join(dir, 'tsconfig.json'), JSON.stringify({
      compilerOptions: { noEmit: true, strict: true, target: 'ES2022', module: 'ESNext', moduleResolution: 'bundler', customConditions: ['browser'], types: [] },
      files: ['example.ts'],
    }));
    execFileSync(process.execPath, [join(root, 'node_modules/typescript/bin/tsc'), '-p', dir], { stdio: 'pipe' });

    const cli = join(dir, 'node_modules/@happendev/docx2pdf/dist/cli.js');
    const doc = join(dir, 'japanese.docx');
    const pdf = join(dir, 'japanese.pdf');
    writeFileSync(doc, makeDocx({ body: p('日本語') }));
    const full = spawnSync(process.execPath, [cli, doc, pdf], { encoding: 'utf8' });
    expect(full.status).toBe(0);
    expect(full.stderr).toBe('');
    // The text has no East Asian font, so it takes Times New Roman's: a serif, like Word's Mincho fallback.
    expect(readFileSync(pdf, 'latin1')).toContain('NotoSerifCJKSC-Regular');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}, 60_000);
