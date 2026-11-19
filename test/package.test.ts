import { execFileSync, execSync } from 'node:child_process';
import { mkdtempSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, it } from 'vitest';

it('ships the font licenses and compiles the documented browser example in an offline consumer', () => {
  const root = fileURLToPath(new URL('../', import.meta.url));
  const dir = mkdtempSync(join(tmpdir(), 'docx2pdf-package-'));
  try {
    const packed = execSync('npm pack --json', {
      cwd: root,
      encoding: 'utf8',
      env: { ...process.env, npm_config_pack_destination: dir },
    });
    const [tarball] = JSON.parse(packed) as Array<{ filename: string; files: { path: string }[] }>;
    expect(tarball.files.map((file) => file.path)).toEqual(expect.arrayContaining([
      'LICENSE', 'NOTICE', 'fonts/NotoSansCJKjp-Regular.ttf', 'fonts/licenses/NotoSansCJK-OFL.txt',
    ]));
    renameSync(join(dir, tarball.filename), join(dir, 'library.tgz'));
    writeFileSync(join(dir, 'package.json'), JSON.stringify({ name: 'docx-browser-consumer', private: true, type: 'module' }));
    execSync('npm install --offline --ignore-scripts --no-audit --no-fund ./library.tgz', { cwd: dir });
    const readme = readFileSync(join(root, 'README.md'), 'utf8');
    const example = readme.match(/### Browser\s+```ts\n([\s\S]+?)\n```/);
    expect(example).toBeTruthy();
    writeFileSync(join(dir, 'example.ts'), 'declare const input: { files: FileList };\n' + example![1]);
    writeFileSync(join(dir, 'tsconfig.json'), JSON.stringify({
      compilerOptions: { noEmit: true, strict: true, target: 'ES2022', module: 'ESNext', moduleResolution: 'bundler', customConditions: ['browser'], types: [] },
      files: ['example.ts'],
    }));
    execFileSync(process.execPath, [join(root, 'node_modules/typescript/bin/tsc'), '-p', dir], { stdio: 'pipe' });
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}, 60_000);
