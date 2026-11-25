#!/usr/bin/env node
// Keeps the shared core of kanera-docx, kanera-xlsx and kanera-pptx identical. Each repo carries
// its own copy of the core files (CORE below; this script is one of them), so there is no separate
// package to publish. Run it from any of the three:
//
//   node scripts/sync-core.mjs [status]  lists core files that differ between the repos (exit 1 if any)
//   node scripts/sync-core.mjs list      lists the core files
//   node scripts/sync-core.mjs push      copies this repo's core files to the other repos
//     --dry-run  report what push would do without writing anything
//     --force    also overwrite (or delete) files another repo has changed since the last sync
//
// The other repos are found beside this one (../kanera-docx, ../kanera-xlsx, ../kanera-pptx), or
// listed in KANERA_REPOS, separated by the platform's path delimiter. A repo that isn't there is
// skipped.
//
// scripts/core-files.json holds every core file's hash as of the last push, so the script can tell
// a repo that is only behind from one that has changed a file of its own. Push brings repos that are
// behind up to date, leaves another repo's own changes alone where this repo hasn't changed the
// file, and stops without writing anything if both have changed it (merge the two here, then push
// with --force).

import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { basename, delimiter, dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

/** The core: whole directories (recursively), single files, and paths left out of them. */
const CORE = {
  dirs: ['src/fonts', 'src/image', 'src/ooxml', 'src/pdf', 'src/raster', 'src/util', 'fonts'],
  files: [
    'src/layout/bidi.ts',
    'src/layout/box.ts',
    'src/layout/chart.ts',
    'src/layout/drawing.ts',
    'src/layout/gradient.ts',
    'src/layout/math.ts',
    'src/layout/metafile.ts',
    'src/layout/paragraph.ts',
    'src/layout/shadow.ts',
    'src/layout/shaper.ts',
    'src/layout/types.ts',
    'src/layout/vertical.ts',
    'scripts/build.mjs',
    'scripts/sync-core.mjs',
  ],
  // Each package bundles one emoji font, which its src/profile.ts names.
  exclude: ['fonts/Noto-COLRv1.ttf', 'fonts/NotoEmoji-VariableFont_wght.ttf', 'fonts/licenses/NotoColorEmoji-OFL.txt', 'fonts/licenses/NotoEmoji-OFL.txt'],
};

const REPOS = ['kanera-docx', 'kanera-xlsx', 'kanera-pptx'];
const LOCK = 'scripts/core-files.json';

const here = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const command = args.find((a) => !a.startsWith('--')) ?? 'status';
const dryRun = args.includes('--dry-run');
const force = args.includes('--force');

/** Core files in a repo, by path relative to it (with / separators). */
function coreFiles(root) {
  const out = new Map();
  const add = (rel) => {
    if (!CORE.exclude.includes(rel)) out.set(rel, join(root, ...rel.split('/')));
  };
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const full = join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (entry.isFile()) add(relative(root, full).split(sep).join('/'));
    }
  };
  for (const dir of CORE.dirs) if (existsSync(join(root, dir))) walk(join(root, dir));
  for (const file of CORE.files) if (existsSync(join(root, file)) && statSync(join(root, file)).isFile()) add(file);
  return out;
}

const hashOf = (file) => createHash('sha256').update(readFileSync(file)).digest('hex').slice(0, 16);

function readLock(root) {
  const file = join(root, LOCK);
  if (!existsSync(file)) return undefined;
  return new Map(Object.entries(JSON.parse(readFileSync(file, 'utf8')).files ?? {}));
}

function writeLock(root, hashes) {
  const files = Object.fromEntries([...hashes].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)));
  const body = { note: 'Written by scripts/sync-core.mjs push: each core file\'s hash at the last sync. Commit it with the core files.', files };
  writeFileSync(join(root, LOCK), `${JSON.stringify(body, null, 2)}\n`);
}

/** This repo first, then the others that are present. */
function findRepos() {
  const listed = process.env.KANERA_REPOS?.split(delimiter).filter(Boolean).map((p) => resolve(p));
  const candidates = listed ?? REPOS.map((name) => resolve(here, '..', name));
  const repos = [{ name: basename(here), root: here }];
  const missing = [];
  for (const root of candidates) {
    if (root === here) continue;
    if (existsSync(join(root, 'package.json'))) repos.push({ name: basename(root), root });
    else missing.push(root);
  }
  for (const repo of repos) {
    repo.files = coreFiles(repo.root);
    repo.hashes = new Map([...repo.files].map(([rel, file]) => [rel, hashOf(file)]));
    repo.lock = readLock(repo.root);
  }
  return { repos, missing };
}

function status() {
  const { repos, missing } = findRepos();
  for (const root of missing) console.log(`Skipped ${root}: not found.`);
  const paths = [...new Set(repos.flatMap((r) => [...r.hashes.keys()]))].sort();
  const drift = [];
  for (const rel of paths) {
    const hashes = repos.map((r) => r.hashes.get(rel));
    if (hashes.every((h) => h !== undefined && h === hashes[0])) continue;
    const states = repos.map((r, i) => {
      const h = hashes[i];
      const base = r.lock?.get(rel);
      if (h === undefined) return `${r.name}: missing`;
      if (h === base) return `${r.name}: as synced`;
      return `${r.name}: ${base === undefined ? 'new' : 'changed'}`;
    });
    const changed = repos.filter((r, i) => hashes[i] !== undefined && hashes[i] !== r.lock?.get(rel));
    const conflict = new Set(changed.map((r) => r.hashes.get(rel))).size > 1;
    drift.push(`  ${rel}${conflict ? '  CONFLICT' : ''}\n      ${states.join(' · ')}`);
  }
  const names = repos.map((r) => r.name).join(', ');
  if (!drift.length) {
    console.log(`${paths.length} core files are identical in ${names}.`);
    return 0;
  }
  console.log(`${drift.length} of ${paths.length} core files differ between ${names}:\n${drift.join('\n')}`);
  console.log('\nPush from the repo that has the version to keep (pnpm core:push there); merge CONFLICT files by hand first.');
  return 1;
}

function push() {
  const { repos, missing } = findRepos();
  const [source, ...targets] = repos;
  for (const root of missing) console.log(`Skipped ${root}: not found.`);
  if (!targets.length) {
    console.log('No other Kanera repos found; nothing to push to.');
    return 1;
  }
  const plans = [];
  const conflicts = [];
  for (const target of targets) {
    const plan = { target, copy: [], remove: [], kept: [] };
    const paths = new Set([...source.hashes.keys(), ...target.hashes.keys(), ...(source.lock?.keys() ?? [])]);
    for (const rel of [...paths].sort()) {
      const s = source.hashes.get(rel);
      const t = target.hashes.get(rel);
      if (s === t) continue;
      // A file the target has changed since the last sync is its own work: keep it, or stop if this
      // repo has changed the file too.
      const targetChanged = t !== target.lock?.get(rel);
      const sourceChanged = s !== source.lock?.get(rel);
      if (targetChanged && !force) {
        if (sourceChanged) conflicts.push(`  ${rel}: changed in both ${source.name} and ${target.name}`);
        else plan.kept.push(rel);
        continue;
      }
      if (s === undefined) plan.remove.push(rel);
      else plan.copy.push(rel);
    }
    plans.push(plan);
  }
  if (conflicts.length) {
    console.log(`Nothing pushed: ${conflicts.length} file(s) changed on both sides since the last sync.\n${conflicts.join('\n')}`);
    console.log(`\nMerge the other repo's changes into ${source.name}'s copy, then run push --force.`);
    return 1;
  }
  for (const { target, copy, remove, kept } of plans) {
    const verb = dryRun ? 'Would update' : 'Updated';
    console.log(`${verb} ${target.name}: ${copy.length} copied, ${remove.length} removed${kept.length ? `, ${kept.length} of its own changes kept` : ''}.`);
    for (const rel of copy) console.log(`  + ${rel}`);
    for (const rel of remove) console.log(`  - ${rel}`);
    for (const rel of kept) console.log(`  = ${rel} (changed in ${target.name}; push from there)`);
    if (dryRun) continue;
    for (const rel of copy) {
      const to = join(target.root, ...rel.split('/'));
      mkdirSync(dirname(to), { recursive: true });
      copyFileSync(source.files.get(rel), to);
    }
    for (const rel of remove) rmSync(join(target.root, ...rel.split('/')));
  }
  if (dryRun) return 0;
  // The source's files are now the synced state everywhere; a kept file still differs from it, so
  // it shows as that repo's change.
  for (const repo of repos) writeLock(repo.root, source.hashes);
  console.log(`\nBuild and test the updated repos: ${targets.map((t) => `pnpm -C ${relative(here, t.root) || '.'} build && pnpm -C ${relative(here, t.root) || '.'} test`).join('; ')}`);
  return 0;
}

const commands = {
  status,
  push,
  list: () => {
    for (const rel of [...coreFiles(here).keys()].sort()) console.log(rel);
    return 0;
  },
};
if (commands[command]) process.exitCode = commands[command]();
else {
  console.error(`Unknown command "${command}". Use status, list or push.`);
  process.exitCode = 2;
}
