# Agent guidance

- Do not commit, amend, push or create branches. Leave changes uncommitted in the working tree;
  the maintainer reviews and commits them.
- Keep private documents out of this repository (see the README's Development section).
- Update [CHANGELOG.md](CHANGELOG.md) only for important user-facing changes or fixes. Keep entries short.
- Before marking work complete, run `pnpm build` and `pnpm test`. The build must succeed and all tests must pass.

## Shared core

kanera-docx, kanera-xlsx and kanera-pptx share a core that must stay identical, byte for byte, in all
three repos: `src/fonts`, `src/image`, `src/ooxml`, `src/pdf`, `src/raster` and `src/util`; the shared
`src/layout` files (all but each package's own engine); the bundled `fonts` (all but the emoji
font); and `scripts/sync-core.mjs`. `node scripts/sync-core.mjs list` prints every core file. The
repos sit side by side (`../kanera-xlsx` and so on); there is no separate core package.

- Run `pnpm core:status` before changing a core file. If another repo has core changes of its own,
  bring them in first by running `pnpm core:push` in that repo. Don't build on a stale copy.
- After changing core files, run `pnpm core:push` to copy them to the other repos, then run
  `pnpm build` and `pnpm test` in each repo it updated. A core change is complete only when all three
  repos build and pass their tests.
- If `core:push` stops on a conflict (both repos changed a file since the last sync), merge the other
  repo's change into this repo's copy, then run `pnpm core:push --force`.
- Keep product code out of the core. Put behaviour that must differ between the packages in
  `src/profile.ts` (each repo's own file, read by the core), or behind an option that the product's
  entry points pass. Extend shared types from product code with module augmentation (see
  kanera-pptx's `src/pptx/model.ts`).
- Changes the core makes in another repo stay uncommitted there like any other change. Commit
  `scripts/core-files.json`, which records the last sync, together with the core files.
