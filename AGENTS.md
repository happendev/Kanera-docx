# Agent guidance

- Do not commit, amend, push or create branches. Leave changes uncommitted in the working tree;
  the maintainer reviews and commits them.
- Keep private documents out of this repository (see the README's Development section).
- Update [CHANGELOG.md](CHANGELOG.md) only for important user-facing changes or fixes. Keep entries short.
- Before marking work complete, run `pnpm build` and `pnpm test`. The build must succeed and all tests must pass.
