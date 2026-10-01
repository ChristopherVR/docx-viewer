# Contributing to docx-viewer

Read [AGENTS.md](AGENTS.md) for the working agreements (one framework-neutral model, one web-component editor, honest reporting of unsupported features, no logic that belongs in `@christophervr/ooxml-core`).

## Getting set up

You need [Bun](https://bun.sh/) 1.3 and Node.js 24.

```sh
bun install
bun run typecheck
bun run test
bun run test:browser   # Playwright contract tests (bun x playwright install chromium once)
bun run build:packages && bun run pack:smoke
```

Formatting is `oxfmt` (`bun run fmt`, `bun run fmt:check`): tabs, single quotes.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org). This is load-bearing, not cosmetic: each of the seven published packages is versioned and released independently, and **the bump level is derived from your commit type**. A mislabelled commit mis-versions a published package, and a non-conforming commit is silently dropped from the changelog. The `Conventional Commits` check on every pull request validates the PR title and every commit subject (`scripts/check-conventional-commits.mjs`).

```
<type>(<scope>): <subject>

<body>

<footer>
```

- **type**: `feat` (minor bump); `fix`, `perf`, `refactor`, `docs`, `test`, `build`, `ci`, `style`, `chore`, `revert` (patch bump). A `!` after the type/scope, or a `BREAKING CHANGE:` footer, bumps major.
- **scope**: the package or area: `core`, `legacy`, `document`, `layout`, `web-component`, `bindings`, `viewer`, `demo`, `ci`, `deps`, `docs`.
- **subject**: imperative, lower-case, no trailing period, header at most 72 characters.

Which package a commit versions is decided by the **paths it touches**, not by its scope, so keep a commit within one package where practical. A change to a package also re-releases every package that depends on it (see [docs/releasing.md](docs/releasing.md)). Changes that only touch tests, docs or the demos never release anything, whatever their type.

Examples: `feat(web-component): add a table properties dialog`, `fix(legacy): read headers of fast-saved .doc files`, `feat(core)!: rename the section model`.

The release workflow also writes `chore(release): bump versions and update changelogs [skip ci]` commits to `main`; do not edit versions or `CHANGELOG.md` files by hand.

## Pull requests

Keep them focused, describe how you checked the change, and add a regression test for every parsing, preservation, editing or binding change. Do not bump versions in a pull request. CI must be green (`ci-success`).

## License

By contributing you agree that your contributions are licensed under the [Apache License 2.0](LICENSE).
