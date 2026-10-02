# Releasing packages

Seven packages are published under the `@christophervr` npm scope: `@christophervr/docx-core` and one self-contained editor package per framework. Each has its **own version line**, bumped only when it actually changes, and its own git tag `<npm-name>@<version>` (for example `@christophervr/docx-core@0.2.0`). The flow is the one used by [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer), adapted to this repository; the scripts that are copies are marked in their header comments. Like pptx-viewer, only user-facing packages are published: the shared internals are private and bundled into each framework package.

| Package                              | Dir                | Depends on (published)                    | Internal packages bundled in |
| ------------------------------------ | ------------------ | ----------------------------------------- | ---------------------------- |
| `@christophervr/docx-core`           | `packages/core`    | `@christophervr/ooxml-core`               | none                         |
| `@christophervr/docx-react-viewer`   | `packages/react`   | `docx-core`, ProseMirror, `react`         | all five (below)             |
| `@christophervr/docx-vue-viewer`     | `packages/vue`     | `docx-core`, ProseMirror, `vue`           | all five                     |
| `@christophervr/docx-angular-viewer` | `packages/angular` | `docx-core`, ProseMirror, `@angular/core` | all five                     |
| `@christophervr/docx-svelte-viewer`  | `packages/svelte`  | `docx-core`, ProseMirror, `svelte`        | all five                     |
| `@christophervr/docx-solid-viewer`   | `packages/solid`   | `docx-core`, ProseMirror, `solid-js`      | all five                     |
| `@christophervr/docx-vanilla-viewer` | `packages/vanilla` | `docx-core`, ProseMirror                  | all five                     |

The internal packages are `private: true` and are never published: `web-component` and `bindings`. `scripts/build-packages.mjs` bundles them into every framework package, so a published tarball imports only `@christophervr/docx-core`, `@christophervr/ooxml-core`, the ProseMirror libraries and its framework peer. `ole2` is inlined inside `ooxml-core`, so it is not a dependency of anything published and does not have to be released first. The ProseMirror libraries stay real dependencies (not bundled) so a package manager can dedupe one copy of them with other ProseMirror code in the application; `docx-core` stays a dependency (not bundled) so `DocumentModel` is the same module on both sides of the application's own imports.

Dependencies between published packages are read from the manifests, never from a hand-kept list; the directories that are bundled in are the `triggers` of each package in `scripts/release-plan.mjs`. All DOCX logic lives in `@christophervr/ooxml-core` (a separate repository with its own releases); `@christophervr/ole2` is released from its own repository.

## Status: not published yet

None of these packages has ever been published, and their npm trusted publishers are **not configured**, so:

- the scheduled trigger in `.github/workflows/release.yml` is commented out (manual dispatch only), and
- the first release of each package publishes its manifest version (`0.1.0`) as is, with no bump.

Enable the schedule only after the checklist under [One-time setup](#one-time-setup) is complete.

## How a release is decided

Everything is computed by `scripts/release-plan.mjs` from git history; run it locally at any time (it writes the git-ignored `release-plan.json` and changes nothing else):

```sh
bun run release:plan            # or: node scripts/release-plan.mjs --no-npm   (offline)
```

For each package it finds the newest `<npm-name>@x.y.z` tag that is an ancestor of `HEAD` (the baseline) and looks at the files changed since:

- **Own files.** Any non-test file under `packages/<name>/` releases that package. Test files, `__tests__`, `CHANGELOG.md` and manifest edits that only touch `version`, `scripts`, `devDependencies` or ranges on sibling packages never do (so a release commit cannot retrigger a release).
- **Bundled internal packages (triggers).** A non-test change under `packages/web-component` or `bindings` releases **every** framework package, because each one ships that code. A bump of the `@christophervr/ooxml-core` range changes each framework manifest and releases them too.
- **Published dependencies.** A package re-releases whenever a package it depends on is released, because its dependency range changes. Releasing `docx-core` therefore releases all six framework packages; releasing one framework package releases nothing else.
- **Shared build pipeline.** `scripts/build-packages.mjs` and `tsconfig.release.json` change every artifact and release everything.
- **Never published.** A package with no tag that is not on npm releases at its manifest version.

On the first run the plan is seven packages, all `0.1.0` with bump `initial`, and no internal package appears in it:

```text
core           0.1.0 -> 0.1.0 (initial; no previous tag)  tag @christophervr/docx-core@0.1.0
react          0.1.0 -> 0.1.0 (initial; no previous tag)  tag @christophervr/docx-react-viewer@0.1.0
vue            ... angular ... svelte ... solid ... vanilla   (same shape)
```

The bump level is the highest Conventional Commit level among commits since the baseline that touch published files in the package's scope (its own directory, its bundled internal directories and its dependencies'): `!` / `BREAKING CHANGE:` is major, `feat` is minor, anything else is patch. A test-only `feat` does not raise the level. The new version is that bump applied to the highest of the package's tags, the npm `latest` and its manifest version.

`--write` (used by the workflow) stamps the new versions and repoints every sibling dependency range at the version being released, keeping an existing `^` or `~` prefix. The framework packages depend on `@christophervr/docx-core` with a `^` range (`^0.1.0` today); `--write` rewrites it to the exact version being released, so the tarball a consumer installs resolves a `docx-core` released in the same run. The private packages use `workspace:*` and are never rewritten.

## Commit conventions

The bump level comes from the commit type, so conforming commits are enforced. Rules and examples are in [CONTRIBUTING.md](https://github.com/ChristopherVR/docx-viewer/blob/main/CONTRIBUTING.md#commit-conventions). The `PR hygiene / Conventional Commits` workflow validates the PR title and every commit subject (`scripts/check-conventional-commits.mjs`): a missing or unknown type fails; header length, casing and a trailing period only warn. Make that check required in the ruleset.

## The release workflow

`.github/workflows/release.yml` has two jobs and two ways to start.

**Scheduled or manual run without input** (`release` job, then `publish` job):

1. Check out `main` with full history using `RELEASE_TOKEN`, run the planner with `--write`. No changed package means the run ends here; that is what a quiet hour looks like.
2. Fail early if the `NPM_PUBLISH` repository variable is not `true`, so nothing is tagged that cannot be published.
3. `bun install`, `bun run build:packages`, `bun run check:published` (no tarball imports an internal package or `ole2`), `bun run pack:smoke` (installs the seven tarballs into a clean consumer and loads DOCX and legacy `.doc` through every framework package). Any failure stops the run before anything is committed or tagged; the next run retries from the same baseline.
4. Prepend the new section to each released package's `CHANGELOG.md` with [git-cliff](https://git-cliff.org) (`cliff.toml`, scoped to the package's paths and tag pattern) and a dated section to the root `CHANGELOG.md` (`cliff-root.toml`) listing the tags released. Changelogs are prepend-only: old tags are pruned, so history is never regenerated.
5. Commit `chore(release): bump versions and update changelogs [skip ci]` straight to `main` (version bumps, changelogs, `bun.lock`), retrying with a rebase if `main` moved. `[skip ci]` keeps it from starting CI.
6. Tag each released package at that commit through a GitHub release (`scripts/release-notes.mjs` writes the body), upload the plan, and prune superseded releases (`scripts/prune-releases.mjs --keep 1`; tags are kept).
7. `publish` job (environment `npm`): check out the release commit, `bun install --frozen-lockfile`, `bun run build:packages`, then `node scripts/publish-released.mjs --plan release-plan.json`.

**Manual dispatch with `tag`** re-publishes one existing tag (skips the `release` job), for when a publish failed or was skipped:

```sh
gh workflow run release.yml -f tag=@christophervr/docx-core@0.1.0
```

`scripts/publish-released.mjs` publishes in dependency order with `npm publish --provenance --access public`. For each package it first checks that the manifest on disk is the version being published, that no dependency uses `workspace:` or `file:`, that no dependency names an internal package or `@christophervr/ole2`, that sibling ranges match the siblings' versions, and that the version is not already on npm (it is skipped if so, making re-runs safe). A version older than the registry's `latest` is published under the `old` dist-tag. `--dry-run` prints the commands without publishing.

### Authentication: trusted publishing, no secrets

There is no npm token anywhere. The `publish` job requests an OIDC token (`id-token: write`), npm >= 11.5.1 exchanges it for a short-lived publish credential, and `--provenance` attaches the attestation. Nothing else in the repository can publish.

### Why the release job is not gated on a second test run

Like pptx-viewer, the release job does not repeat typecheck, unit and browser tests; it only builds and smoke-tests the packages it will publish. `main` is protected by the `ci-success` check (which `ci.yml` produces by funnelling every required job), so what is on `main` has passed CI. Do not push untested code straight to `main`.

## One-time setup

1. **Rulesets / branch protection on `main`**: require the `ci-success` check and the `Conventional Commits` check; add a bypass for repository admins (the release commit is pushed straight to `main`).
2. **`RELEASE_TOKEN` secret**: a fine-grained personal access token with `Contents: Read and write` on this repository, owned by a repository admin. The default `GITHUB_TOKEN` cannot be granted ruleset bypass, so without it the release commit cannot be pushed.
3. **`NPM_PUBLISH` repository variable** set to `true`.
4. **`npm` environment** (Settings, Environments), optionally with required reviewers.
5. **npm trusted publisher for each of the seven packages** (`@christophervr/docx-core`, `-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer`, `-vanilla-viewer`; none of the internal packages, they are never published) (package settings on npmjs.com, "Trusted Publisher", GitHub Actions): organization/user `ChristopherVR`, repository `docx-viewer`, workflow `release.yml`, environment `npm`. Unverified: npm may only allow configuring a trusted publisher on a package that already exists. If so, the very first version of each package must be bootstrapped once by other means (a short-lived granular token, then revoke it) before the workflow can publish; the planner treats an already-published `0.1.0` as released only once its tag exists, so create the `<npm-name>@0.1.0` tags on the commit that was published.
6. Uncomment the `schedule` trigger in `.github/workflows/release.yml` (hourly), and run the workflow once by hand first.

## Before the first release

- Make sure `@christophervr/ooxml-core@^0.2.0` is on npm; `docx-core` depends on it. `@christophervr/ole2` does not need to be published: its codecs are bundled into the framework packages.
- Check the plan: `bun run release:plan` should list seven packages at `0.1.0` with bump `initial` and no internal package.
- Dispatch the workflow by hand. The seven packages are tagged `@christophervr/docx-core@0.1.0` and `@christophervr/docx-<framework>-viewer@0.1.0`, and published with `docx-core` first.

Local checks before dispatching:

```sh
bun install --frozen-lockfile
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run check:published && bun run pack:smoke
bun run release:plan
node scripts/publish-released.mjs --plan release-plan.json --dry-run
```

## Housekeeping

- `prune-releases.yml` runs daily and keeps only the newest GitHub release per package (tags and npm versions are untouched). Run it manually with `dry_run` to preview.
- `scripts/check-changelog-sections.mjs` (CI) fails if a changelog contains a stub section such as `_Releases: _`, because changelogs are prepend-only and cannot be regenerated.
- Scripts that are candidates to move into a shared package: `release-plan.mjs` (only the package table differs), `check-conventional-commits.mjs`, `check-changelog-sections.mjs`, `release-notes.mjs`, `prune-releases.mjs`, `publish-released.mjs`. They are copied between this repository, `ooxml-core` and `pptx-viewer`.
