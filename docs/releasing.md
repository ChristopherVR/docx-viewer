# Releasing packages

Seven packages are published under the `@christophervr` npm scope. The umbrella package is `@christophervr/docx-viewer`; installing it brings in the rest. Each package has its **own version line**, bumped only when it actually changes, and its own git tag `<npm-name>@<version>` (for example `@christophervr/docx-core@0.2.0`). The flow is the one used by [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer), adapted to this repository; the scripts that are copies are marked in their header comments.

| Package                             | Depends on (inside this repo)                   |
| ----------------------------------- | ----------------------------------------------- |
| `@christophervr/docx-core`          | none (consumes `@christophervr/ooxml-core`)     |
| `@christophervr/docx-legacy`        | core                                            |
| `@christophervr/docx-document`      | core, legacy                                    |
| `@christophervr/docx-layout`        | core                                            |
| `@christophervr/docx-web-component` | core, document, layout                          |
| `@christophervr/docx-bindings`      | core, web-component                             |
| `@christophervr/docx-viewer`        | bindings, core, document, legacy, web-component |

Dependencies are read from the manifests, never from a hand-kept list. All DOCX logic lives in `@christophervr/ooxml-core` (a separate repository with its own releases); `@christophervr/ole2` is released from its own repository.

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
- **Internal dependencies.** A package re-releases whenever a package it depends on is released, transitively, because its dependency range changes. Releasing `core` therefore releases all seven; releasing `layout` releases `layout`, `web-component`, `bindings` and `viewer`.
- **Shared build pipeline.** `scripts/build-packages.mjs` and `tsconfig.release.json` change every artifact and release everything.
- **Never published.** A package with no tag that is not on npm releases at its manifest version.

The bump level is the highest Conventional Commit level among commits since the baseline that touch published files in the package's scope (its own directory plus its dependencies'): `!` / `BREAKING CHANGE:` is major, `feat` is minor, anything else is patch. A test-only `feat` does not raise the level. The new version is that bump applied to the highest of the package's tags, the npm `latest` and its manifest version.

`--write` (used by the workflow) stamps the new versions and repoints every sibling dependency range at the version being released, keeping an existing `^` or `~` prefix. Sibling ranges are exact pins today, so a consumer always installs one consistent set.

## Commit conventions

The bump level comes from the commit type, so conforming commits are enforced. Rules and examples are in [CONTRIBUTING.md](https://github.com/ChristopherVR/docx-viewer/blob/main/CONTRIBUTING.md#commit-conventions). The `PR hygiene / Conventional Commits` workflow validates the PR title and every commit subject (`scripts/check-conventional-commits.mjs`): a missing or unknown type fails; header length, casing and a trailing period only warn. Make that check required in the ruleset.

## The release workflow

`.github/workflows/release.yml` has two jobs and two ways to start.

**Scheduled or manual run without input** (`release` job, then `publish` job):

1. Check out `main` with full history using `RELEASE_TOKEN`, run the planner with `--write`. No changed package means the run ends here; that is what a quiet hour looks like.
2. Fail early if the `NPM_PUBLISH` repository variable is not `true`, so nothing is tagged that cannot be published.
3. `bun install`, `bun run build:packages`, `bun run pack:smoke`. Any failure stops the run before anything is committed or tagged; the next run retries from the same baseline.
4. Prepend the new section to each released package's `CHANGELOG.md` with [git-cliff](https://git-cliff.org) (`cliff.toml`, scoped to the package's paths and tag pattern) and a dated section to the root `CHANGELOG.md` (`cliff-root.toml`) listing the tags released. Changelogs are prepend-only: old tags are pruned, so history is never regenerated.
5. Commit `chore(release): bump versions and update changelogs [skip ci]` straight to `main` (version bumps, changelogs, `bun.lock`), retrying with a rebase if `main` moved. `[skip ci]` keeps it from starting CI.
6. Tag each released package at that commit through a GitHub release (`scripts/release-notes.mjs` writes the body), upload the plan, and prune superseded releases (`scripts/prune-releases.mjs --keep 1`; tags are kept).
7. `publish` job (environment `npm`): check out the release commit, `bun install --frozen-lockfile`, `bun run build:packages`, then `node scripts/publish-released.mjs --plan release-plan.json`.

**Manual dispatch with `tag`** re-publishes one existing tag (skips the `release` job), for when a publish failed or was skipped:

```sh
gh workflow run release.yml -f tag=@christophervr/docx-core@0.1.0
```

`scripts/publish-released.mjs` publishes in dependency order with `npm publish --provenance --access public`. For each package it first checks that the manifest on disk is the version being published, that no dependency uses `workspace:` or `file:`, that sibling ranges match the siblings' versions, and that the version is not already on npm (it is skipped if so, making re-runs safe). A version older than the registry's `latest` is published under the `old` dist-tag. `--dry-run` prints the commands without publishing.

### Authentication: trusted publishing, no secrets

There is no npm token anywhere. The `publish` job requests an OIDC token (`id-token: write`), npm >= 11.5.1 exchanges it for a short-lived publish credential, and `--provenance` attaches the attestation. Nothing else in the repository can publish.

### Why the release job is not gated on a second test run

Like pptx-viewer, the release job does not repeat typecheck, unit and browser tests; it only builds and smoke-tests the packages it will publish. `main` is protected by the `ci-success` check (which `ci.yml` produces by funnelling every required job), so what is on `main` has passed CI. Do not push untested code straight to `main`.

## One-time setup

1. **Rulesets / branch protection on `main`**: require the `ci-success` check and the `Conventional Commits` check; add a bypass for repository admins (the release commit is pushed straight to `main`).
2. **`RELEASE_TOKEN` secret**: a fine-grained personal access token with `Contents: Read and write` on this repository, owned by a repository admin. The default `GITHUB_TOKEN` cannot be granted ruleset bypass, so without it the release commit cannot be pushed.
3. **`NPM_PUBLISH` repository variable** set to `true`.
4. **`npm` environment** (Settings, Environments), optionally with required reviewers.
5. **npm trusted publisher for each of the seven packages** (package settings on npmjs.com, "Trusted Publisher", GitHub Actions): organization/user `ChristopherVR`, repository `docx-viewer`, workflow `release.yml`, environment `npm`. Unverified: npm may only allow configuring a trusted publisher on a package that already exists. If so, the very first version of each package must be bootstrapped once by other means (a short-lived granular token, then revoke it) before the workflow can publish; the planner treats an already-published `0.1.0` as released only once its tag exists, so create the `<npm-name>@0.1.0` tags on the commit that was published.
6. Uncomment the `schedule` trigger in `.github/workflows/release.yml` (hourly), and run the workflow once by hand first.

## Before the first release

- Publish `@christophervr/ole2` (the sibling repository) and make sure `@christophervr/ooxml-core@^0.1.0` is on npm; the legacy and core packages depend on them.
- Check the plan: `bun run release:plan` should list seven packages at `0.1.0` with bump `initial`.
- Dispatch the workflow by hand. The seven packages are tagged `@christophervr/docx-<name>@0.1.0` and published in dependency order.

Local checks before dispatching:

```sh
bun install --frozen-lockfile
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run pack:smoke
node scripts/publish-released.mjs --plan release-plan.json --dry-run
```

## Housekeeping

- `prune-releases.yml` runs daily and keeps only the newest GitHub release per package (tags and npm versions are untouched). Run it manually with `dry_run` to preview.
- `scripts/check-changelog-sections.mjs` (CI) fails if a changelog contains a stub section such as `_Releases: _`, because changelogs are prepend-only and cannot be regenerated.
- Scripts that are candidates to move into a shared package: `release-plan.mjs` (only the package table differs), `check-conventional-commits.mjs`, `check-changelog-sections.mjs`, `release-notes.mjs`, `prune-releases.mjs`, `publish-released.mjs`. They are copied between this repository, `ooxml-core` and `pptx-viewer`.
