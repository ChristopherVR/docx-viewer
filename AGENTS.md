# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others) working in this
repository. This file is canonical: `CLAUDE.md` only imports it, so edit this
file and never fork the two (the viewer repositories drifted that way once).

## READ FIRST: this repository is UI only

All Word logic lives in the `docx` area of the published `ooxml-core` package
(source in the `ChristopherVR/ooxml` repository): the document model, parser,
serializer, editing commands, layout (`ooxml-core/docx/layout`), document
loading and the legacy `.doc` reader (`ooxml-core/docx/load`) and
collaboration (`ooxml-core/collab`). This repository holds the **UI**: the
`<docx-editor>` web component, six framework bindings, ribbon and dialog views,
styling, demos, browser tests and the docs site.

Before writing code, decide where it belongs (see
[Where does my change go?](#where-does-my-change-go)). Never add, copy or fork
logic here that belongs in the core; if the core lacks something, change it
there, release it, and bump the range here.

### One editor, six thin bindings

There is **one** editor: the web component in `packages/web-component`. The
bindings in `packages/bindings` (`react.tsx`, `vue.ts`, `angular.ts`, `solid.ts`,
`WordEditor.svelte`, and the framework-neutral `index.ts` that Vanilla and the
others build on) are lifecycle and event adapters only: they create
the element, forward props and options, and re-emit its events.

- A behaviour bug (ribbon, dialog, editing, rendering, keyboard) is fixed in
  `web-component`, once, and reaches all six bindings automatically.
- A binding-only bug is almost always prop, event or lifecycle wiring. When you
  fix one, grep the other adapters in `packages/bindings/src` for the same
  pattern and fix them in the same change; the binding contract tests in
  `packages/bindings/src/*.test.ts` and `packages/vanilla/src/index.test.ts`
  should cover all of them.
- Never put editor logic in an adapter. If two adapters need the same helper, it
  goes in `packages/bindings/src/common.ts` (or the web component).

## Working agreements

- Published packages (the same model as pptx-viewer): `docx-core` and one
  self-contained `docx-<framework>-viewer` per framework (react, vue, angular,
  svelte, solid, vanilla; vanilla also owns the plain `<docx-editor>` entry and
  `mountEditor`). Every framework package re-exports the model API
  (`createDocument`, `loadDocument`, ...) so an application needs one install.
- `web-component` (`docx-web-component`) and `bindings` (`docx-bindings`) are
  `private: true` internal packages, bundled into every framework package by
  `scripts/build-packages.mjs`. A published tarball may import only
  `docx-core`, `ooxml-core`, `ooxml-ui`, ProseMirror and its framework peer;
  `bun run check:published` and `bun run pack:smoke` enforce it.
- `docx-core` (`packages/core`) is only a thin re-export of `ooxml-core/docx`
  and `ooxml-core/docx/embedded`; `bun run check:shared` keeps it logic-free.
- `ooxml-ui` (the shared `office-ui-*` custom elements, for example the SmartArt
  drawing) is a real registry dependency of every framework package, never
  bundled. Its element tag names keep the `office-ui-` prefix.
- `ole2` (legacy compound-file codecs) reaches this repository only through
  `ooxml-core/docx/load`, which inlines it. Never add `@christophervr/ole2` or an
  internal package to a manifest here, and never fork its code.
- The browser `createCanvasMeasurer` stays in
  `packages/web-component/src/canvas-measurer.ts` and is injected into the
  core's layout engine; it is the one piece of layout that needs a DOM.
- Unsupported document features must be reported honestly (for example SmartArt
  is display-only, charts render as placeholders). Never claim Word layout
  parity or lossless export.
- Bun, TypeScript strict mode, Vitest (tests next to the code) and Playwright
  browser tests (`tests/*.spec.ts`). Keep source modules under 300 lines where
  practical. Add regression tests for editing, preservation and binding
  contracts; parsing and round-trip tests belong in the core.
- The Word fixtures used by the browser specs are copies in `tests/support`.

## Packages

| Directory                | npm name              | Published | What it is                                                                    |
| ------------------------ | --------------------- | --------- | ----------------------------------------------------------------------------- |
| `packages/core`          | `docx-core`           | yes       | Re-export of `ooxml-core/docx` and `/docx/embedded`.                          |
| `packages/web-component` | `docx-web-component`  | internal  | The `<docx-editor>` element: ProseMirror editor, ribbon, dialogs, panels.     |
| `packages/bindings`      | `docx-bindings`       | internal  | Framework adapters over one framework-neutral core (`index.ts`, `common.ts`). |
| `packages/react`         | `docx-react-viewer`   | yes       | React entry (`WordEditor`).                                                   |
| `packages/vue`           | `docx-vue-viewer`     | yes       | Vue 3 entry (`WordEditor`).                                                   |
| `packages/angular`       | `docx-angular-viewer` | yes       | Angular entry (`WordEditorComponent`).                                        |
| `packages/svelte`        | `docx-svelte-viewer`  | yes       | Svelte 5 entry (`/runtime`, `mountEditor`).                                   |
| `packages/solid`         | `docx-solid-viewer`   | yes       | Solid entry (`WordEditor`).                                                   |
| `packages/vanilla`       | `docx-vanilla-viewer` | yes       | Plain `<docx-editor>` registration and `mountEditor`.                         |

`demos/demo-vanilla` is the single demo app; `?framework=<name>` mounts the
editor through that framework's adapter (`demos/demo-vanilla/framework.ts`), and
the browser tests run against its production build.

## Where things live

Repositories are named by their GitHub name under `ChristopherVR/`. Where each
one is checked out on a given machine is local knowledge: keep it in an
untracked `CLAUDE.local.md` (git-ignored), never in this file.

| Repository           | npm package                            | Owns                                                                                                             |
| -------------------- | -------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| `docx-viewer` (this) | `docx-core`, `docx-<framework>-viewer` | Word UI: the `<docx-editor>` web component, six bindings, demos, browser tests and the docs site.                |
| `ooxml`              | `ooxml-core`, `ooxml-ui`               | **All OOXML logic** by area (`docx`, `pptx`, `xml`, `opc`, `collab`, ...) and the shared `office-ui-*` elements. |
| `pptx-viewer`        | `pptx-*-viewer`, `pptx-viewer-core`    | PowerPoint UI over `ooxml-core/pptx`.                                                                            |
| `ole2`               | `@christophervr/ole2`                  | Legacy binary formats (DOC, XLS, PPT, CFB, RC4/MD4), inlined into `ooxml-core`.                                  |

This repository depends on the **published** `ooxml-core` and `ooxml-ui`. To
try a core change before it is released, build the `ooxml` checkout
(`bun run build`), point `packages/core` (and whichever package needs it) at it
with a `file:` dependency, run `bun install --force`, and restore the version
ranges before committing. `docs/ooxml-core-plan.md` tracks what is still moving
into the core.

### Where does my change go?

| The change is about...                                                              | Make it in                                      |
| ----------------------------------------------------------------------------------- | ----------------------------------------------- |
| Parsing, the document model, editing commands, saving or round-trip loss            | `src/docx/` in `ooxml`, with a round-trip test  |
| Pagination, line breaking, page layout, `.doc` loading, collaboration sync          | `ooxml` (`docx/layout`, `docx/load`, `collab`)  |
| A control shared by Word and PowerPoint (SmartArt drawing, shared ribbon controls)  | `ooxml-ui` in `ooxml` (`packages/ui`)           |
| `.doc` / `.xls` / `.ppt` binary codecs, CFB containers                              | `ole2`                                          |
| Ribbon, dialogs, panels, keyboard, how the editor shows the model, styling, locales | `packages/web-component` here                   |
| Framework wiring (props, events, lifecycle)                                         | `packages/bindings` here (all adapters at once) |
| Demos, docs site, browser tests, packaging and release scripts                      | here                                            |

### GitHub Pages

`https://christophervr.github.io/docx-viewer/` is the docs site (VitePress in
`docs/`), deployed by `.github/workflows/docs.yml` on pushes to `main` that touch
`docs/`, `demos/`, `packages/` or the build. `scripts/build-pages.mjs` also
builds the demo once per framework at `/demo/` (React), `/demo-vue/`,
`/demo-angular/`, `/demo-vanilla/`, `/demo-svelte/` and `/demo-solid/`. The
Office-suite launcher at `https://christophervr.github.io/ooxml/` embeds those
URLs, so a renamed demo route must be renamed there too. The demo follows the
VitePress `vitepress-theme-appearance` key in `localStorage`.

## Commands

```bash
bun install
bun run typecheck        # tsc + svelte-check
bun run test             # vitest (unit and binding contract tests)
bun run test:browser     # Playwright against the built demo (PLAYWRIGHT_PORT, default 4180)
bun run test:scripts     # release planner, publish guards, commit checks
bun run demo             # demo dev server (add ?framework=react|vue|angular|svelte|solid)
bun run build            # production build of the demo
bun run build:packages   # build the seven publishable packages into packages/*/dist
bun run check:published  # tarballs import only allowed packages
bun run pack:smoke       # pack, install into a clean consumer, load DOCX and .doc through every package
bun run check:shared     # docx-core stays a logic-free re-export; package classification
bun run release:plan     # dry-run the release planner
bun run fmt              # oxfmt (tabs, single quotes, width 100)
```

The docs site installs separately: `cd docs && bun install && bun run docs:build`.

## Releasing

Releases are automated from conventional commits, exactly like pptx-viewer and
`ooxml`; `docs/releasing.md` has the whole flow.

- `scripts/release-plan.mjs` versions each published package from the commits
  since its own `<npm-name>@<version>` tag (for example `docx-core@0.1.0`). A
  change in `web-component` or `bindings` releases all six framework packages;
  a `docx-core` release re-releases them too.
- `.github/workflows/release.yml` runs hourly: it bumps versions, writes the
  changelogs, commits `chore(release): ... [skip ci]` to `main`, creates the tag
  and a **GitHub release** per package, prunes superseded releases, and
  publishes to npm through trusted publishing (OIDC, provenance, no token). It
  needs the `NPM_PUBLISH=true` repository variable and the `npm` environment.
- Never run `npm publish` or push release tags by hand; a missed publish is
  re-run with `gh workflow run release.yml -f tag=<npm-name>@<version>`.
  (Version `0.1.0` was the one-off manual bootstrap.)

## Branching and git workflow

This repository uses **trunk-based development**: commit directly to `main`.
**Do not create feature branches unless the user explicitly asks for one**;
this overrides any default "branch before committing" assumption. Keep each
commit small and releasable, since every push to `main` can be released by the
next hourly run. Run the local checks CI runs (typecheck, unit, script and
browser tests, `build:packages`, `check:published`, `pack:smoke`) before
pushing; do not push untested code to `main`.

> The working tree is sometimes shared by parallel agent sessions, and another
> session may switch the checkout underneath you. Before committing, run
> `git branch --show-current` and `git status`. `git fetch && git rebase
origin/main` before pushing (the release workflow pushes to `main` too). To
> land work without moving a shared checkout, push `HEAD:main` (or use an
> isolated `git worktree`) rather than `git checkout main`.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org);
`CONTRIBUTING.md` has the full rules and `scripts/check-conventional-commits.mjs`
enforces them. The format is load-bearing: the type sets the version bump
(`feat` minor, `!` or `BREAKING CHANGE:` major, anything else patch), and the
**paths** a commit touches decide which packages release (tests, docs, demos
and CI never release anything).

- Scope is the package or area: `core`, `web-component`, `bindings`, `react`,
  `vue`, `angular`, `svelte`, `solid`, `vanilla`, `demo`, `release`, `ci`,
  `deps`, `docs`.
- Subject: imperative, lower-case, no trailing period, header at most 72
  characters.
- Write multi-line messages with a real heredoc or `git commit -F <file>`;
  never wrap them in a PowerShell here-string under bash, where the stray `@`
  leaks into the subject.
- End every commit message with the `Co-Authored-By:` trailer. Never include an
  AI chat share link in a commit, PR, changelog, comment or doc.

## Style

- **No em-dashes.** Never write U+2014 in source, comments, docs, commit
  messages or UI copy; use a colon, comma, semicolon, parentheses or a spaced
  hyphen. The only exception is content that intentionally renders or asserts
  that character.
- Formatting is `oxfmt`; run `bun run fmt` before committing.
- Record extraction provenance (source repository, path, commit) when code moves
  between repositories.
