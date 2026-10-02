# Handover: Office suite programme (docx-viewer / ooxml-core / pptx-viewer)

Uncommitted on purpose (delete or commit as you see fit). Written 2026-10-02 at the end of a long session. The user's own words are paraphrased only where noted; ask them when in doubt.

## 1. Goal and decisions already made

A unified Office suite. All logic lives in **one published package**, `@christophervr/ooxml-core`, with an area per subpath (`xml`, `opc`, `units`, `color`, `geometry`, `diagram`, `docx`, `docx/layout`, `docx/load`, `pptx`, `collab`; later `drawingml`, `chart`, xlsx, Visio `.vsdx`). Docx and pptx are **equal, symmetrical areas**. ole2 (`D:\Development\ole2`) stays legacy-binary only and is inlined into the core's bundles, never forked.

The viewer repos hold **only UI**. Shared UI web components that all Office products need live in a second package in the core repo: `@christophervr/office-ui` (`packages/ui`, `office-ui-*` tags). It is a normal dependency of each editor package, so **users never install it separately**; same for `docx-core`. Yjs/collab is part of the core (`@christophervr/ooxml-core/collab`).

Working rules from the user: **trunk-based on `main`, no PRs, no long-lived feature branches** (delete a branch right after it lands), merge once verified locally, prefer one CI run on `main`. Never claim Word layout parity or lossless export (docx-viewer AGENTS.md). No unnecessary questions: recommend, decide.

## 2. Repos, state and URLs

| Repo | Local path | `main` at handover | Open issues |
|---|---|---|---|
| ooxml-core (public) | `D:\Development\ooxml-core` | `14df8c8` release commit + later docs | 0 |
| docx-viewer | `D:\Development\docx-viewer` (this checkout) | `d99f3ef` | 0 |
| pptx-viewer | `D:\Development\pptx-wt\pptx-viewer` (main worktree; many `w*` worktrees exist) | `d7d4042be` (release bot commits land here too) | #342 #373 #386 #396 #402 |

npm (verified): `@christophervr/ooxml-core@0.2.0` and `@christophervr/office-ui@0.1.1` are **published with provenance via OIDC**. `office-ui@0.1.0` is a bad version (core `^0.1.0` range: no `/diagram` entry): deprecate it (needs the user's OTP): `npm deprecate @christophervr/office-ui@0.1.0 "needs @christophervr/ooxml-core ^0.2.0; use 0.1.1 or later" --otp=<code>`.

docx-viewer packages are NOT published yet (see section 6).

## 3. How to work in this environment (read this first)

* **Windows + Git Bash.** Heredocs collapse backslashes (use Write/Edit for backslash-heavy files); CRLF warnings are noise; Git-Bash `tar` breaks some local pack scripts (docx `pack:smoke` works locally; pptx `check-*-package.mjs` binding script hits the `C:` tar bug locally, CI is fine).
* **Pushing to main.** The user added permission rules, and the pptx ruleset has an admin bypass: `git push origin HEAD:main` works from the main session. **Subagents are still refused by the auto-mode classifier** (reasons "Git Destructive"/"Merge Without Review"/"CI Bypass"): they push a branch, you rebase + land it from your session, then delete the branch (`git push origin --delete <branch>`). Workflow dispatch, `gh variable set`, and `gh run rerun` of in-progress runs may be blocked: give the user `! gh ...` commands to run.
* **Always `git fetch && git rebase origin/main` before pushing.** pptx-viewer's hourly release automation and the ooxml-core release workflow push `chore(release): ... [skip ci]` commits to `main`.
* **Local gate before every push to pptx-viewer main** (full e2e is ~4,170 tests, 1.5 h; CI runs once on main afterwards): `bun install`, `bun run build`, `bun run typecheck`, `oxlint --deny-warnings` on changed files, all unit suites, plus the e2e specs for the touched area on all five bindings. A reusable script is saved next to this file: `HANDOVER-gate-pptx.sh <worktree>` (typecheck, lint, parallel unit suites; ends with `GATE_EXIT=n`).
* **Machine is slow when several agents run.** Use `--workers=2` and a unique `PPTX_E2E_PORT_OFFSET` (docx: `PLAYWRIGHT_PORT`). Known load-only failures (pass alone, also fail on main under load): shared `deck-save-encryption.test.ts` and `modify-password-check.test.ts` (PBKDF timeouts), `ssr.test.ts`, occasionally `collaboration-component.test.ts`. Do NOT dismiss a failure as load without re-running it alone and comparing against `origin/main`; I did once and it was a real stale test (fixed).
* **e2e hygiene.** After any Playwright run: `git checkout -- e2e/fixtures` (it regenerates ~33 tracked decks). If e2e says "Stale build output": `bun run build`; if it names the Angular demo: `rm -rf demos/demo-angular/node_modules/.vite`. After pulling a new core version run `bun install` (docx-viewer needs it for `@christophervr/ooxml-core@0.2.0`).
* **pptx CI.** `ci.yml` has `cancel-in-progress` on pushes: rapid pushes to `main` cancel each other, so `main` currently has **no completed green run for the latest heads**. Ask the user to dispatch one: `! gh workflow run ci.yml -R ChristopherVR/pptx-viewer`. CI now uses **10 e2e shards** and a 120 s budget for dense-panel specs (these were genuine capacity timeouts).
* Commit conventions are enforced (`Conventional Commits` check, subject <= 72 chars, types feat/fix/perf/refactor/docs/test/build/ci/style/chore/revert). Every commit ends with the `Co-Authored-By` + `Claude-Session` trailers.

## 4. Release / publishing machinery (all three repos match pptx-viewer's model)

* Conventional-commit driven planner `scripts/release-plan.mjs`, `publish-released.mjs`, git-cliff changelogs, hourly `release.yml` (+ manual dispatch, tag re-publish mode), `ci-success` aggregate check, OIDC trusted publishing (no npm tokens), release pruning.
* **ooxml-core**: schedule ON, repo variable `NPM_PUBLISH=true` set, `npm` environment exists, no ruleset, trusted publishers configured for both packages. The ui's core dependency is declared `"*"` (root cannot be its own workspace member) and rewritten to `^<core version>` only in the published manifest (`isWorkspaceRange` covers `*`). `publish-released.mjs --manual` exists for hand publishes without provenance. Note: new core versions take ~8 minutes to become visible on the registry; the ui publish guard checks it.
* **docx-viewer**: workflow written, **cron commented out** until the user confirms. Needs per-package trusted publishers (below).
* **pptx-viewer**: pre-existing hourly release; it has been publishing packages from `main`.

## 5. What was completed this session (high level)

* ooxml-core: symmetric docx/pptx (ESM+CJS), fixtures committed in-repo (no pptx-viewer clone in CI), parallel build (`scripts/build.mjs`, ~52 s), release automation, new areas `diagram` (format-neutral SmartArt model; docx SmartArt parse/preserve), `docx/layout`, `docx/load`, `collab` (Yjs), plus `packages/ui` (`office-ui`). Plans: `docs/agnostic-core-plan.md`, `docs/office-ui-plan.md`, `docs/collab-area.md`, `docs/docx-layout-adoption.md`.
* docx-viewer: published set restructured like pptx: `@christophervr/docx-core` + one self-contained package per framework (`docx-react-viewer`, `-vue-`, `-angular-`, `-svelte-`, `-solid-`, `-vanilla-viewer`); private internals `web-component`, `bindings`. Adopted core layout/load/collab (private layout/document/legacy packages deleted); SmartArt now **displays** from the cached drawing (read-only, via `office-ui-smartart`), honestly labelled. Each framework package re-exports the model API (`createDocument` etc.).
* pptx-viewer: engine is `@christophervr/ooxml-core/pptx` (`pptx-viewer-core` is a thin re-export, `^0.1.0` -> bump to `^0.2.0` when convenient). Shared web-component migration landed for Design/Review/contextual, Draw, View, Transitions, Animations, Insert, Home (all strips), status bar, title bar, notes toolbar, context menus, banner/paste/toasts/mobile/slide-show toolbar/dialog footer (4 dialogs), rail/section/sorter menus, inspector reset actions, control tokens + shared select/checkbox/search everywhere (native selects/checkboxes gone, drift tests guard it), group drill-in lock fix, Svelte inline-editor/bootstrap race fix, SmartArt regression tests.
* Issues closed: #374 #375 #377 #378 #380 #387 #393 #394 #395 #397 #398.

## 6. Pending user actions (cannot be done by the assistant)

1. npm: add **trusted publishers** for the 7 docx packages (`@christophervr/docx-core`, `docx-{react,vue,angular,svelte,solid,vanilla}-viewer`: repo `ChristopherVR/docx-viewer`, workflow `release.yml`, environment `npm`); first publish of each may need a one-off manual publish (npm may require the package to exist first). Then uncomment the cron in docx-viewer's `release.yml` and set `NPM_PUBLISH=true` + create the `npm` environment there.
2. `npm deprecate` of `office-ui@0.1.0` (OTP).
3. **Rotate the npm token** that was pasted into the chat earlier.
4. Dispatch a full CI run on pptx-viewer `main` (see section 3).
5. (Done in this session's last step: stale local branches/worktrees deleted, see section 8.)
6. SmartArt report ("can't double-click to inline edit / can't see colour options"): **not reproducible** on `main` (all five bindings, real decks, dev + prod build; specs pass). Need from the user: binding, built vs dev, and the deck (or whether a `noDrilldown`/lock flag is involved). Note the contextual tab never auto-switches (click the tab; PowerPoint behaves the same).

## 7. Open work, in priority order

Agents were told to wrap up; check `git log origin/main`, `gh issue list` and the branch lists first, because some may have landed after this was written (section 9 is updated if so).

1. **pptx-viewer open issues** (all in `D:\Development\pptx-wt\pptx-viewer`, `docs/guide/ui-migration.md` is the running tracker):
   * #396 adopt `pptx-ui-dialog-footer` in all remaining dialogs (branch `ui/dialog-footers-all`).
   * #373 remaining Home controls (font family/size selects, colour pickers, bullets/numbering, line spacing, direction, columns, Select menu, Group/Merge/Crop, popover contents) (branch `ui/home-remaining-controls`).
   * #342 umbrella: remaining = shared radio group + shared switch (AutoSave) (branch `ui/radio-switch-primitives`); then close with the drafted comment answering the reporter's four questions (in the agent report / `docs/guide/ui-migration.md`).
   * #386 closes when #396 is done (follow-ups #393-#398 are done except #396; #402 is new).
   * #402 sorter multi-selection/zoom/multi-slide clipboard exist only in React.
2. **Ribbon visual parity vs real PowerPoint** (user: "heavy UI parity issues"). Two agents audit with PowerPoint installed locally (`C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE`): Home/Insert/Draw/Design (`docs/guide/ribbon-parity.md`) and the other tabs + contextual tabs (`docs/guide/ribbon-parity-other-tabs.md`). Includes the SmartArt Design tab gap (Add Shape, Text Pane, Layouts, Reset, SmartArt Format tab; today only Change Colors and Styles).
3. **ooxml-core migration, wave 2** (user said finish issues first, then adoptions):
   * pptx-viewer adopts core `/collab` (replace `packages/shared/src/render/collaboration-*.ts` presence/publisher/departure/teardown/sync-gate/load-origin/broadcast-follow/assets modules; see core `docs/collab-area.md`; keep slide schema/reconcile as a `DocumentAdapter`; pass `LEGACY_PPTX_*` names where the old embedder contract must hold), bump `pptx-viewer-core` to core `^0.2.0`.
   * pptx-viewer shared web components -> `office-ui` (tag migration: each `pptx-ui-*` becomes a thin subclass of the `office-ui-*` constructor; plan in core `docs/office-ui-plan.md`; wave 2 needs a shared popover first: split button, menu/popover, colour picker, tooltip).
   * Core areas: extract the SmartArt layout engine/interpreters into `diagram`, then `drawingml`, then `chart`; move the pptx drawing-shape parser onto `diagram` with a parity test; port pptx off `XmlObject` onto the shared `xml` model; tighten pptx TS flags (relaxed today); SmartArt editing in docx; xlsx and Visio areas; third-party licence review.
   * Core follow-ups from agents: make `write-drawing.ts` unchanged-placeholder check key-order-insensitive (docx currently aligns JSON key order); docx floating SmartArt gets no reserved space in Print Layout (`adapt-floats.ts` skips images with empty `partName`); docx renderer limits (one text line per shape, no autofit, alpha/arcs dropped, 3D flattened); copy/paste of a diagram not tested; charts render as placeholders.
4. **docx-viewer**: publish setup (section 6), `oxfmt --check` shows 15 pre-existing files (CRLF); `tests/header-footer-link.spec.ts` flaked once in CI on `main` (6/6 locally; failed job re-run requested).
5. Housekeeping: Vue warns `Invalid prop x/y NaN got Undefined` for the new shared context menu adapter (found in e2e logs; harmless but a real adapter bug); **Vanilla read-only AI menu must match the other four bindings (user decision: restore the two icons it lost in the context-menu migration, #393).** Concretely: in `packages/vanilla/src/viewer/` (`context-menu-surface.ts` and the read-only AI menu), pass the same row icons the other bindings show (compare React/Vue/Angular/Svelte AI menu rows), extend the shared `pptx-ui-context-menu` row model with an optional `icon` if it has no icon slot yet (additive, with unit tests and token-coloured/forced-colours-safe rendering), and add a cross-binding e2e that asserts the AI menu rows (labels, order, icons present) are identical in all five bindings (`e2e/context-menu-migration.spec.ts` is the natural home); update the #393 section in `docs/guide/ui-migration.md` to remove the 'loses its two icons' note; Vanilla bug: a notes edit followed immediately by a thumbnail click is swallowed (blur commit re-renders the rail): needs its own issue; the Vue-only `pptx.notesToolbar.*` locale keys are unused.

## 8. Cleanup (done at the end of this session, at the user's request)

* **Remote:** no feature branches in any of the three repos (every branch was verified to be fully in `main` before deletion). Remote branches only appear temporarily when a subagent hands over work; land it and delete it.
* **Local branches:** ooxml-core and docx-viewer have only `main`. pptx-viewer: 177 stale branches deleted; remaining = `main` + the five in-flight agent branches (`ui/radio-switch-primitives`, `ui/home-remaining-controls`, `ui/dialog-footers-all`, `ui/ribbon-parity-home-insert`, `ui/ribbon-parity-other-tabs`): delete each right after it lands.
* **Backups** of commits that existed nowhere else (restore with `git fetch <bundle> <branch>:<branch>`): `D:\Development\_branch-backups\pptx-viewer-local-only.bundle` (37 branches, mostly pre-rebase copies of work that landed, plus older branches from earlier sessions) and `ooxml-core-local-only.bundle` (4 pre-rebase branches). docx-viewer had nothing to back up. The bundles can be deleted once the user is comfortable.
* **Worktrees:** ooxml-core and docx-viewer: only the main checkouts remain. pptx-viewer remaining: `pptx-viewer-new` (main), `pptx-wt\pptx-viewer` (assistant's scratch worktree, detached; fine to remove), five in-flight agent worktrees (`w342c`, `w373c`, `w396`, `wparity1`, `wparity2`), and three **kept on purpose because they hold uncommitted tracked source edits from earlier PR-review work**: `pptx-review-266` (20 files), `pptx-viewer-pr-merge-check` (2), `pptx-viewer-pr228` (1). Ask the user before touching those.
* **Folders Windows would not delete** (a process holds a handle; delete by hand after closing editors/dev servers or after a reboot; they are NOT registered worktrees any more): `D:\Development\docx-viewer\.claude\worktrees\agent-afe5e066b9c35a747`, `D:\Development\w386b\pptx-statusbar`, `D:\Development\w373c\before`, `D:\Development\wsmartart\pptx-viewer`. All other leftover agent folders and their empty parents were removed.
* Reusable script: none kept; the approach was "unregister non-force, bundle local-only commits, then `git branch -D`".

## 9. Agent outcomes after this document was written (updated as work landed)

Landed on pptx-viewer `main` after the first draft: #373 Home remainder (`daecae20b`, issue closed), #342 radio + switch primitives (`37f8c8f06`, issue closed with the reporter-facing summary posted), the docx SmartArt tests. Still open issues: #386 (closes with #396), #396, #402.

* **Ribbon parity, other tabs** (branch `ui/ribbon-parity-other-tabs`, worktree `D:\Development\wparity2\pptx-viewer`): rebased on `37f8c8f06`; a full rebuild + typecheck + focused e2e was running; if it is green, push to `main` and delete the branch. Fixed a stale shared test (`gallery-registry.test.ts`) that insisted the SmartArt Design tab have exactly two galleries. The gap tables are in `docs/guide/ribbon-parity-other-tabs.md` (priority list of unfixed gaps: contextual command groups, missing Table Layout / Chart Format / SmartArt Format / Video-Audio tabs, Transitions has 9 types vs ~50 and no Delay, Slide Show/Record/View missing groups that need models the viewer lacks). SmartArt Design now has Create Graphic / Layouts / Styles / Reset groups (Add Shape, Add Bullet, Layouts, Reset Graphic work; Promote/Demote/Move/Text Pane/Right-to-Left/Convert are disabled with a reason).
* **Ribbon parity, Home/Insert/Draw/Design** (branch `ui/ribbon-parity-home-insert`, worktree `D:\Development\wparity1\pptx-viewer`): agent `a147883576fbc9ec9` was asked to rebase onto the landed #373 (heavy overlap in the Home strips), keep both the shared controls and the PowerPoint-like tiles, run the all-five-binding e2e for the touched ribbon specs, and report the head sha. Land it from the main session when it reports (rebase again first; expect a conflict with the other-tabs branch in shared ribbon styles). Its doc: `docs/guide/ribbon-parity.md`. Gotchas it found: `pptx-ui-ribbon-command`/`-group` must not have instance fields named like their attributes (`label`, `launcher`; React 19 assigns element properties); Svelte's `button { min-height: 24px }` outranks the 22px rows (the Home CSS repeats `[data-pptx-chrome]`); Home groups now collapse below ~1900 px and Arrange extras below ~2600 px (`COLLAPSING_GROUPS` in `ribbon-overflow.ts` allows only home/insert/draw/design: extend it for the other tabs and update `ribbon-compact-layout`).
* **#396 dialog footers**: agent `a0a196cbad73ce160` was still working (branch `ui/dialog-footers-all`, worktree `D:\Development\w396\pptx-viewer`); land it, close #396 and #386.
* **Build facts**: after shared edits run `bun run build` (about 20 min under load), `bun run inline-shared` in `packages/angular`, and `rm -rf demos/demo-angular/node_modules/.vite` before e2e, or the dist-freshness guard aborts. Vanilla `PptxViewer.test.ts` can OOM a worker under load: rerun alone with `NODE_OPTIONS=--max-old-space-size=6144`.
* An agent ran `taskkill /IM chrome.exe` once (it may have closed a browser window of the user); agents were told not to do that again.
