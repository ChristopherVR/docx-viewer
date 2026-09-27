# Outstanding parity work

Status as of 2026-09-27. Six parity workstreams were started in parallel. Two are
merged into `main`; four were interrupted before completion and are preserved as
unmerged work-in-progress commits on local branches. None of the WIP branches has
been verified end-to-end, so each must be finished, rebased and re-tested before
merging.

## Merged

| Workstream | Summary |
| --- | --- |
| Numbering and lists | `numbering.xml` catalog, style-inherited `numPr`, computed list labels (decimal, Roman, letter, ordinal, cardinal/ordinal text, bullet, legal, multilevel `lvlText`), bulleted/numbered list and level commands, Enter-on-empty exits a list, Tab/Shift+Tab levels, numbering part surgery on save. |
| Sections, breaks, headers/footers, notes | Every `sectPr` modeled in `DocumentModel.sections`; page/column breaks and `pageBreakBefore` modeled and editable (Ctrl+Enter); headers/footers (default/first/even) and footnotes/endnotes parsed read-only and rendered around the continuous surface. The body parser moved to `packages/core/src/block-parser.ts` so parts share it. |

Remaining gaps for the merged work are listed in the [parity roadmap](/parity-roadmap)
and emitted as model warnings (picture bullets, `numStyleLink`, exact `lvlRestart`
cascade, section/header/footer/note editing, multi-column rendering, field
recalculation, editing paragraphs that contain note references).

## Interrupted work in progress

Each branch was created from `cf5c660` and holds one WIP commit. They predate the
`block-parser.ts` extraction: any change they make to `parseRun`, `parseParagraph`
or `parseTable` in `parse.ts` must be ported into `block-parser.ts` when rebasing.
Several also reformatted unrelated files with oxfmt (LF line endings); drop those
hunks when rebasing.

### Images, hyperlinks, bookmarks — `worktree-agent-a575211d2cd290b94`

Done so far (core): `drawing.ts`, `write-drawing.ts`, `hyperlink.ts`,
`write-hyperlink.ts`, `bookmarks.ts`, `relationship-allocator.ts`,
`package-parts.ts`, `write-safety.ts`, `parse-warnings.ts`, and
`inline-content.test.ts`; web component `inline-content-schema.ts`.

Outstanding:
- Finish `run-adapter.ts` (`appendInlineNode`) for the image node and link mark.
- Editor commands: Insert Picture (media part, relationship, content type, `wp:inline`
  with unique `docPr` id), insert/edit/remove link (Ctrl+K), Ctrl+Click for
  http/https/mailto only, and scrolling to bookmark anchors.
- Placeholders for floating (`wp:anchor`) images, VML `w:pict`, charts and SmartArt.
- Localization (EN/FR), roadmap update, and a full test/typecheck/format pass.

### Theme, character styles, table fidelity — `worktree-agent-a4e4fe04fbfc622c5`

Done so far (core): theme parsing and theme color transforms (`theme.ts`,
`theme-color.ts`, `theme-model.ts`), character styles and run inheritance with
toggle semantics (`character-styles.ts`, `run-formatting.ts`, `run-properties.ts`,
`run-style-model.ts`, `underline.ts`, `write-run-extra.ts`,
`write-run-validation.ts`), and table model, parsing, borders and styles
(`table-model.ts`, `parse-table.ts`, `resolve-table.ts`, `table-borders.ts`,
`table-styles.ts`); web component `table-model-adapter.ts`, `table-render.ts`.

Outstanding:
- Wire resolved run formatting into `run-adapter.ts` so the editor renders
  inherited fonts, colors and toggles while the toolbar still edits direct formatting.
- Character style picker; hidden text (`vanish`) display toggle.
- Editor rendering of colspan/rowspan, borders, shading and cell alignment, plus
  read-only nested tables.
- Localization, roadmap update, and a full verification pass.

### Tracked changes and comments — `worktree-agent-a134163dd3fc29c5c`

Done so far (core): `parse-revisions.ts`, `write-revisions.ts`,
`revision-commands.ts` (accept/reject transforms), `comments.ts`,
`write-comments.ts`, `settings.ts` (`w:trackRevisions`), `zip-parts.ts`, with tests;
web component `review-schema.ts`, `review-display.ts`, `review-commands.ts`,
`review-controller.ts`, `review-ribbon.ts`, `comment-commands.ts`,
`comments-panel.ts`, `track-changes-mode.ts`.

Outstanding:
- The agent was mid-way through moving Review ribbon construction from `ribbon.ts`
  into `review-ribbon.ts`; finish that and reconcile with the Language group in the
  Review tab.
- Confirm Track Changes mode (insertions attributed to the collaboration identity,
  deletions kept as struck-through `w:del`, deleting your own insertion removes it)
  and markup display modes in the browser.
- Localization, roadmap/docs page, and a full verification pass.

### Pagination / Print Layout — `worktree-agent-af4d15a11cb644e14`

Done so far: new `packages/layout` (`@christophervr/docx-layout`) pagination
engine with an injectable text measurer, wired into the workspace; web component
`print-layout.ts`, `print-layout-view.ts`, `print-layout-cursor.ts` with tests. The
agent had reached a clean typecheck and was running the full suite and formatter.

Outstanding:
- Adapt the engine's section input to the merged `DocumentModel.sections` shape
  (multi-column flow, per-section page geometry) and honor `pageBreakBefore` and
  `break: 'page' | 'column'` runs.
- Browser test for the Print Layout toggle; confirm the responsive ribbon test
  still passes with the View tab additions.
- Roadmap section 3 update and a full verification pass.

## Merge checklist for each branch

1. Rebase onto `main`; port parser changes into `block-parser.ts`; drop formatting-only hunks.
2. Keep ribbon additions inside the desktop width (`tests/responsive.spec.ts`);
   prefer existing groups or non-Home tabs.
3. `bun run test`, `bun run typecheck`, `bun run test:browser` must pass.
4. Update `docs/parity-roadmap.md` and model warnings to state what remains unsupported.

## Known repository issue

`bun run fmt:check` reports the same files on a clean `main` checkout: with
`core.autocrlf=true` files check out with CRLF, while oxfmt expects LF. Consider a
`.gitattributes` with `* text=auto eol=lf`.
