# Outstanding parity work

Status as of 2026-09-27. Six parity workstreams were run in parallel and are now
all merged into `main`. Four of them were interrupted before their authors
finished, so they were completed during integration. The gaps below are what
remains. None of this is Word parity; see the [parity roadmap](/parity-roadmap)
and the model warnings emitted at parse time.

## Merged workstreams

| Workstream | What landed |
| --- | --- |
| Numbering and lists | `numbering.xml` catalog, style-inherited `numPr`, computed labels (decimal, Roman, letter, ordinal, cardinal/ordinal text, bullet, legal, multilevel `lvlText`), list commands, Enter-on-empty exits a list, Tab/Shift+Tab levels, numbering part surgery on save. |
| Sections, breaks, headers/footers, notes | `DocumentModel.sections`; editable page/column breaks and `pageBreakBefore` (Ctrl+Enter); read-only headers/footers and footnotes/endnotes. |
| Pagination / Print Layout | `@christophervr/docx-layout` engine (line breaking, spacing, indents, widow/orphan, keep rules, table row splitting, columns) and a read-only paginated Print Layout view with printing, driven by the section model. |
| Tracked changes and comments | Revisions on runs and paragraph marks, markup display modes, accept/reject (one/all), next/previous, Track Changes mode, comments pane with replies and resolve, `comments.xml` writing. |
| Theme, character styles, tables | Theme colors/fonts, character styles with toggle-property inheritance, extra run properties (caps, small caps, double strike, underline style/color, spacing, shading, hidden text), table grid/widths/merges/borders/shading/table styles. |
| Pictures, hyperlinks, bookmarks | Inline pictures rendered from package media, floating/unsupported drawings as placeholders, hyperlink targets (relationships, anchors, simple `HYPERLINK` fields), read-only bookmarks preserved across edits. |

### Fixes made during integration

- Track Changes mode applied each edit twice; the tracked replacement now replaces the original step.
- New hyperlink/image relationships could reuse an id declared only in `document.xml.rels` (for example the styles relationship).
- Comment reference runs were parsed as text and duplicated on every save.
- Editing text dropped character styles, caps and theme colors (and then stripped them on save); editing a simple table dropped its grid, borders and style.
- Pictures had no image source in the editor; imported `javascript:` link targets were rendered as live `href`s.

## Remaining gaps

### Editor commands and UI

- Insert Picture command (core package surgery exists: media part, relationship, content type, `wp:inline`), and picture resize.
- Insert/edit/remove hyperlink (Ctrl+K), Ctrl+Click to follow http/https/mailto links, and scrolling to bookmark anchors.
- Character style picker; show/hide hidden text; resolved (inherited) run formatting and theme colors are modeled in core but not yet rendered by the editor.
- Header/footer, footnote/endnote and section editing; multi-column rendering on the continuous surface.
- Table rendering in the editor for colspan/rowspan, borders, shading and cell alignment (`table-render.ts` exists; verify against Word-authored fixtures), and read-only nested tables.

### Fidelity

- Field recalculation (`PAGE`, `NUMPAGES`, `DATE`, …); complex `fldChar`/`instrText` hyperlinks remain protected.
- Picture crop and effects; floating object position and text wrapping.
- Picture bullets, `numStyleLink`, Word's exact `lvlRestart` cascade.
- Move linkage for `moveFrom`/`moveTo`, prior-formatting snapshots for `rPrChange`/`pPrChange`, table-structure revisions, and comments spanning multiple paragraphs.
- Pagination against real font metrics, footnote placement, floating objects, and Word-rendered reference comparisons.
- Headers, footers and notes do not resolve their own relationship parts, so pictures and external links inside them are not shown.

### Engineering follow-ups

- Modules over the 300-line guideline: `web-component/src/component.ts` (449), `localization.ts` (411), `ribbon.ts` (394), `schema.ts` (364), `core/src/model.ts` (355).
- Superseded helpers from the parallel branches remain (`write-hyperlink.ts`, `write-safety.ts`, and `buildInlineNodes`/`gatherOldRuns` in `write-revisions.ts`) now that `write-inline.ts` is the single inline writer; remove them once no tests depend on them.
- The Track Changes plugin handles transactions of plain replace steps; multi-step transactions mixing deletions and insertions need position mapping between steps.
- The browser suite is intermittently flaky on first load (the same behavior reproduces on earlier commits).
- `bun run fmt:check` fails on a clean checkout because `core.autocrlf=true` produces CRLF files while oxfmt expects LF. A `.gitattributes` with `* text=auto eol=lf` would fix it.
