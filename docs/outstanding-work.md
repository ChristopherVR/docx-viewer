# Outstanding parity work

Status as of 2026-09-27. Six parity workstreams were run in parallel and are now
all merged into `main`. Four of them were interrupted before their authors
finished, so they were completed during integration. The gaps below are what
remains. None of this is Word parity; see the [parity roadmap](/parity-roadmap)
and the model warnings emitted at parse time.

## Merged workstreams

| Workstream                               | What landed                                                                                                                                                                                                                                                        |
| ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Numbering and lists                      | `numbering.xml` catalog, style-inherited `numPr`, computed labels (decimal, Roman, letter, ordinal, cardinal/ordinal text, bullet, legal, multilevel `lvlText`), list commands, Enter-on-empty exits a list, Tab/Shift+Tab levels, numbering part surgery on save. |
| Sections, breaks, headers/footers, notes | `DocumentModel.sections`; editable page/column breaks and `pageBreakBefore` (Ctrl+Enter); read-only headers/footers and footnotes/endnotes.                                                                                                                        |
| Pagination / Print Layout                | `@christophervr/docx-layout` engine (line breaking, spacing, indents, widow/orphan, keep rules, table row splitting, columns) and a read-only paginated Print Layout view with printing, driven by the section model.                                              |
| Tracked changes and comments             | Revisions on runs and paragraph marks, markup display modes, accept/reject (one/all), next/previous, Track Changes mode, comments pane with replies and resolve, `comments.xml` writing.                                                                           |
| Theme, character styles, tables          | Theme colors/fonts, character styles with toggle-property inheritance, extra run properties (caps, small caps, double strike, underline style/color, spacing, shading, hidden text), table grid/widths/merges/borders/shading/table styles.                        |
| Pictures, hyperlinks, bookmarks          | Inline pictures rendered from package media, floating/unsupported drawings as placeholders, hyperlink targets (relationships, anchors, simple `HYPERLINK` fields), read-only bookmarks preserved across edits.                                                     |

### Fixes made during integration

- Track Changes mode applied each edit twice; the tracked replacement now replaces the original step.
- New hyperlink/image relationships could reuse an id declared only in `document.xml.rels` (for example the styles relationship).
- Comment reference runs were parsed as text and duplicated on every save.
- Editing text dropped character styles, caps and theme colors (and then stripped them on save); editing a simple table dropped its grid, borders and style.
- Pictures had no image source in the editor; imported `javascript:` link targets were rendered as live `href`s.

## Remaining gaps

### Editor commands and UI

Done since the merge: per-section page setup and section breaks (undoable), in-place header/footer and footnote/endnote editing, Insert Footnote/Endnote (creating the notes part when needed), headers/footers with live PAGE/NUMPAGES in Print Layout, table style rendering, picture resize and alt text, Word-style window chrome (title bar with quick access, "Tell me"
command search, Editing/Viewing mode; File backstage; status bar with zoom and compatibility
notes), Insert Picture (PNG, JPEG, GIF, BMP), insert/edit/remove hyperlink with Ctrl+K and a
"Place in this document" bookmark list, Ctrl+Click to follow links, a character style picker,
inherited run formatting and theme fonts/colors rendered in the editor, and a hidden-text toggle.

Still missing:

- SVG pictures (they need a PNG fallback part).
- Section editing covers size/orientation, margins, columns, vertical alignment, all four
  section-break types, a different first page, page numbering and the document-wide odd/even
  headers setting; line numbering and page borders are still protected (only their presence is
  modeled). The continuous surface shows columns only for single-section documents (Print Layout
  shows all); vertically justified sections lay out top-aligned.
- Ctrl+Alt+F / Ctrl+Alt+D (insert footnote/endnote) are bound, but Windows browsers can report
  Ctrl+Alt as AltGr and not deliver them; the Insert ribbon buttons always work.
- Pictures and links in headers, footers and notes are resolved, shown and saved with each part's own
  relationships; while a header, footer or note is being edited, the ribbon (formatting, Insert
  Picture, Link) targets it.
- Table style run formatting and cell margins render; tables' own default cell margins (`tblCellMar`) and row heights are not modeled yet.
- Underline/strikethrough cancelled by the style hierarchy still render (bold, italic and caps are handled).

### Fidelity

- Print Layout recalculates PAGE, NUMPAGES, SECTIONPAGES, DATE and TIME (with `\@` pictures); TOC,
  cross-references and other fields show Word's saved result, and paragraphs containing complex
  fields stay protected from edits.
- Picture crop and effects; floating pictures follow their wrapping approximately (floats), but exact
  positions, vertical offsets and in-front/behind layering are not reproduced.
- Picture bullets, `numStyleLink`, Word's exact `lvlRestart` cascade.
- Move linkage for `moveFrom`/`moveTo`, prior-formatting snapshots for `rPrChange`/`pPrChange`, table-structure revisions, and comments spanning multiple paragraphs.
- Pagination against real font metrics, footnote placement, floating objects, and Word-rendered reference comparisons.

### Engineering follow-ups

- Modules over the 300-line guideline: `web-component/src/component.ts` (~530: the element's public API, lifecycle and ribbon dispatch) and the `localization-strings.ts` data table.
- The Track Changes plugin handles transactions of plain replace steps; multi-step transactions mixing deletions and insertions need position mapping between steps.
