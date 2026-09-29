# Outstanding parity work

Status as of 2026-09-28. Six parity workstreams were run in parallel and are now
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
notes), Insert Picture (PNG, JPEG, GIF, BMP, and SVG with a PNG fallback), insert/edit/remove hyperlink with Ctrl+K and a
"Place in this document" bookmark list, Ctrl+Click to follow links, a character style picker,
inherited run formatting and theme fonts/colors rendered in the editor, and a hidden-text toggle.
Fields are editable (complex `fldChar`/`instrText` and `fldSimple`): the result text edits in place
and the field structure is kept. A References tab has Table of Contents and Update Table: entries
come from heading styles (`\o` levels), and page numbers come from Print Layout pagination. Paragraph
tab stops (`w:tabs`) are modeled and saved, and Print Layout honors them (alignment and leaders),
along with paragraph indents and list labels.

Still missing:

- Section editing covers size/orientation, margins, columns, vertical alignment, all four
  section-break types, a different first page, page numbering and the document-wide odd/even
  headers setting; line numbering and page borders are still protected (only their presence is
  modeled). The continuous surface shows columns only for single-section documents (Print Layout
  shows all); vertically justified sections lay out top-aligned.
- Ctrl+Alt+F / Ctrl+Alt+D (insert footnote/endnote) are bound, but Windows browsers can report
  Ctrl+Alt as AltGr and not deliver them; the References ribbon buttons always work.
- Table of contents entries link to `_Toc` bookmarks on their headings with `PAGEREF` page numbers,
  as Word writes them. TC fields and custom style mappings (`\t`) are not collected. Tab stops
  are saved and honored in Print Layout and on the editing surface (measured after rendering, with
  leaders); the implicit hanging-indent stop is applied in Print Layout only.
- Pictures and links in headers, footers and notes are resolved, shown and saved with each part's own
  relationships; while a header, footer or note is being edited, the ribbon (formatting, Insert
  Picture, Link) targets it.
- Tables: default cell margins (`tblCellMar`) and row properties (height with its rule, keep-together, repeat-as-header) are modeled and render in both views; Print Layout clips exact-height rows and repeats header rows. They are read-only (no editor controls) and preserved on save; row properties are dropped from the model when rows are added or removed.
- Paragraph borders (`w:pBdr`, with `w:space` and grouped `between` lines) and shading render in both views from direct formatting and styles; they are read-only (no controls) and preserved on save.
- Paragraph keep options (`keepNext`, `keepLines`, `widowControl`, `contextualSpacing`) are modeled from styles and direct formatting, preserved through editing and used by Print Layout pagination (headings stay with the next paragraph); there are no editor controls for them yet.
- Toggle properties follow ECMA-376: explicit offs (`w:val="0"`, `w:u w:val="none"`) are kept and cancel styles, styles inherit through basedOn and XOR across style types, and direct formatting is absolute. Bold, Italic, Underline and Strikethrough toggle what the text shows, writing an explicit off for style-inherited formatting as Word does (with a selection; a collapsed caret toggles the typing marks).

### Ribbon parity with Word

Every tab uses Word's layout language: SVG icons, large stacked buttons, inline icon-and-label
buttons, dropdown commands (a transparent native `<select>` over an icon, caption and caret, so
`select.value` and assistive technology keep working), split colour buttons with swatch popovers,
editable font family and size combo boxes, a Styles gallery of preview tiles, group dialog
launchers, contextual tabs, and one baseline for group captions. Toggle buttons reflect the
selection (alignment, lists, bold and the rest), disabled controls look disabled, and tooltips
carry the shortcut ("Bold (Ctrl+B)"). A Word command is shown only where a real command exists.

What each tab offers that Word also has:

- **Home:** Paste, Cut, Copy (disabled without a selection), Format Painter (character formatting
  only), font family and size combo boxes showing the effective font (styles and theme included),
  Grow/Shrink Font, Change Case, Clear Formatting, Bold through Superscript, Text Highlight, Font
  Color, lists and levels, indents, Show/Hide paragraph marks (¶ only), alignment, Line Spacing,
  Shading and Borders (written to `w:shd` and `w:pBdr`), the Styles gallery with a "more" menu and
  a character-style menu, Find, Replace and Select All. The Font and Paragraph launchers open real
  dialogs: Font (family, style, size, colour, underline style and colour, strikethrough, double
  strikethrough, super/subscript, small and all caps, hidden, character spacing, with a preview)
  and Paragraph (alignment, indents, first-line/hanging, spacing, line rule and amount, contextual
  spacing, widow/orphan, keep with next, keep lines, page break before). Both apply only the
  fields you changed and show mixed selections as blank or indeterminate.
- **Insert:** Table, Pictures, Format picture, Link, Blank page, page and column breaks, Header,
  Footer and Page Number (creates the part, relationship, content type and `sectPr` reference on
  save, for existing packages and new documents), Date and Time (inserted as text in the display
  locale, not an updating field) and Symbol (27 common glyphs, not the full Symbol dialog).
- **Layout:** Margins (Normal, Narrow, Moderate, Wide), Size (Letter, Legal, Tabloid, Executive,
  A3, A4, A5, B5), Orientation, Vertical alignment, Columns, page-number format and start,
  different first page, odd and even, section breaks, Indent Left/Right (inches) and paragraph
  Before/After spacing.
- **References:** Table of Contents, Add Text (heading levels 1-3), Update Table, Footnote and
  Endnote.
- **Review:** Spelling (toggles the browser's spell checker; no bundled dictionary or grammar
  checker), Word Count (selection or document), Read Aloud (the browser's speech synthesis),
  tracked-change and comment commands, and the language and direction controls.
- **View:** Hidden text, Gridlines, thumbnails, Zoom, Zoom to 100%, One page, Page width and Print
  Layout.
- **Table (contextual):** appears only while the selection is in a table, as in Word.
- **Ribbon:** collapse (button, double-click a tab, or Ctrl+F1) with click-to-peek, and tab KeyTips
  after Alt or F10 (H, N, P, S, R, W, T; F opens File).
- **Shortcuts:** Ctrl+L, E, R, J (alignment), Ctrl+Shift+. and , (grow and shrink font), Ctrl+= and
  Ctrl+Shift+= (sub and superscript) and Ctrl+Space (clear formatting), listed in the help dialog.

Known limits of what is implemented:

- Inserting a header, footer or page number is not undone by Ctrl+Z. Header and footer content lives
  in the model outside the editor document, like in-place header edits, and the editor's history
  records only the section layout. The first section's header and footer are the ones shown and
  changed.
- Page Number adds a PAGE field to the default header or footer; it does not offer Word's numbered
  gallery styles, "Page X of Y" or first/odd/even variants.
- Shading is a solid hex fill and Borders use Word's default 0.5 pt automatic pen; there are no line
  styles, widths, colours, patterns or the Borders and Shading dialog.
- The Font dialog has no Advanced tab (ligatures, kerning, scale, position) or text effects.

Still not at parity:

- **Home:** Text Effects, Sort, Multilevel List. Colour palettes are short lists, not Word's theme,
  standard and More Colors grid. Line Spacing offers presets; the Paragraph dialog has the rest.
- **Insert:** Cover Page, Shapes, Icons, 3D Models, SmartArt, Charts, Screenshot, Text Box, WordArt,
  Drop Cap, Equation, Bookmark, Cross-reference, Comment gallery, Signature Line, Object.
- **Layout and References:** Line Numbers, Hyphenation, Watermark, Page Color, Page Borders, Position
  and Wrap Text, Citations, Bibliography, Captions, Index and Table of Authorities.
- **Review and View:** Editor pane, Thesaurus, Translate, Accessibility, Compare, Protect, Ink;
  Ruler, Navigation Pane headings, Read Mode, Web and Outline views, Multiple Pages, New Window,
  Split and Macros.
- **Dropdown galleries:** Margins, Size, Orientation, Columns and the others are text lists, not
  Word's thumbnail galleries; Zoom is a percentage list, not the Zoom dialog.
- **Ribbon:** no per-command KeyTips (only tabs), no overflow menus for narrow widths (panels
  scroll horizontally), no customisation UI, no table-design or layout tools beyond row and column
  commands.

### Fidelity

- Print Layout recalculates PAGE, NUMPAGES, SECTIONPAGES, DATE and TIME (with `\@` pictures); TOC,
  cross-references and other fields show Word's saved result. Field results are editable; the
  begin/code/separate/end structure is preserved and guarded against partial deletion.
- Picture crop and effects. Print Layout draws inline pictures on their lines and floating pictures
  at their `wp:positionH`/`wp:positionV` positions, in front of or behind text. Body text wraps
  around square, tight and through pictures on their larger side (using the picture's rectangle,
  not its outline) and below top-and-bottom pictures; `character`/`line` frames approximate to
  the column/paragraph, and wrap distances use Word's defaults. Pictures
  in table cells paginate inline; header and footer pictures (inline and floating) are drawn on
  every page; floats inside table cells are not positioned. The editing surface still approximates
  floats with CSS floats.
- Picture bullets, `numStyleLink`, Word's exact `lvlRestart` cascade.
- Prior-formatting snapshots for `rPrChange`/`pPrChange`, and table-structure revisions. Tracked moves are linked by name (accepting or rejecting either side resolves both) and saved with their range markers. Under Track Changes, dragging text or cutting and pasting the same text records a move; moving formatted content across table cells is recorded as a deletion and an insertion. Comments may span paragraphs (one range per comment); range edges outside any run move to the nearest commented text when edited.
- Print Layout measures and draws text with formatting inherited from document defaults, paragraph
  and character styles and theme fonts. It uses metric-compatible substitutes (Carlito, Caladea,
  Arimo, Tinos, Cousine) when Word's fonts are missing, and re-paginates once web fonts load. Single
  line height comes from the font's ascent and descent. Kerning, per-script fonts (East Asian and
  complex-script faces) and comparisons against Word-rendered references are still missing.
- Print Layout places footnotes at the bottom of the page where their reference lands (reserving
  the space, under Word's separator rule) and endnotes after the last paragraph; note references
  and marks print as superscript numbers. Footnotes that do not fit are not continued onto the
  next page, and multi-column pages put notes below the whole page, not under each column.
- New documents carry Word's modern defaults (Calibri 11pt, 8pt after, 1.08 lines; Normal,
  Heading 1–3, Title, TOC 1–3 and Hyperlink styles) and save them in `styles.xml`, so they open in
  Word as edited. The editing surface and Print Layout apply the same spacing and line heights.
- Print Layout tables use the document's grid widths, cell margins, borders (table style, table
  and cell), shading, vertical alignment and table indent or alignment. Vertically merged cells are
  drawn as one cell, but their text stays in the merge's first row.

### Engineering follow-ups

- Modules over the 300-line guideline: `web-component/src/component.ts` (~520: the element's public API, lifecycle and ribbon dispatch) and the `localization-strings.ts` data table.
