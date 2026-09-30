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
  headers setting and line numbering (Layout > Line Numbers: none, continuous, restart each page or
  section; the count-by, start and distance values are read and kept but have no controls, and line
  numbers are drawn in Print Layout only, counting every line of body paragraphs, tables excluded,
  with no per-paragraph suppression); page borders are still protected (only their presence is
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

The **File** tab follows Word's backstage: Info (file name and save state, compatibility notes,
tracked-change and comment counts, and a properties panel of words, characters, paragraphs,
tables, sections, page size and orientation), New, Open, Save, Save As (takes a file name and keeps
the opened format), Print, Export (PDF through the print dialog, DOCX, plain text) and Options
(display language, theme, review author). Not implemented: Home and recent files, Share, Account,
Close, Protect Document, Version History, author/title/tags properties (the model has no core
properties), and the page count on Info (pagination is a Print Layout result).

What each tab offers that Word also has:

- **Home:** Paste, Cut, Copy (disabled without a selection), Format Painter (character formatting
  only), font family and size combo boxes showing the effective font (styles and theme included),
  Grow/Shrink Font, Change Case, Clear Formatting, Bold through Superscript, Text Highlight, Font
  Color, lists and levels, indents, Sort (paragraphs, locale collation, numbers as numbers),
  Show/Hide paragraph marks (¶ only), alignment, Line Spacing, Shading and Borders (written to
  `w:shd` and `w:pBdr`, plus a Horizontal Line), a Styles gallery with previous/next/more controls,
  an expanded gallery of recommended paragraph and character styles with Clear Formatting, and a
  docked Styles pane listing every style; Find (with Go To headings, bookmarks, tables and
  pictures), Replace, and Select (All, Objects). The Font and Paragraph launchers open real
  dialogs: Font (family, style, size, colour, underline style and colour, strikethrough, double
  strikethrough, super/subscript, small and all caps, hidden, character spacing, with a preview)
  and Paragraph (alignment, indents, first-line/hanging, spacing, line rule and amount, contextual
  spacing, widow/orphan, keep with next, keep lines, page break before). Both apply only the
  fields you changed and show mixed selections as blank or indeterminate.
- **Insert:** Text boxes in an opened document (`wps:txbx`, including ones inside
  `mc:AlternateContent`) show their text read-only in a boxed placeholder on the editing surface and
  are kept byte-for-byte on save; they cannot be inserted or edited, and Print Layout does not draw
  them (it says so). Drop Cap (Dropped or In margin, three lines: the first letter moves into its own
  `w:framePr` paragraph as in Word; the editing surface floats it, while Print Layout folds it into the
  next paragraph as a raised initial with a warning; no Drop Cap Options dialog for font, lines or
  distance), Table (a size grid with keyboard control and typed rows and columns), Cover Page (one
  plain design), Pictures, Format picture, Link, Bookmark (add, move, delete, go to, Word's naming
  rule), Blank page, page and column breaks, Header,
  Footer and Page Number (creates the part, relationship, content type and `sectPr` reference on
  save, for existing packages and new documents), Date and Time (inserted as text in the display
  locale, not an updating field) and Symbol (27 common glyphs, not the full Symbol dialog).
- **Layout:** a Page Setup dialog (custom margins, gutter, header and footer distances,
  orientation, paper size and custom width and height, validated so text keeps room), Margins (Normal, Narrow, Moderate, Wide), Size (Letter, Legal, Tabloid, Executive,
  A3, A4, A5, B5), Orientation, Vertical alignment, Columns, page-number format and start,
  different first page, odd and even, section breaks, Page Color (the colour picker; written as
  `w:background` with `displayBackgroundShape`, drawn on the editing surface and in Print Layout, not
  printed, and undoable; theme and gradient page fills are not modeled), Indent Left/Right (inches) and paragraph
  Before/After spacing.
- **References:** Table of Contents (levels 1-3, 1-2 or 1-5, or Remove), Add Text (heading levels
  1-3), Update Table, Insert Caption (a `SEQ` field; every caption of a label is renumbered when
  one is added, and the Caption style is used when the document has one), Cross-reference (`REF` or
  `PAGEREF` fields to headings, bookmarks and captions; hidden `_Ref` bookmarks are added as
  needed and the result text is computed when inserted, not updated later), Footnote and Endnote.
- **Review:** Spelling (toggles the browser's spell checker; no bundled dictionary or grammar
  checker), Word Count (selection or document), Read Aloud (the browser's speech synthesis),
  tracked-change commands, comments (New, Delete, Previous, Next, Show), and the language and
  direction controls.
- **View:** Ruler (inch ticks from the left margin, shaded margins and the current paragraph's
  indent markers; drag a marker to change the paragraph's indents, snapped to sixteenths of an inch and undone as one step; margin edges and tab stops are not draggable), Hidden text, Gridlines,
  thumbnails, Zoom, Zoom to 100%, One page, Page width and Print Layout.
- **Table (contextual):** appears only while the selection is in a table, as in Word.
- **Ribbon:** collapse (button, double-click a tab, or Ctrl+F1) with click-to-peek, and tab KeyTips
  after Alt or F10 (H, N, P, S, R, W, T; F opens File); choosing a tab shows a tip on each of its
  commands (Alt, H, 1 is Bold), and Escape steps out.
- **Shortcuts:** Ctrl+L, E, R, J (alignment), Ctrl+Shift+. and , (grow and shrink font), Ctrl+= and
  Ctrl+Shift+= (sub and superscript) and Ctrl+Space (clear formatting), listed in the help dialog.

Ribbon layout: every panel is one fixed height (104 px, as Word's is), group captions are pinned
to the bottom and rows sit at the top, Clipboard has Paste with labelled Cut, Copy and Format
Painter beside it, and the Styles gallery shows the built-in names capitalised ("Heading 1").
Font Color and Shading open Word's colour picker (Theme Colors as ten columns over five variants,
Standard Colors, and More Colors through the browser's colour dialog; Shading adds No Color) using
the Office theme palette, not the document's theme. Margins, Size, Orientation, Vertical alignment
and Columns open galleries of page thumbnails with detail lines (inches for margins and paper); the
native select stays underneath, so the keyboard and assistive technology still work. Icons carry
Word's accent hues and a pressed control is darker than a hovered one. Styles live in
`styles/ribbon*.css`, loaded in order.

View > Zoom opens Word's Zoom dialog (200%, 100%, 75%, page width, whole page or a percentage, with
a preview); Many pages, Text width and Wrap to window are not offered. Large icons are fitted to a
common frame (measured from each drawing) so they fill their slot evenly.

Layout > Hyphenation is None or Automatic (`w:autoHyphenation`, undoable): the editing surface asks the
browser to hyphenate by the text's language, while Print Layout's line breaker does not hyphenate yet,
so it lays lines out as if it were off. View > Multiple pages fits two Print Layout pages side by side
(any other zoom returns to one column, and the Zoom dialog offers it as Two pages); zoom now also
applies to Print Layout, which ignored it before.

Known limits of what is implemented:

- When a tab is wider than the window, its rightmost groups fold into one dropdown button each
  (the group's live controls appear in a panel under it; a command closes the panel, Escape too),
  and unfold when the window widens. Folding goes strictly right to left, with no priority order
  or intermediate sizes as in Word (a group does not first shrink its buttons), and the panel only
  scrolls when even folded groups do not fit.
- Home fits at 1280 px only because the Styles gallery shows about three tiles (Word shows more);
  in French and German the Editing group folds at that width.
- Icons are one line-icon set with a tinted stroke; Word's are two-tone and drawn per command.

- Inserting a header, footer or page number is not undone by Ctrl+Z. Header and footer content lives
  in the model outside the editor document, like in-place header edits, and the editor's history
  records only the section layout. The first section's header and footer are the ones shown and
  changed.
- Page Number adds a PAGE field to the default header or footer; it does not offer Word's numbered
  gallery styles, "Page X of Y" or first/odd/even variants.
- The Borders menu presets use Word's default 0.5 pt automatic pen. The Borders and Shading dialog
  (paragraphs) sets sides, one style (single, double, dotted, dashed), width and colour for all chosen
  sides, and a solid fill; there is no per-side pen, Box/Shadow/3-D setting, patterns, page borders or
  table/cell targeting, and no preview.
- The Font dialog has no Advanced tab (ligatures, kerning, scale, position) or text effects.

Still not at parity:

- **Home:** Text Effects; Multilevel List offers two styles (1. 1.1. 1.1.1. and 1. a) i.), not Word's gallery of
  heading-linked and bullet outlines, and has no Define New Multilevel List dialog. Text Highlight keeps the 17-colour list (Word's is a short
  list too). Line Spacing offers presets; the Paragraph dialog has the rest.
- **Insert:** Cover Page galleries (one plain design exists), Shapes, Icons, 3D Models, SmartArt,
  Charts, Screenshot, inserting Text Boxes, WordArt, Equation, Signature Line, Object.
- **Layout and References:** Manual Hyphenation and its options, Watermark, Page Borders, Position
  and Wrap Text, Citations, Bibliography, Table of Figures, Index and Table of Authorities, and
  automatic updating of `REF`, `PAGEREF` and `SEQ` results (they are computed when inserted; captions
  renumber immediately, cross-references do not).
- **Review and View:** Editor pane, Thesaurus, Translate, Accessibility, Compare, Protect, Ink;
  Navigation Pane headings, Read Mode, Web and Outline views, New Window,
  Split and Macros.
- **Dropdown galleries:** Margins, Size, Orientation, Vertical alignment, Columns, Borders and Line
  spacing are galleries with thumbnails (Line spacing ends with Line Spacing Options, which opens the
  Paragraph dialog; Borders ends with Borders and Shading).
  Page number format, Sort, Change case and similar stay plain text lists, as Word's Change Case
  does. Zoom is a percentage list next to a Zoom dialog. Margins ends with Custom Margins, which opens Page Setup; Columns has no Left, Right or More Columns.
- **Ribbon:** command KeyTips use Word's keys for the commands that share a name with Word's (Bold `1`,
  Paste `V`, Font Color `FC` and so on) and two letters from the label for the rest, so a few differ
  from Word's; there are no group-level tips and they are not localised. Customisation is limited to showing and hiding commands (File > Customize Ribbon; no reordering,
  custom tabs or groups, and nothing is stored between sessions unless the host saves the ids from
  the `ribbon-customize` event), no table-design or layout tools beyond row and column
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
