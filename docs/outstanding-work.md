# Outstanding parity work

Status as of 2026-09-30. Six parity workstreams were run in parallel and are now
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
  section, plus Line Numbering Options for count-by, start and automatic/custom distance from text;
  numbers are drawn in Print Layout only, with per-paragraph suppression in the menu and Paragraph
  dialog, including style inheritance and explicit offs. Tables are excluded.
  The model uses visible start numbers and converts Word's zero-based `w:start` at the package boundary);
  page borders are still protected (only their presence is
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
- Tables: default cell margins (`tblCellMar`) and row properties (height with its rule, keep-together, repeat-as-header) are modeled and render in both views; Print Layout clips exact-height rows and repeats header rows. Table > Properties now edits these settings for rectangular tables, with mixed selected rows, validation and one-step undo. Row properties follow surviving rows during structural edits. Opened DOCX packages save changes to these fields while preserving unrelated table XML. Merged/nested/complex tables remain protected; individual cell margins have no controls. Exact-height rows are a minimum on the continuous editing surface and clipped in Print Layout.
- Paragraph borders (`w:pBdr`, with `w:space` and grouped `between` lines) and shading render in both views from direct formatting and styles. Home offers presets and a Borders and Shading dialog; art borders and a table/cell border editor remain missing.
- Paragraph keep options (`keepNext`, `keepLines`, `widowControl`, `contextualSpacing`) are modeled from styles and direct formatting, preserved through editing and used by Print Layout pagination (headings stay with the next paragraph). The Paragraph dialog exposes these settings, including mixed selections.
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
Close, Protect Document, Version History, and the page count on Info (pagination is a Print Layout result).
Info edits Title, Subject, Author, Tags and Comments (`docProps/core.xml`, created when missing; other
elements such as created/modified dates are kept untouched, and the last-modified fields are not
updated on save). Edits mark the document dirty but are not part of undo history.

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
  them (it says so). Drop Cap (Dropped or In margin, with font, line count and distance options: the first letter moves into its own
  `w:framePr` paragraph as in Word; the editing surface floats it, while Print Layout folds it into the
  next paragraph as a raised initial with a warning), Table (a size grid with keyboard control and typed rows and columns), Cover Page (one
  plain design), Pictures, Format picture, Link, Bookmark (add, move, delete, go to, Word's naming
  rule), Blank page, page and column breaks, Header,
  Footer and Page Number (creates the part, relationship, content type and `sectPr` reference on
  save, for existing packages and new documents), Date and Time (inserted as text in the display
  locale, not an updating field) and Symbol (27 common glyphs, not the full Symbol dialog).
- **Layout:** a Page Setup dialog (custom margins, gutter, header and footer distances,
  orientation, paper size and custom width and height, validated so text keeps room), Margins (Normal, Narrow, Moderate, Wide), Size (Letter, Legal, Tabloid, Executive,
  A3, A4, A5, B5), Orientation, Vertical alignment, Columns (One, Two, Three, Left and Right;
  More Columns edits count, equal or individual widths, spacing and separator lines, validates
  available width, targets the current section and supports undo), page-number format and start,
  different first page, odd and even, section breaks, Page Color (the colour picker; written as
  `w:background` with `displayBackgroundShape`, drawn on the editing surface and in Print Layout, not
  printed, and undoable; theme and gradient page fills are not modeled), Indent Left/Right (inches) and paragraph
  Before/After spacing.
- **References:** Table of Contents (levels 1-3, 1-2 or 1-5, or Remove), Add Text (heading levels
  1-3), Update Table, Insert Caption (a `SEQ` field; every caption of a label is renumbered when
  one is added, and the Caption style is used when the document has one), Cross-reference (`REF` or
  `PAGEREF` fields to headings, bookmarks and captions; hidden `_Ref` bookmarks are added as
  needed), Table of Figures (`TOC \h \z \c "Figure"` for Figure, Table or Equation captions: entries are the captions with Print Layout page numbers and links to `_Toc` bookmarks, rebuilt by Update Table of Figures; custom `\t` style mappings and `\a` are not collected), Update Fields (renumbers every `SEQ` caption, then refreshes `REF` and `PAGEREF`
  results from their bookmarks and Print Layout pages, then rebuilds the table of contents and any tables of figures; the field changes are one undo step, each table rebuild is its own; fields with other
  switches such as `\n`, `\r`, `\w` or `\p`, and fields whose bookmark is missing, are left
  as they are; it is a command, not automatic on edit), Footnote and Endnote.
- **Review:** Spelling (toggles the browser's spell checker; no bundled dictionary or grammar
  checker), Word Count (selection or document), Read Aloud (the browser's speech synthesis),
  tracked-change commands, comments (New, Delete, Previous, Next, Show), and the language and
  direction controls.
- **View:** Ruler (inch ticks from the left margin, shaded margins and the current paragraph's
  indent markers; drag a marker to change the paragraph's indents, snapped to sixteenths of an inch and undone as one step; margin edges and tab stops are not draggable), Hidden text, Gridlines,
  thumbnails, Navigation pane (body headings, inherited heading styles, collapsible hierarchy,
  keyboard navigation, current-heading selection and live refresh; also works in Viewing mode),
  Zoom, Zoom to 100%, One page, Page width and Print Layout.
- **Table (contextual):** appears only while the selection is in a table, as in Word. Properties opens row height (At least/Exactly), Allow row to break across pages, Repeat as header row, and default cell margins in inches. Row settings apply to selected rows; existing individual cell margin overrides stay in place. These controls are disabled for complex tables and Viewing mode. Word 16 COM confirmed row semantics and the 0/0/0.075/0.075-inch default cell margins.
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
- The Styles gallery shows about three tiles (Word shows more). Compact multilevel-list and
  paragraph controls keep the full command row inside its group, so Show/Hide ¶ cannot overlap
  the Styles gallery. In French and German the Editing group folds at 1280 px.
- Icons are one line-icon set with a tinted stroke; Word's are two-tone and drawn per command.

- Inserting headers, footers and page numbers and editing header/footer text share the body undo
  history. Ctrl+Z/Ctrl+Y, ribbon history and Quick Access history restore both content and previews.
  Browser contracts cover header insertion, typing and body history across all six bindings.
  The selection's section supplies the preview and receives inserted parts; edits to a part linked
  by several sections update all its uses. Word COM confirmed that a page number inserted in section
  2 exports as one field in that section and no fields in section 1. Header and Footer menus insert
  blank default, first-page or even-page parts. First-page insertion enables Different first page in
  that section; even-page insertion enables the document-wide odd/even setting, in the same undo step.
  Page Number follows the first/even/default slot being edited, or the default slot from the body.
  Word COM confirmed a three-page export enables both settings, with PAGE/NUMPAGES in the first-page
  footer, PAGE in the even-page footer and no fields in the default footer. Decorative header/footer
  galleries remain missing. Insert > Link to Previous toggles the active
  first/even/default header/footer story (disabled in section 1, outside a story and in Viewing mode).
  Inherited content now appears and edits in the continuous preview. Unlinking creates an independent
  copy, preserving the source XML and relationships; linking removes the local reference and uses the
  previous section's story. Both are undoable. Word COM confirmed independent section-2 header text
  with LinkToPrevious false, and matching inherited text with LinkToPrevious true after relinking.
  A contextual Header & Footer tab appears and opens while editing a story, with Go to Header/Footer,
  Previous/Next Section, Link to Previous, the first/odd/even settings and Close Header and Footer.
  Navigation keeps the story variant and updates the body section selection. Empty stories can be
  visited without changing the model; their parts are created only on typing, with undo removing
  that first edit and its part. Closing removes the empty preview and returns focus to the body;
  switching to Viewing mode also closes an active story editor.
  Header/footer formatting controls now reflect the active story rather than the body selection.
  Browser contracts cover context, navigation, first typing, undo and close across all six bindings.
  Word COM confirmed that typing into section 2's previously empty footer creates independent text
  there and leaves section 1's footer blank.
  Position controls edit Header from Top and Footer from Bottom in inches for the active section,
  preserving other measurements exactly, validating the same 0–22-inch range as Page Setup and
  supporting undo/redo. Word COM confirmed a two-section export keeps section 1 at 0.5/0.5 inches
  while section 2 uses 0.75/0.875 inches. Review the distances in Print Layout; the continuous story
  preview stays outside the body editing surface.
- Page Number offers page-thumbnail choices for plain numbers or "Page X of Y", at the left,
  center or right of the default header or footer. These are PAGE and NUMPAGES fields: Print Layout
  and Word update both (Word COM confirmed a two-page export reads "Page 1 of 2"). Switching the
  generated pattern is undoable and preserves surrounding paragraphs. Imported custom page-number
  text and complex codes are kept; adding a total extends that paragraph, while Plain changes only
  its alignment. Word's decorative gallery styles remain missing; slot selection is through editing
  the corresponding first/even/default header or footer, rather than a separate page-number gallery.
- The Borders menu presets use Word's default 0.5 pt automatic pen. The Borders and Shading dialog
  (paragraphs) sets sides, one style (single, double, dotted, dashed), width and colour for all chosen
  sides, and a solid fill; there is no per-side pen, Box/Shadow/3-D setting, patterns, page borders or
  table/cell targeting, and no preview.
- The Font dialog now has Font and Advanced tabs. Advanced sets horizontal scale, expanded or
  condensed character spacing, raised/lowered position and a kerning threshold. It preserves
  untouched mixed values, validates edited values, and applies all dialog changes as one undo step.
  Combined Font/Advanced changes are atomic; header/footer formatting uses the body history so an
  undo reverses formatting without also undoing prior header typing.
  Superscript/Subscript toggles reflect inherited formatting and write `<w:vertAlign w:val="baseline"/>`
  when turning it off, including at a collapsed caret. Clearing script in the Font dialog handles
  inherited and mixed selections with one undo step; baseline renders at full size in both views.
  Word COM confirmed an exported reset is 12 pt with both script flags off while Normal retains
  its superscript setting. The source style XML remains unchanged.
  Tabs support arrow keys, Home/End and all five display locales. Ligatures now offers all 16
  Office 2010 combinations, preserves inherited and untouched mixed settings, and writes explicit
  None overrides. Editing surfaces, story previews, the dialog preview and Print Layout apply
  the selected OpenType features when the font supports them. Print Layout measures explicit
  selections with browser text shaping; exact Word glyph metrics remain approximate. Word COM
  confirmed Standard + Contextual as 3 and the exported None reset as 0. Text effects remain missing.
  The core now models horizontal scale (0–600%), kerning threshold and signed baseline position,
  including inherited style values and explicit neutral overrides. These survive editor typing,
  undo/redo and export in all six bindings. Invalid/unmodeled properties remain protected; Word COM
  verified an edited 125% / 12 pt kerning / lowered 3 pt run and a neutral reset. Print Layout now
  applies scale, additional character spacing, baseline position and kerning thresholds, including
  inherited settings. Scaled widths also drive wrapping and decimal tabs; Word COM confirmed that
  character spacing is added after scaling. Browser font metrics remain approximate. The continuous
  editor and header/footer/note previews now render inherited and direct spacing, baseline position
  and kerning thresholds, with script sizing retained. Paragraphs without an explicit pStyle now
  resolve the default Normal font settings, checked against Word COM. Horizontal scale now also
  renders on the continuous surface and in header/footer/note previews, with scaled glyph advances,
  spacing added after scale, word wrapping and native caret/selection. Unicode grapheme clusters
  stay intact. A scaled word split across formatting runs can still wrap at that split in a narrow
  column; the browser regression records this as an expected failure pending logical-word grouping.
  Scaled word boxes also suppress browser automatic hyphenation inside those words; Print Layout
  remains the required reference for pagination and keeps its existing hyphenation limitation.
  The dialog directs
  users to Print Layout to review these settings. Word COM verified a UI-exported run with 125%
  scale, 2 pt expanded spacing, lowered 3 pt position and 12 pt kerning.

The Home > Multilevel List menu now includes Define New Multilevel List. It edits nine levels,
with number style, `%1`–`%9` format tokens, start value, restart rule, legal numbering, alignment,
number/text positions and tab/space/nothing suffix. It starts from the selected list's supported
properties and creates a new independent definition; imported definitions remain untouched.
The dialog has a live marker preview, validation, Cancel/Escape, editing-permission checks and
all five display locales. Applying the list is one undoable paragraph edit; unused definitions
may remain in the additive catalog after undo. All six bindings verify apply/undo/redo/export
and preservation of the imported abstract definition. Live editing markers also recalculate on
list commands and history changes, including in header/footer/note editors. Word COM verified
the exported markers, start 4, never restart, right alignment, 0.25-inch number position,
1-inch text position and space suffix. The dialog fits a 1280×720 viewport with its action
buttons visible. For left-to-right paragraphs, the editor and Print Layout now place left,
center and right markers around the defined number position, with tab/space/nothing separators
and continuation lines at the text indent. Print Layout keeps a marker such as "Part 12."
indivisible and does not repeat it on a continuation page. Editor markers are generated CSS,
so they never become editable characters. Browser contracts compare editor and Print Layout
positions and verify typing at the paragraph start in all six bindings. The editor refreshes
its marker font measurements when web fonts load. Right-to-left marker placement and precise
editor positioning at center/right/decimal custom tabs remain approximate; The dialog's
Number font group sets a level's marker font family, size, bold, italic and colour (`w:lvl/w:rPr`,
written in schema order for new definitions, parsed from imports, and applied to the editor and Print
Layout markers over the paragraph's text formatting). Other `rPr` properties (underline, effects,
theme colours, East Asian fonts) and style-inherited marker fonts are not modeled or editable.

Imported heading outlines now resolve `w:lvl/w:pStyle` associations, rather than treating every
style-inherited item as level one. Direct paragraph numbering and explicit `numId=0` removal
retain precedence; effective level overrides are respected. Text editing preserves the original
styles and numbering XML. All six browser bindings verify nested markers, editing and opening
Define New Multilevel List at the inherited level with its imported pattern. Additive definition
export writes style links in schema order. Word COM confirmed the five-item sequence
`1., 1.1., 1.2., 2., 2.1.` when styles also specify matching `ilvl`. Installed Word discards the
second heading's numbering when that redundant style level is omitted; the viewer follows the
OOXML association rule for those imports. Each level has a Link level to style
select (a style belongs to one level). Applying the list writes `w:numPr` into those existing
paragraph styles in `styles.xml` (schema-ordered, the rest of the style XML untouched) and a
`w:pStyle` on the new level; any other style-catalog edit is still refused on save. The style link
is not part of editor history: Undo removes the list from the selection, but styles stay linked to
the (retained) definition. Creating new styles from the dialog is not supported.

Still not at parity:

- **Home:** Text Effects; Multilevel List offers five styles (1. 1.1. 1.1.1., 1. a) i., and heading-linked 1 / 1.1,
  Article I. / Section 1.1 and I. / A. outlines that link the document's existing Heading styles), not Word's full gallery (no bullet outlines or previews; "Section 1.01" shows as 1.1). Define New Multilevel List still needs Word's graphical number placeholders and
  its apply-to/list-style options. (Add tab stop at: a level's `w:tab w:val="num"` is read, written and
  honoured by the marker's tab suffix in the editor and Print Layout.) Text Highlight keeps the 17-colour list (Word's is a short
  list too). Line Spacing offers presets; the Paragraph dialog has the rest.
- **Insert:** Cover Page galleries (one plain design exists), Shapes, Icons, 3D Models, SmartArt,
  Charts, Screenshot, inserting Text Boxes, WordArt, Equation, Signature Line, Object.
- **Layout and References:** Manual Hyphenation and its options, Watermark, Page Borders, Position
  and Wrap Text, Citations, Bibliography, Index and Table of Authorities, and
  automatic updating of `REF`, `PAGEREF` and `SEQ` results (Update Fields refreshes them on request;
  only captions renumber on insert).
- **Review and View:** Editor pane, Thesaurus, Translate, Accessibility, Compare, Protect, Ink;
  Read Mode, Web and Outline views, New Window,
  Split and Macros.
- **Dropdown galleries:** Margins, Size, Orientation, Vertical alignment, Columns, Page Number, Borders and Line
  spacing are galleries with thumbnails (Line spacing ends with Line Spacing Options, which opens the
  Paragraph dialog; Borders ends with Borders and Shading).
  Page number format, Sort, Change case and similar stay plain text lists, as Word's Change Case
  does. Zoom is a percentage list next to a Zoom dialog. Margins ends with Custom Margins, which opens Page Setup. Columns includes Left and Right presets and ends with More Columns. Individual widths and gaps are editable; changing a width adjusts the adjacent column to retain the available text width. Print Layout uses unequal widths and reflows paragraph continuations. The continuous editing surface shows unequal-column sections at full text width, so Print Layout is required to review their columns.
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
- Picture bullets and `numStyleLink`. Multilevel counters now respect `lvlRestart=0` (never),
  one-based higher-level restart triggers, the omitted previous-level default, skipped ancestor
  levels and per-instance level/start overrides. Nested items initialize omitted ancestors,
  so a subsequent explicit ancestor advances its counter. Word COM comparisons cover default,
  never and restart-after-Level-1 sequences, including a skipped level whose intermediate
  ancestor never restarts; all six browser bindings display the same markers. Text editing and
  export preserve the original numbering XML. Out-of-range triggers use the OOXML default;
  Word's import normalization of redundant/invalid explicit triggers is not replicated (this
  installed Word discarded a level with an explicit previous-level trigger, rather than treating
  it like an omitted trigger). Remaining Define New Multilevel List controls are listed above.
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

- Modules over the 300-line guideline (non-test, non-data): `web-component/src/ribbon-tabs.ts` (~490, the tab definitions), `ribbon-home.ts`, `backstage-pages.ts`, `model-adapter.ts` and a few files just over 300; the locale tables and generated schema types are data.
- `ssr.test.ts` can time out (5 s) when the whole suite runs under load; it passes alone.
