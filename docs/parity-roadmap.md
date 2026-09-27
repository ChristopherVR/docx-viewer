# Word parity milestones

Parity is measured separately for import, visual layout, editing, export, preservation and accessibility. Reading a field or retaining its XML does not prove it renders or recalculates correctly.

## 1. Verified editing foundation (implemented subset)

Shared editor, five lifecycle adapters, DOCX paragraphs/direct formatting/alignment/simple tables, a restricted plain-paragraph DOC text path, no-op preservation, original-format save, shared CFB and legacy Word binary code, unit and browser contracts. The editor currently uses a continuous editing surface and does not reproduce Word pagination. This establishes the architecture; it is not a Word replacement release.

Direct formatting includes bold, italic, underline, strikethrough, highlight,
superscript/subscript, font family/size/color and paragraph spacing/indentation.
The shared ribbon offers line-spacing multiples from single to triple, including
1.15 and 1.5, and shows imported exact/minimum spacing in points. Choosing inherited
spacing removes the direct line-spacing override. Supported paragraph style properties
resolve through document defaults and basedOn inheritance without flattening exports. Shift+Enter inserts a line break inside the
current paragraph, while Enter creates another paragraph. Both persist through
DOCX save and reload in every framework binding.
Highlight values follow the [WordprocessingML color enumeration](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.highlightcolorvalues?view=openxml-3.0.1).
Simple rectangular tables support row/column insertion and deletion, whole-table
deletion, and independent undo/redo. Existing cell XML follows retained paragraph
identities when rows or columns move. Merged, nested and complex imported tables
disable structural commands; export independently checks the source XML. Widths,
borders and other unmodeled table styling are retained where supported, but are
not rendered with Word layout fidelity or exposed as editing controls.

The shared editor also supports literal Unicode-aware find/replace, paragraph
LTR/RTL direction, three Word language metadata fields, run direction overrides,
and Unicode word counting. See [editing text](/editing) for the behavior and limits.
Language metadata does not provide translation, spellchecking, UI localization or
proof of full complex-script/IME parity.

The [collaboration protocol](/collaboration) synchronizes versioned ProseMirror
steps through an authority, rebases supported concurrent edits and preserves
pending edits while waiting for acknowledgements. The [local coauthoring demo](/demo/collaboration.html)
provides two editors and paused delivery for testing concurrency. Production
networking, permissions, persistence and shared-save coordination
remain application responsibilities. Transient peer cursors/selections and English/French
interface localization are available in the shared editor. Structural table commands are disabled during
collaboration because their whole-table replacements are not yet conflict-aware;
editing text inside tables remains available.

## 2. Fidelity and document model

Resolve document defaults, paragraph/character styles, theme fonts/colors, numbering and tabs. Model sections, breaks, headers/footers, images, relationships, hyperlinks, notes and merged tables. Use ordered OOXML preservation with explicit unsupported-edit errors. Introduce Word-authored fixtures with corresponding expected text and package checks before extending each feature.

Add shared OOXML capabilities only when their contracts are established for both consumers. Keep the canonical DOCX model/parser/writer in `docx-core`; `ole2` remains the owner of CFB and binary Word code.

Paragraph styles now have a read-only catalog with separate direct and resolved
paragraph formatting. Document defaults and `basedOn` chains resolve in the core
with cycle detection and are rendered by the shared editor. The style picker can
select existing definitions; inherited paragraph values remain out of direct formatting. Style creation must update `styles.xml`, its
relationship and content type before a style picker can safely assign new IDs.
Character styles, theme references and Word toggle-property semantics need
separate fixtures before claiming full inheritance support.

List numbering is now modeled. `word/numbering.xml` (`abstractNum`, `num`, `lvlOverride`,
`startOverride`) is parsed into a read-only catalog, and a paragraph's direct `w:numPr`
(`numId`/`ilvl`) or numbering inherited through its paragraph style (`pStyle` → style `numPr`)
resolves to a computed marker: decimal, upper/lower Roman, upper/lower letter, ordinal,
decimal-zero, ordinal/cardinal text, bullet and legal (`isLgl`) numbering, including multilevel
`lvlText` placeholders (`%1`..`%9`) and per-level restart across the document. Markers render as
non-editable generated content with the level's indentation as a fallback when the paragraph has
no direct indent. The shared editor adds bulleted list, numbered list, increase/decrease list
level and remove list commands; applying a list to a plain paragraph mints a fresh, independent
`numId` (and its `abstractNum`) so unrelated lists never share counters. Enter on an empty list
item ends the list, and Tab/Shift+Tab change level inside one. Saving appends new numbering
definitions to `word/numbering.xml` (creating the part, its relationship and content-type override
the first time a document gains a list) while existing `abstractNum`/`num` entries are read-only;
editing one throws rather than silently rewriting it, and a no-op save still returns the original
bytes. Picture bullets, style-linked numbering (`numStyleLink`/`numPicBulletId`), other
custom/legacy `numFmt` tokens (they render as Decimal Number), and Word's exact per-level
`lvlRestart` cascade (a simplified "any shallower level restarts every deeper level" rule is used
instead) remain unsupported or approximated; see the numbering model warnings emitted at parse
time for the specifics of a given document.
Sections are now modeled: every paragraph-level `sectPr` (section break) and the
final body section resolve into `DocumentModel.sections`, each with page size,
orientation, margins, header/footer/gutter distances, column layout (count,
spacing, equal/individual widths, separator), section type, `titlePg`, vertical
alignment, page-number start/format, and line-numbering/page-border presence.
The pagination workstream is the primary consumer of this shape; the shared
editor still renders one continuous surface and does not lay out columns or
per-page headers/footers itself.

Page and column breaks (`w:br` type page/column) are distinguished from
ordinary line breaks and `pageBreakBefore` is modeled on the paragraph; both
round-trip and are editable, with a visible break marker and an Insert > Page
break command (Ctrl+Enter). A break mixed into a run alongside other text (not
how Word authors documents) remains protected rather than silently
demoted to a line break. Other non-line breaks (e.g. `w:clear`) are still
unsupported and protect their paragraph from edits.

Headers and footers (default/first/even, honoring `settings.xml`
`evenAndOddHeaders`) resolve via relationships into read-only paragraph/table
content and render above/below the continuous surface; they cannot be edited
and their parts are always byte-preserved. Simple `PAGE`/`NUMPAGES`/etc. fields
inside them show as static bracketed placeholders, not recalculated values.

Footnotes and endnotes parse from `footnotes.xml`/`endnotes.xml` (skipping
Word's separator marks) and render at the end of the surface, footnotes then
endnotes, numbered by first-reference order in the requested `numFmt`.
Footnote/endnote reference marks render as superscript numbers inline in the
body; paragraphs containing a reference mark remain protected from edits,
since relocating that mark safely is not supported.

## 3. Pagination and WYSIWYG layout

Separate semantic editing from page layout. Build line breaking, font metrics/substitution, widow/orphan control, keep-with-next, page/section breaks, repeating headers, table row fragmentation, footnote placement and floating object wrapping. Define a stable page/line mapping for selection and hit testing. Browser CSS min-height and page size are insufficient evidence of Word parity.

Validate against a checked-in corpus rendered by Word (and cross-check LibreOffice), with reference images, text bounding boxes and tolerances. Include variable fonts, complex scripts, RTL, vertical writing, equations, large documents and printer/PDF pagination. Keep selection/IME/accessibility tests alongside visual comparisons.

## 4. Editing breadth

Add style/numbering controls, tables and images, find/replace, links, comments, tracked changes, fields, review tools, clipboard fidelity, printing, collaboration and accessibility. Implement every command once in the editor controller; run the same browser contract in each framework adapter. Keep DOC output explicitly constrained until corresponding binary structures are supported and validated.

## 5. Release gates

No unsupported feature silently disappears. Every feature has import/edit/export/layout coverage with real fixtures. Cross-framework and cross-browser tests pass. Package artifacts work in independent consumer apps. CI checks out/builds the shared repository at a pinned revision (or consumes a published release); source ownership checks reject copied OLE implementations. Publish only after these gates are met.

Outstanding and in-progress workstreams are tracked in [outstanding work](/outstanding-work).
