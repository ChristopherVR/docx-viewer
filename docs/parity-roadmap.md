# Word parity milestones

Parity is measured separately for import, visual layout, editing, export, preservation and accessibility. Reading a field or retaining its XML does not prove it renders or recalculates correctly.

## 1. Verified editing foundation (implemented subset)

Shared editor, five lifecycle adapters, DOCX paragraphs/direct formatting/alignment/simple tables, a restricted plain-paragraph DOC text path, no-op preservation, original-format save, shared CFB and legacy Word binary code, unit and browser contracts. The editor currently uses a continuous editing surface and does not reproduce Word pagination. This establishes the architecture; it is not a Word replacement release.

Direct formatting includes bold, italic, underline, strikethrough, highlight,
superscript/subscript, font family/size/color and paragraph spacing/indentation.
The shared ribbon offers line-spacing multiples from single to triple, including
1.15 and 1.5, and shows imported exact/minimum spacing in points. Choosing inherited
spacing removes the direct line-spacing override; style inheritance is not yet
resolved on the editing surface. Shift+Enter inserts a line break inside the
current paragraph, while Enter creates another paragraph. Both persist through
DOCX save and reload in every framework binding.
Highlight values follow the [WordprocessingML color enumeration](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.highlightcolorvalues?view=openxml-3.0.1).
Simple rectangular tables support row/column insertion and deletion, whole-table
deletion, and independent undo/redo. Existing cell XML follows retained paragraph
identities when rows or columns move. Merged, nested and complex imported tables
disable structural commands; export independently checks the source XML. Widths,
borders and other unmodeled table styling are retained where supported, but are
not rendered with Word layout fidelity or exposed as editing controls.

## 2. Fidelity and document model

Resolve document defaults, paragraph/character styles, theme fonts/colors, numbering and tabs. Model sections, breaks, headers/footers, images, relationships, hyperlinks, notes and merged tables. Use ordered OOXML preservation with explicit unsupported-edit errors. Introduce Word-authored fixtures with corresponding expected text and package checks before extending each feature.

Add shared OOXML capabilities only when their contracts are established for both consumers. Keep the canonical DOCX model/parser/writer in `docx-core`; `ole2` remains the owner of CFB and binary Word code.

Paragraph styles require a style catalog and separate direct and resolved
formatting. Resolve document defaults and `basedOn` chains in the core (with
cycle detection), then expose the same result to the single editor. Preserve
explicit off values and inherited properties when editing; do not flatten the
resolved appearance into every run. Style creation must update `styles.xml`, its
relationship and content type before a style picker can safely assign new IDs.
Character styles, theme references and Word toggle-property semantics need
separate fixtures before claiming full inheritance support.

## 3. Pagination and WYSIWYG layout

Separate semantic editing from page layout. Build line breaking, font metrics/substitution, widow/orphan control, keep-with-next, page/section breaks, repeating headers, table row fragmentation, footnote placement and floating object wrapping. Define a stable page/line mapping for selection and hit testing. Browser CSS min-height and page size are insufficient evidence of Word parity.

Validate against a checked-in corpus rendered by Word (and cross-check LibreOffice), with reference images, text bounding boxes and tolerances. Include variable fonts, complex scripts, RTL, vertical writing, equations, large documents and printer/PDF pagination. Keep selection/IME/accessibility tests alongside visual comparisons.

## 4. Editing breadth

Add style/numbering controls, tables and images, find/replace, links, comments, tracked changes, fields, review tools, clipboard fidelity, printing, collaboration and accessibility. Implement every command once in the editor controller; run the same browser contract in each framework adapter. Keep DOC output explicitly constrained until corresponding binary structures are supported and validated.

## 5. Release gates

No unsupported feature silently disappears. Every feature has import/edit/export/layout coverage with real fixtures. Cross-framework and cross-browser tests pass. Package artifacts work in independent consumer apps. CI checks out/builds the shared repository at a pinned revision (or consumes a published release); source ownership checks reject copied OLE implementations. Publish only after these gates are met.
