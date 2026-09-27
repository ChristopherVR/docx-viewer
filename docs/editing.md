# Editing text

Every framework mounts the same web-component editor. Commands, keyboard
shortcuts, search, language metadata and document conversion have one implementation.

## Find and replace

Use **Find and replace** or **Ctrl+F** (**Command+F** on macOS) while the editor is
focused. Enter a literal search phrase and use **Find next** or **Find previous**
to move through matches. **Match case** controls case sensitivity. Escape closes
the panel and returns focus to the document.

Search spans adjacent formatted runs within a paragraph. It does not match across
paragraphs, table cells or hard line breaks. It avoids matches that split a Unicode
grapheme, such as part of a combining sequence or joined emoji. Searches are literal;
regular expressions and locale-specific linguistic matching are not supported.

**Replace** changes the current match; **Replace all** changes the document's
matches in one undoable operation. Replacement text takes the first matched
character's formatting. Text outside the match retains its formatting. Searching
works in read-only mode; replacement is disabled. Existing DOCX preservation checks
still apply when exporting edits around unsupported source content.

## Multilingual documents

Paragraph direction and text language are separate properties. Use **Paragraph
direction** for left-to-right or right-to-left layout, and **Run direction** for an
explicit direction override within a paragraph. The language controls store Word's
regular, East Asian and complex-script language tags. Language metadata does not
translate text or install a spellchecker.

The editor preserves Unicode text, including Arabic, Hebrew, CJK, combining
characters and emoji. Browser fonts, shaping and bidirectional layout determine
the on-screen appearance. These controls do not establish Word-equivalent complex
script typography, pagination or operating-system IME behavior.

Word counts use the browser's Unicode word segmentation, including dictionary
segmentation for unspaced scripts where supported. Counts can differ between ICU
versions and from Microsoft Word. Older browsers use a less precise fallback.

## Coauthoring integration

The [collaboration guide](/collaboration) describes the versioned editing protocol
and how an application connects the shared editor to its authority. This is an
integration API, not a hosted collaboration service. Access control, networking,
storage and reconnect policy belong to the application.

## Interface language

Set the shared editor's `locale` property (or the binding's `locale` prop) to `fr`
for French or `en` for English. Regional tags such as `fr-CA` use the French UI;
unsupported locales fall back to English. Changing the interface language preserves
the current document, selection, undo history and document-language metadata.

## Imported paragraph styles

The Home ribbon's Style picker lists paragraph styles from the loaded DOCX.
The editor resolves supported alignment, direction, spacing and indentation through
document defaults and the selected style's inheritance chain. Direct paragraph
formatting still takes precedence. Choosing a style changes its reference, without
flattening inherited properties into every paragraph.

The original styles part is preserved when saving the loaded document. Creating or
editing style definitions, character styles, inherited run formatting and theme
resolution are not supported yet. Exporting a new package with a style catalog is
rejected; use Save on the loaded document to preserve its original style definitions.
