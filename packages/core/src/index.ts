// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export * from './model.js';
export * from './highlight.js';
export { LIGATURE_VALUES, isLigatures, type Ligatures } from './ligatures.js';
export * from './language.js';
export {
	parseParagraphStyleCatalog,
	resolveParagraphFormatting,
	resolveStyleNumbering,
} from './paragraph-styles.js';
export * from './underline.js';
export { parseRunStyleCatalog } from './character-styles.js';
export { parseTableStyleCatalog } from './table-styles.js';
export { resolveTableStyleFormatting } from './resolve-table.js';
export { resolveRunFormatting, type RunFormattingContext } from './run-formatting.js';
export { parseTheme, parseColorSchemeMapping } from './theme.js';
export {
	resolveThemeColorToken,
	resolveThemeColorReference,
	applyThemeTint,
	applyThemeShade,
	isThemeColorToken,
} from './theme-color.js';
export { loadDocx } from './parse.js';
export { saveDocx } from './save.js';
export {
	assertValidDocumentModel,
	DocxModelValidationError,
	validateDocumentModel,
	type ValidationIssue,
} from './validate-model.js';
export * from './numbering-model.js';
export { parseNumberingCatalog, resolveNumberingLevel } from './numbering-parse.js';
export {
	computeListLabels,
	displayListLabel,
	formatListNumber,
	resolveParagraphNumbering,
} from './numbering-format.js';
export {
	headingListLevels,
	isHeadingListKind,
	type HeadingListKind,
} from './heading-list-kinds.js';
export {
	ensureListDefinition,
	createListDefinition,
	linkStylesToList,
	type ListKind,
} from './numbering-editing.js';
export { formatNoteNumber, numberNotesInOrder } from './notes.js';
export {
	acceptAllRevisions,
	acceptRevision,
	findRevision,
	listRevisions,
	rejectAllRevisions,
	rejectRevision,
	type RevisionEntry,
} from './revision-commands.js';
export { DEFAULT_TABLE_BORDERS } from './table-defaults.js';
export { fieldName } from './field-runs.js';
export { dateFieldResult, datePicture, formatWordDate } from './field-date.js';
export * from './table-of-contents.js';
export { resolveCellVisuals, type CellPlacement } from './table-visuals.js';
export { DEFAULT_STYLES_XML, defaultStyleCatalogs } from './default-styles.js';
export * from './generated/wml-simple-types.js';
export * from './units.js';
export * from './simple-types.js';
export { alignFromJustification, parseJustification } from './paragraph-alignment.js';
export type { ParagraphAlign, ParsedJustification } from './paragraph-alignment.js';
export { withParseWarnings, enumValue, reportParseWarning } from './parse-diagnostics.js';
export { PROPERTY_FIELDS, type DocumentProperties } from './core-properties.js';
export {
	captionParagraphs,
	isCaptionOf,
	tableOfFiguresInstruction,
	tocCaptionLabel,
} from './table-of-figures.js';
