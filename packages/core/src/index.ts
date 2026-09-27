// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export * from './model.js';
export * from './highlight.js';
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
export * from './numbering-model.js';
export { parseNumberingCatalog, resolveNumberingLevel } from './numbering-parse.js';
export {
	computeListLabels,
	formatListNumber,
	resolveParagraphNumbering,
} from './numbering-format.js';
export { ensureListDefinition, type ListKind } from './numbering-editing.js';
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
