// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export * from './model.js';
export * from './highlight.js';
export * from './language.js';
export {
	parseParagraphStyleCatalog,
	resolveParagraphFormatting,
	resolveStyleNumbering,
} from './paragraph-styles.js';
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
