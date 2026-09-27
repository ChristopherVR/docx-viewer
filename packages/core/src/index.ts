// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export * from './model.js';
export * from './highlight.js';
export * from './language.js';
export * from './underline.js';
export { parseParagraphStyleCatalog, resolveParagraphFormatting } from './paragraph-styles.js';
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
