// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export type {
	ThemeColorSlot,
	ThemeColorToken,
	ThemeFontSet,
	ThemeCatalog,
	ThemeColorReference,
	ThemeFontScript,
	ThemeFontRole,
	WordUnderlineStyle,
} from './theme-model.js';
import type { ThemeColorReference, ThemeFontScript, ThemeFontRole, WordUnderlineStyle } from './theme-model.js';
export type {
	RunFormatting,
	CharacterStyleDefinition,
	RunStyleCatalog,
} from './run-style-model.js';
import type { RunStyleCatalog } from './run-style-model.js';
export type {
	TableBorderSide,
	TableBorders,
	TableCellMargins,
	NestedTablePreview,
	TableLook,
	TableConditionalRegion,
	TableStyleConditionalFormatting,
	TableStyleDefinition,
	TableStyleCatalog,
} from './table-model.js';
import type { TableCell as TableCellShape, TableStyleCatalog } from './table-model.js';

export interface TextRun {
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	/** Word named highlight color, such as `yellow` or `lightGray`. */
	highlight?: string;
	verticalAlign?: 'superscript' | 'subscript';
	/** Direct Word run language tag (`w:lang/@w:val`), without automatic detection. */
	language?: string;
	/** Direct East Asian script language tag (`w:lang/@w:eastAsia`). */
	eastAsiaLanguage?: string;
	/** Direct complex-script language tag (`w:lang/@w:bidi`). */
	bidiLanguage?: string;
	/** Explicit run-level bidirectional override; `false` means direct off, undefined inherits. */
	rtl?: boolean;
	/** Font size in points. */
	fontSize?: number;
	fontFamily?: string;
	/** Direct RGB color, e.g. #28665E. May coexist with `colorTheme` as Word's stored fallback. */
	color?: string;
	/** Direct `w:color/@w:themeColor` (+ themeTint/themeShade); resolution happens in a separate layer. */
	colorTheme?: ThemeColorReference;
	/** Character style reference (`w:rStyle/@w:val`); preserved and editable, not flattened. */
	style?: string;
	caps?: boolean;
	smallCaps?: boolean;
	/** `w:dstrike`; kept distinct from the single-line `strike` toggle. */
	doubleStrike?: boolean;
	/** Hidden text (`w:vanish`); the editor renders it dimmed rather than removing it. */
	vanish?: boolean;
	/** Non-single underline style, e.g. `double`/`wave`; `underline` stays the simple on/off toggle. */
	underlineStyle?: WordUnderlineStyle;
	underlineColor?: string;
	/** `w:spacing/@w:val` character spacing, in twips (positive expands, negative condenses). */
	characterSpacingTwips?: number;
	/** Direct `w:shd/@w:fill` run shading. */
	shadingFill?: string;
	/** Direct `w:shd` theme fill; kept alongside `shadingFill` without flattening. */
	shadingThemeFill?: ThemeColorReference;
	/** Direct `w:rFonts` theme font references, per script; resolved via the document theme. */
	fontTheme?: Partial<Record<ThemeFontScript, ThemeFontRole>>;
}
export interface Paragraph {
	type: 'paragraph';
	/** Stable identity across edits, unique throughout the document. */
	id: string;
	runs: TextRun[];
	align?: 'left' | 'center' | 'right' | 'justify';
	/** Paragraph base direction from direct `w:bidi`; undefined inherits. */
	direction?: 'ltr' | 'rtl';
	style?: string;
	/** Word paragraph spacing and indentation values, kept in their native twip units (1/20 pt). */
	spacingBeforeTwips?: number;
	spacingAfterTwips?: number;
	lineSpacingTwips?: number;
	/** `auto` uses 240ths of a line; `exact` and `atLeast` use twips. */
	lineSpacingRule?: 'auto' | 'exact' | 'atLeast';
	indentLeftTwips?: number;
	indentRightTwips?: number;
	indentStartTwips?: number;
	indentEndTwips?: number;
	firstLineTwips?: number;
	hangingTwips?: number;
}
/** Direct paragraph properties supported by the editor, in native Word units. */
export type ParagraphFormatting = Pick<
	Paragraph,
	| 'align'
	| 'direction'
	| 'spacingBeforeTwips'
	| 'spacingAfterTwips'
	| 'lineSpacingTwips'
	| 'lineSpacingRule'
	| 'indentLeftTwips'
	| 'indentRightTwips'
	| 'indentStartTwips'
	| 'indentEndTwips'
	| 'firstLineTwips'
	| 'hangingTwips'
>;
export interface ParagraphStyleDefinition {
	id: string;
	name?: string;
	basedOn?: string;
	isDefault?: boolean;
	formatting: ParagraphFormatting;
}
/** Parsed source style data. It is retained for inheritance and never flattened on save. */
export interface ParagraphStyleCatalog {
	docDefaults: ParagraphFormatting;
	styles: Record<string, ParagraphStyleDefinition>;
	warnings: string[];
}
export type TableCell = TableCellShape;
export interface Table {
	type: 'table';
	id: string;
	/** Whether row, cell and paragraph structure can be edited without losing unsupported table XML. */
	structureEditable?: boolean;
	rows: TableCell[][];
	/** `w:tblGrid/w:gridCol` widths in twips, one per grid column. */
	grid?: number[];
	widthTwips?: number;
	alignment?: 'left' | 'center' | 'right';
	indentTwips?: number;
	borders?: import('./table-model.js').TableBorders;
	/** `w:tblStyle/@w:val`; conditional formatting resolves through `tableStyles` without flattening. */
	style?: string;
	look?: import('./table-model.js').TableLook;
}
export type Block = Paragraph | Table;
export interface DocumentModel {
	blocks: Block[];
	/** CSS pixels at 96 DPI. Page height is currently a surface minimum, not pagination. */
	page: {
		width: number;
		height: number;
		marginTop: number;
		marginRight: number;
		marginBottom: number;
		marginLeft: number;
	};
	warnings: string[];
	/** Source paragraph defaults/styles; editing the catalog itself is not supported. */
	paragraphStyles?: ParagraphStyleCatalog;
	/** Source run-level defaults/character styles; editing the catalog itself is not supported. */
	characterStyles?: RunStyleCatalog;
	/** Source table style catalog; editing the catalog itself is not supported. */
	tableStyles?: TableStyleCatalog;
	/** Parsed word/theme/theme1.xml and settings.xml color scheme mapping. */
	theme?: import('./theme-model.js').ThemeCatalog;
}
export interface LoadedDocument {
	model: DocumentModel;
	/** Returns bytes in the original format; unsupported edits reject instead of silently degrading it. */
	save(model?: DocumentModel): Promise<Uint8Array>;
}
export function createDocument(): DocumentModel {
	return {
		blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text: '' }] }],
		page: {
			width: 816,
			height: 1056,
			marginTop: 96,
			marginRight: 96,
			marginBottom: 96,
			marginLeft: 96,
		},
		warnings: [],
	};
}
